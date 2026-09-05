//! A generic COM event sink.
//!
//! This module knows nothing about Excel. Any IDispatch that names a default
//! source dispinterface can be advised on, which is what makes Application,
//! Workbook, a worksheet ActiveX control and a UserForm control all the same
//! case rather than four.

use std::collections::HashMap;
use std::ffi::c_void;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::mpsc::Sender;
use windows::core::{IUnknown, IUnknown_Vtbl, Interface, BSTR, GUID, HRESULT, PCWSTR, VARIANT};
use windows::Win32::Foundation::{E_NOINTERFACE, E_NOTIMPL, E_POINTER, RPC_E_WRONG_THREAD, S_OK};
use windows::Win32::System::Com::{
    CoGetApartmentType, IAdviseSink, IAdviseSink2, IConnectionPoint, IConnectionPointContainer,
    IDispatch, IDispatch_Vtbl, ITypeInfo, ITypeLib, APTTYPE, APTTYPEQUALIFIER, APTTYPE_MAINSTA,
    APTTYPE_STA, DISPATCH_FLAGS, DISPPARAMS, EXCEPINFO, FUNCFLAG_FRESTRICTED, IMPLTYPEFLAG_FDEFAULT,
    IMPLTYPEFLAG_FSOURCE, TKIND_DISPATCH, TYPEKIND,
};
use windows::Win32::System::Ole::{
    IAdviseSinkEx, IPropertyNotifySink, IProvideClassInfo, IProvideClassInfo2,
    GUIDKIND_DEFAULT_SOURCE_DISP_IID,
};
use crate::dispatch::{ComError, ComResult, LOCALE_USER_DEFAULT};

/// One callback, as it came off the wire from Excel. VARIANTs, not JSON:
/// converting them needs the handle table, which lives in the session.
pub struct RawEvent {
    pub handle: u64,
    pub dispid: i32,
    pub args: Vec<VARIANT>,
}

/// The sink object Excel calls back into: a COM object whose memory layout is
/// the ABI, laid out by hand rather than by `#[implement(IDispatch)]`.
///
/// The macro was the obvious choice and does not work here. Its generated
/// `QueryInterface` starts with `let iid = &*iid;`, and the IID an
/// out-of-process caller passes is not guaranteed to be aligned: marshalling
/// this sink over to Excel goes through `NdrInterfacePointerMarshall`, which
/// hands `QueryInterface` a pointer straight into RPC's packed NDR format
/// string. Measured under Wine, that pointer landed on an odd address, and
/// forming a `&GUID` from it is a misaligned reference -- undefined behaviour,
/// which a debug build turns into `panic_misaligned_pointer_dereference`.
/// Inside an `extern "system"` function that panic cannot unwind, so it aborts
/// the process: not one failing test but the whole test binary, every time the
/// sink is advised. `read_unaligned` in `query_interface` below is the entire
/// fix, and owning the vtable is what makes it reachable.
///
/// `#[repr(C)]` with `vtable` first is load-bearing: a `*mut EventSink` is
/// handed to COM directly as the interface pointer.
///
/// Not `Send`/`Sync`, and deliberately not marked so. The object is created
/// on, advised from, and called back on one STA thread; COM's apartment rules,
/// not Rust's, are what keep that true -- and [`require_sta`] is what keeps
/// COM's side of that bargain from being assumed rather than checked.
///
/// Private, and constructed only by [`advise`]: nothing outside this module
/// may hand this object to COM, because nothing outside this module can
/// establish the apartment invariant that makes it sound.
#[repr(C)]
struct EventSink {
    vtable: &'static IDispatch_Vtbl,
    refcount: AtomicU32,
    /// The source dispinterface this sink was advised for.
    ///
    /// A connection point is entitled to ask for that exact IID before it
    /// accepts a sink -- ATL's `IConnectionPointImpl::Advise` does precisely
    /// `pUnk->QueryInterface(m_iid, ...)` -- and answering it with our
    /// `IDispatch` is what being a dispinterface sink means. Excel's own
    /// connection point settles for `IDispatch`, but a worksheet ActiveX
    /// control's will not, and this module claims to serve both.
    source_iid: GUID,
    tx: Sender<RawEvent>,
    handle: u64,
}

static SINK_VTABLE: IDispatch_Vtbl = IDispatch_Vtbl {
    base__: IUnknown_Vtbl {
        QueryInterface: EventSink::query_interface,
        AddRef: EventSink::add_ref,
        Release: EventSink::release,
    },
    GetTypeInfoCount: EventSink::get_type_info_count,
    GetTypeInfo: EventSink::get_type_info,
    GetIDsOfNames: EventSink::get_ids_of_names,
    Invoke: EventSink::invoke,
};

impl EventSink {
    /// A new sink, owned by the returned `IUnknown`: it holds the one
    /// reference, and its `Drop` releases it.
    fn new(tx: Sender<RawEvent>, handle: u64, source_iid: GUID) -> IUnknown {
        let boxed = Box::new(EventSink {
            vtable: &SINK_VTABLE,
            refcount: AtomicU32::new(1),
            source_iid,
            tx,
            handle,
        });
        unsafe { IUnknown::from_raw(Box::into_raw(boxed) as *mut c_void) }
    }

    unsafe extern "system" fn query_interface(
        this: *mut c_void,
        iid: *const GUID,
        interface: *mut *mut c_void,
    ) -> HRESULT {
        if iid.is_null() || interface.is_null() {
            return E_POINTER;
        }
        // Read, never borrow: see the type's doc comment for why this pointer
        // may be unaligned and what borrowing it costs.
        let iid = unsafe { std::ptr::read_unaligned(iid) };
        let obj = unsafe { &*(this as *const EventSink) };
        if iid == IUnknown::IID || iid == IDispatch::IID || iid == obj.source_iid {
            obj.refcount.fetch_add(1, Ordering::Relaxed);
            unsafe { *interface = this };
            S_OK
        } else {
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }
    }

    unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
        let obj = unsafe { &*(this as *const EventSink) };
        // `wrapping_add`, not `+ 1`: an overflow panic here could not unwind
        // out of an `extern "system"` function and would abort the process
        // instead. Every arithmetic step in this vtable follows that rule.
        obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
    }

    unsafe extern "system" fn release(this: *mut c_void) -> u32 {
        let ptr = this as *mut EventSink;
        let remaining = unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
        if remaining == 0 {
            drop(unsafe { Box::from_raw(ptr) });
        }
        remaining
    }

    /// No typeinfo. A dispinterface sink is called by DISPID, and the names
    /// this bridge reports come from the *source's* typeinfo (`dispid_names`),
    /// never from the sink's.
    unsafe extern "system" fn get_type_info_count(_this: *mut c_void, count: *mut u32) -> HRESULT {
        if count.is_null() {
            return E_POINTER;
        }
        unsafe { *count = 0 };
        S_OK
    }

    unsafe extern "system" fn get_type_info(
        _this: *mut c_void,
        _index: u32,
        _lcid: u32,
        type_info: *mut *mut c_void,
    ) -> HRESULT {
        if !type_info.is_null() {
            unsafe { *type_info = std::ptr::null_mut() };
        }
        E_NOTIMPL
    }

    unsafe extern "system" fn get_ids_of_names(
        _this: *mut c_void,
        _riid: *const GUID,
        _names: *const PCWSTR,
        _count: u32,
        _lcid: u32,
        _ids: *mut i32,
    ) -> HRESULT {
        E_NOTIMPL
    }

    /// NEVER BLOCK IN HERE. This runs inside a COM call from Excel; anything
    /// that waits stalls Excel itself. Copy the arguments, push, return.
    ///
    /// The argument ORDER, which is the whole of the subtlety, lives in
    /// [`ordered_args`].
    unsafe extern "system" fn invoke(
        this: *mut c_void,
        dispid: i32,
        _riid: *const GUID,
        _lcid: u32,
        _flags: DISPATCH_FLAGS,
        params: *const DISPPARAMS,
        _result: *mut std::mem::MaybeUninit<VARIANT>,
        _excep: *mut EXCEPINFO,
        _arg_err: *mut u32,
    ) -> HRESULT {
        let obj = unsafe { &*(this as *const EventSink) };
        let args = if params.is_null() { Vec::new() } else { ordered_args(unsafe { &*params }) };
        // A closed receiver means the session is gone; there is nothing to
        // report it to and failing back into Excel helps nobody.
        //
        // This send deliberately does NOT wake the session thread, which every
        // other send in this project has to (see `session::SessionRoute::send`).
        // It does not need to: `require_sta` guarantees this `Invoke` is
        // running ON the session thread, inside its own `drain_messages()`
        // call, so the thread is already awake and drains this channel on its
        // way back round the loop. A wake here would be a wake the thread
        // sends itself.
        let _ = obj.tx.send(RawEvent { handle: obj.handle, dispid, args });
        S_OK
    }
}

/// One callback's arguments, in DECLARATION order.
///
/// This is the exact inverse of what `dispatch::raw_invoke` builds on the way
/// out, and the two are meant to be read against each other rather than as
/// this project holding two opposite beliefs about argument order:
///
/// * `rgvarg[0 .. cNamedArgs]` are NAMED, each identified by the DISPID at the
///   matching index of `rgdispidNamedArgs`. A parameter's DISPID *is* its
///   zero-based position -- the same convention `dispatch::get_dispids` relies
///   on when it sends named arguments out.
/// * `rgvarg[cNamedArgs .. cArgs]` are POSITIONAL, last parameter first.
///
/// **Measured, and not what either the brief or the obvious sink assumes:**
/// every Excel Application event arrives here FULLY NAMED. `SheetChange`
/// comes in with `cArgs=2, cNamedArgs=2, rgdispidNamedArgs=[0, 1]` -- "the
/// argument at position 0", then "the argument at position 1" -- as does
/// `WindowActivate(Wb, Wn)`, and the one-argument events arrive with `[0]`.
/// So the
/// reversal question does not even arise for them: the named half of the rule
/// above places both arguments by position, which is why `rgvarg` reads in
/// declaration order and why reversing it (the brief's `.rev()`) produced
/// `(Range, Worksheet)` instead of `SheetChange(Sh, Target)`.
///
/// That measurement was taken with the `source_iid` arm of the sink's
/// `QueryInterface` both present and absent, because the two arrived
/// together and answering a different set of IIDs could in principle change
/// which RPC stub Wine builds for the sink -- and the stub is what rebuilds
/// this DISPPARAMS. It does not: with the arm removed, `SheetChange` still
/// arrives as `cArgs=2, cNamedArgs=2, [0, 1]` and still delivers
/// `(Worksheet, Range)`. The two decisions are independent.
///
/// The positional half has never been observed inbound on this Wine. It
/// follows the convention that the OUTBOUND path is measured to obey -- Excel
/// reads `rgvarg` last parameter first, which is what
/// `dispatch::test_positional_arguments_reach_excel_in_order` pins -- and it
/// is what an in-process ATL control's `Fire_` method sends, which this
/// module claims to serve. Both halves are pinned by unit tests, so a source
/// that disagrees is one failing assertion away from being noticed rather
/// than a backwards callback for a caller to puzzle over.
///
/// Nothing here can panic (a panic in an `extern "system"` callback aborts the
/// process) and nothing is dropped: if the named DISPIDs are not usable
/// parameter positions -- `DISPID_PROPERTYPUT`, a real member DISPID, the
/// same position twice, a null id array -- the arguments are handed over
/// exactly as they arrived, because there is no honest way to reorder them
/// and a plausible-looking wrong order is the failure this module fears most.
///
/// Every VARIANT is cloned, which is a `VariantCopy`: each argument gets its
/// own reference and outlives this call. The originals belong to the caller
/// and are gone the moment we return.
fn ordered_args(p: &DISPPARAMS) -> Vec<VARIANT> {
    if p.rgvarg.is_null() {
        return Vec::new();
    }
    let count = p.cArgs as usize;
    let all = unsafe { std::slice::from_raw_parts(p.rgvarg, count) };
    let named_count = (p.cNamedArgs as usize).min(count);
    if named_count > 0 && p.rgdispidNamedArgs.is_null() {
        return all.to_vec();
    }
    let named_ids: &[i32] = if named_count == 0 {
        &[]
    } else {
        unsafe { std::slice::from_raw_parts(p.rgdispidNamedArgs, named_count) }
    };

    let mut slots: Vec<Option<VARIANT>> = (0..count).map(|_| None).collect();
    for (i, &id) in named_ids.iter().enumerate() {
        match usize::try_from(id) {
            Ok(position) if position < count && slots[position].is_none() => {
                slots[position] = Some(all[i].clone());
            }
            _ => return all.to_vec(),
        }
    }
    // The positional tail fills whatever the named arguments left empty --
    // the front of the list, in the ordinary case -- last parameter first.
    let free: Vec<usize> = (0..count).filter(|i| slots[*i].is_none()).collect();
    let positional = &all[named_count..];
    if free.len() != positional.len() {
        return all.to_vec();
    }
    for (&position, v) in free.iter().zip(positional.iter().rev()) {
        slots[position] = Some(v.clone());
    }
    let ordered: Vec<VARIANT> = slots.into_iter().flatten().collect();
    if ordered.len() == count {
        ordered
    } else {
        all.to_vec()
    }
}

/// The default source interface a coclass typeinfo declares, or `None` when
/// it declares none it can be sure of.
///
/// This is the evidence `IProvideClassInfo2::GetGUID` summarises: the coclass
/// lists the interfaces it implements, and the one flagged both `FSOURCE` and
/// `FDEFAULT` is its default source. A coclass with exactly one `FSOURCE`
/// entry and no `FDEFAULT` flag leaves no choice either. Two or more sources
/// and no default is the same "nothing says which" as several connection
/// points, and answers `None` so the caller falls through to its own refusal.
fn coclass_default_source(coclass: &ITypeInfo) -> Option<GUID> {
    let attr = unsafe { coclass.GetTypeAttr() }.ok()?;
    let impl_count = unsafe { (*attr).cImplTypes } as u32;
    unsafe { coclass.ReleaseTypeAttr(attr) };

    let mut default: Option<u32> = None;
    let mut sources: Vec<u32> = Vec::new();
    for index in 0..impl_count {
        let Ok(flags) = (unsafe { coclass.GetImplTypeFlags(index) }) else { continue };
        if flags.0 & IMPLTYPEFLAG_FSOURCE.0 == 0 {
            continue;
        }
        sources.push(index);
        if flags.0 & IMPLTYPEFLAG_FDEFAULT.0 != 0 {
            default = Some(index);
            break;
        }
    }
    let index = match (default, sources.as_slice()) {
        (Some(index), _) => index,
        (None, [only]) => *only,
        _ => return None,
    };

    let href = unsafe { coclass.GetRefTypeOfImplType(index) }.ok()?;
    let source = unsafe { coclass.GetRefTypeInfo(href) }.ok()?;
    let attr = unsafe { source.GetTypeAttr() }.ok()?;
    let guid = unsafe { (*attr).guid };
    unsafe { source.ReleaseTypeAttr(attr) };
    Some(guid)
}

/// The IID of `disp`'s default source dispinterface.
///
/// Three ways of being told, tried in order, and no guessing after them.
/// `IProvideClassInfo2` answers directly. Failing that, the coclass typeinfo
/// behind `IProvideClassInfo` names its default source -- the same fact,
/// spelled out in the type library, and the only one of the two that crosses
/// Wine's process boundary: Wine has no proxy for `IProvideClassInfo2`, so
/// out of process the first step never answers. Failing that too, a container
/// with exactly ONE connection point leaves no choice, and taking the only
/// option is not a guess. More than one and this fails: nothing in that list
/// says which is the default, and picking the first is a coin toss that would
/// silently subscribe to the wrong interface.
///
/// The coclass step comes BEFORE the enumeration, not after it as a fallback.
/// Measured on the MSForms 2.0 extenders Excel hands out for its controls
/// (`OLEObject.Object` on a sheet, `Controls.Item(name)` on a UserForm): the
/// worksheet one answers `EnumConnectionPoints` with `E_NOTIMPL`, the UserForm
/// one enumerates three points, so an enumeration-first order never reaches
/// this step for exactly the objects it exists for.
pub fn source_iid(disp: &IDispatch) -> ComResult<GUID> {
    if let Ok(pci) = disp.cast::<IProvideClassInfo2>() {
        if let Ok(guid) = unsafe { pci.GetGUID(GUIDKIND_DEFAULT_SOURCE_DISP_IID.0 as u32) } {
            return Ok(guid);
        }
    }
    if let Ok(pci) = disp.cast::<IProvideClassInfo>() {
        let named = unsafe { pci.GetClassInfo() }.ok().and_then(|ti| coclass_default_source(&ti));
        if let Some(guid) = named {
            return Ok(guid);
        }
    }

    let cpc: IConnectionPointContainer = disp.cast().map_err(|_| {
        ComError::new(
            E_NOINTERFACE,
            "this object is not an event source: it has no IConnectionPointContainer",
        )
    })?;
    let en = unsafe { cpc.EnumConnectionPoints() }.map_err(ComError::from)?;

    let mut found: Vec<IConnectionPoint> = Vec::new();
    loop {
        let mut buf: [Option<IConnectionPoint>; 1] = [None];
        let mut fetched = 0u32;
        // `pcfetched` is a bare `*mut u32` in this windows version, not the
        // `Option<&mut u32>` the other out-params would suggest.
        let hr = unsafe { en.Next(&mut buf, &mut fetched) };
        if hr.is_err() || fetched == 0 {
            break;
        }
        if let Some(cp) = buf[0].take() {
            found.push(cp);
        }
        if found.len() > 1 {
            break;
        }
    }

    match found.len() {
        0 => Err(ComError::new(
            E_NOINTERFACE,
            "this object is not an event source: it offers no connection points",
        )),
        1 => unsafe { found[0].GetConnectionInterface() }.map_err(ComError::from),
        _ => Err(ComError::new(
            E_NOINTERFACE,
            "this object offers several connection points and does not say which is the \
             default source (no IProvideClassInfo2, and no coclass typeinfo naming one), so \
             there is no way to tell which one to subscribe to",
        )),
    }
}

/// A live subscription, and the guarantee that it is torn down.
///
/// Dropping this unadvises. That is deliberately not left to the call site:
/// forgetting costs more than a stale subscription, because the source still
/// holds a reference to the sink, so the sink's `Sender` is never dropped and
/// the receiving end never learns the session is over. One forgotten call
/// leaks a COM object, a channel, and every future event into it.
#[derive(Debug)]
pub struct Advised {
    pub point: IConnectionPoint,
    /// The connection cookie, or `None` once it has been handed back. Private
    /// and taken rather than copied, so an explicit [`Advised::unadvise`] and
    /// the `Drop` below cannot both spend it: a second `Unadvise` with a
    /// stale cookie is at best `CONNECT_E_NOCONNECTION` and at worst hits
    /// whatever new subscription has since been given that number.
    cookie: Option<u32>,
    /// DISPID to event name, read once from the source interface's typeinfo.
    /// Empty when the typeinfo could not be read; callers fall back to a
    /// `DISPID_<n>` name rather than dropping the event.
    pub names: HashMap<i32, String>,
}

impl Advised {
    /// Unadvise now, surfacing the HRESULT, and make the `Drop` below a no-op.
    ///
    /// Idempotent: calling it twice, or calling it and then dropping,
    /// unadvises exactly once.
    ///
    /// Callers that have nothing to do with a failure should just drop the
    /// `Advised`; this exists for the ones that want to know, and for tearing
    /// a subscription down at a chosen moment (COM pointers must be released
    /// while their apartment still exists, so "at the end of the scope" is
    /// not always soon enough).
    pub fn unadvise(&mut self) -> ComResult<()> {
        let Some(cookie) = self.cookie.take() else { return Ok(()) };
        unsafe { self.point.Unadvise(cookie) }.map_err(ComError::from)
    }
}

impl Drop for Advised {
    fn drop(&mut self) {
        // A failure here cannot be returned and must not panic (this can run
        // during an unwind), but it must not vanish either: it means the
        // source still holds the sink, so it is exactly the leak this type
        // exists to prevent. Report it and carry on.
        if let Err(e) = self.unadvise() {
            eprintln!(
                "sink: Unadvise failed while dropping the subscription; the source may still \
                 hold the sink: {e}"
            );
        }
    }
}

/// Refuse to create a sink anywhere but on a single-threaded apartment.
///
/// The sink is neither `Send` nor `Sync`, and its `Invoke` takes `&self` to a
/// `Sender` that is `!Sync`: two concurrent callbacks would be a data race by
/// the type system's own rules. What keeps them from being concurrent is not
/// Rust but COM -- calls to an object owned by an STA are serialised onto that
/// apartment's one thread. Initialise the thread MTA instead and that
/// guarantee is gone: COM delivers `Invoke` on arbitrary RPC threads, in
/// parallel, and the comment about apartments becomes a false claim about
/// live code.
///
/// So the invariant is checked where it is *established* -- the thread that
/// advises is the thread that owns the sink -- rather than where it would be
/// violated. Checking in `Invoke` instead (recording the creating thread id
/// and comparing) was the alternative: it detects the same mistake later, at
/// a point where the only available responses are to abort the process or to
/// silently drop the event, and it cannot report anything to the programmer
/// who made it. Refusing here returns an ordinary `ComResult` error, at the
/// call site that can still fix it, before any object exists to be raced.
fn require_sta() -> ComResult<()> {
    let mut apartment = APTTYPE::default();
    let mut qualifier = APTTYPEQUALIFIER::default();
    unsafe { CoGetApartmentType(&mut apartment, &mut qualifier) }.map_err(|e| {
        ComError::new(
            e.code(),
            "advise: this thread has no COM apartment, so it cannot own an event sink",
        )
    })?;
    if apartment != APTTYPE_STA && apartment != APTTYPE_MAINSTA {
        return Err(ComError::new(
            RPC_E_WRONG_THREAD,
            format!(
                "advise: an event sink may only be created on a single-threaded apartment \
                 (this thread's apartment type is {}, qualifier {}); on an MTA thread COM \
                 delivers callbacks concurrently and the sink is not thread-safe",
                apartment.0, qualifier.0
            ),
        ));
    }
    Ok(())
}

/// Connection interfaces this sink must never advise on, by IID.
///
/// THE RULE: the sink refuses to advise a source interface it cannot actually
/// implement. This list is the explicit half of that rule.
///
/// `source_iid` takes the only connection point on offer when there is exactly
/// one, and the only one on offer is not always a dispinterface. MEASURED on
/// this Excel under this Wine: a Range offers exactly one connection point and
/// it is `IPropertyNotifySink` (9BFBBC02-EFF1-101A-84ED-00AA00341D07), so the
/// Advise used to succeed -- with zero event names, since that IID names
/// nothing in Excel's type library. `IPropertyNotifySink` is a VTABLE
/// interface: its slot 3 is `OnChanged(DISPID)`, and slot 3 of the `IDispatch`
/// vtable this sink answers `QueryInterface(source_iid)` with is
/// `GetTypeInfoCount(*mut u32)`. One property notification and the DISPID
/// arrives where a `u32` pointer is expected and gets a zero written through
/// it. Refusing the Advise is what keeps that unreachable.
///
/// The typekind check below is the general half of the rule and cannot replace
/// this one: `IPropertyNotifySink` is declared in stdole, not in the object's
/// own type library, so `GetTypeInfoOfGuid` fails for it and it lands squarely
/// in the "cannot tell" bucket that check is required to let through.
const UNIMPLEMENTABLE_SOURCE_IIDS: [(GUID, &str); 4] = [
    (IPropertyNotifySink::IID, "IPropertyNotifySink"),
    (IAdviseSink::IID, "IAdviseSink"),
    (IAdviseSink2::IID, "IAdviseSink2"),
    (IAdviseSinkEx::IID, "IAdviseSinkEx"),
];

/// Refuse a connection interface this sink is known not to implement, naming
/// it. See [`UNIMPLEMENTABLE_SOURCE_IIDS`].
fn refuse_unimplementable_source(iid: &GUID) -> ComResult<()> {
    for (denied, name) in UNIMPLEMENTABLE_SOURCE_IIDS {
        if *iid == denied {
            return Err(ComError::new(
                E_NOINTERFACE,
                format!(
                    "advise: this object's only connection point is {name}, a vtable interface \
                     rather than a dispinterface. An IDispatch sink cannot implement it -- \
                     answering its QueryInterface would hand the source a vtable whose slots \
                     mean something else -- so the subscription is refused instead"
                ),
            ));
        }
    }
    Ok(())
}

/// The typekind of the source interface `iid`, or `None` if it cannot be read.
///
/// Same route as [`dispid_names`]: the object's own typeinfo -> the library
/// containing it -> `GetTypeInfoOfGuid(iid)`. Every step is allowed to fail,
/// and failing means "cannot tell", not "not a dispinterface" -- see
/// [`refuse_non_dispinterface_typekind`], which is where that distinction is
/// made.
fn source_typekind(disp: &IDispatch, iid: &GUID) -> Option<TYPEKIND> {
    let ti = unsafe { disp.GetTypeInfo(0, LOCALE_USER_DEFAULT) }.ok()?;

    let mut lib: Option<ITypeLib> = None;
    let mut index = 0u32;
    unsafe { ti.GetContainingTypeLib(&mut lib, &mut index) }.ok()?;
    let source = unsafe { lib?.GetTypeInfoOfGuid(iid) }.ok()?;

    let attr = unsafe { source.GetTypeAttr() }.ok()?;
    let kind = unsafe { (*attr).typekind };
    unsafe { source.ReleaseTypeAttr(attr) };
    Some(kind)
}

/// The general half of the rule, in its PERMISSIVE form: refuse a source
/// interface the type library says is not a dispinterface, and proceed when
/// the type library says nothing at all.
///
/// "Cannot tell" and "told, and it is not a dispinterface" are different
/// outcomes and only the second refuses. Getting that backwards would break a
/// binding requirement rather than tighten one: an object whose typeinfo
/// cannot be read must keep delivering its events as `DISPID_<n>`, which is
/// exactly what `dispid_names` returning an empty map is for. Sources with no
/// readable typeinfo are ordinary here -- every in-process control that ships
/// without a type library is one -- so refusing them would take the module's
/// most permissive documented path and turn it into an error.
///
/// It is defence in depth, not the defence: the IID denylist above is what
/// closes the case this project has actually measured, because that IID is not
/// in the object's library and so arrives here as `None`.
fn refuse_non_dispinterface_typekind(kind: Option<TYPEKIND>, iid: &GUID) -> ComResult<()> {
    match kind {
        Some(kind) if kind != TKIND_DISPATCH => Err(ComError::new(
            E_NOINTERFACE,
            format!(
                "advise: this object's connection interface {iid:?} is declared with typekind \
                 {} rather than TKIND_DISPATCH ({}), so it is not a dispinterface and an \
                 IDispatch sink cannot implement it",
                kind.0, TKIND_DISPATCH.0
            ),
        )),
        _ => Ok(()),
    }
}

pub fn advise(disp: &IDispatch, handle: u64, tx: Sender<RawEvent>) -> ComResult<Advised> {
    require_sta()?;
    let iid = source_iid(disp)?;
    // The sink refuses to advise a source interface it cannot actually
    // implement. Both halves of that rule are checked HERE -- where the IID is
    // chosen and before any sink object exists -- for the same reason
    // `require_sta` is: this is the call site that can still be told, and an
    // ordinary `ComResult` error reaches it. Checking inside `query_interface`
    // instead would be too late to do anything but lie or crash.
    refuse_unimplementable_source(&iid)?;
    refuse_non_dispinterface_typekind(source_typekind(disp, &iid), &iid)?;
    let cpc: IConnectionPointContainer = disp.cast().map_err(ComError::from)?;
    let point = unsafe { cpc.FindConnectionPoint(&iid) }.map_err(ComError::from)?;

    let sink = EventSink::new(tx, handle, iid);
    let cookie = unsafe { point.Advise(&sink) }.map_err(ComError::from)?;

    Ok(Advised { point, cookie: Some(cookie), names: dispid_names(disp, &iid) })
}

/// DISPID to name for a source dispinterface, via the type library that
/// declares it.
///
/// The route is `disp`'s own typeinfo -> the library containing it ->
/// `GetTypeInfoOfGuid(iid)`, which is the source dispinterface's typeinfo
/// whether or not the object exposes `IProvideClassInfo2`. Measured: Excel's
/// Application does NOT expose `IProvideClassInfo2` under this Wine, so the
/// coclass route -- the obvious one, since `source_iid` reaches for that
/// interface first -- finds nothing at all here, and an empty map costs every
/// event its name.
///
/// Best effort: an object with no readable typeinfo returns an empty map, and
/// the caller names the event `DISPID_<n>` instead of discarding it.
fn dispid_names(disp: &IDispatch, iid: &GUID) -> HashMap<i32, String> {
    let mut out = HashMap::new();
    let Ok(ti) = (unsafe { disp.GetTypeInfo(0, LOCALE_USER_DEFAULT) }) else { return out };

    let mut lib: Option<ITypeLib> = None;
    let mut index = 0u32;
    if unsafe { ti.GetContainingTypeLib(&mut lib, &mut index) }.is_err() {
        return out;
    }
    let Some(lib) = lib else { return out };
    let Ok(source) = (unsafe { lib.GetTypeInfoOfGuid(iid) }) else { return out };

    let Ok(attr) = (unsafe { source.GetTypeAttr() }) else { return out };
    let func_count = unsafe { (*attr).cFuncs };
    unsafe { source.ReleaseTypeAttr(attr) };

    for f in 0..func_count as u32 {
        let Ok(fd) = (unsafe { source.GetFuncDesc(f) }) else { continue };
        let memid = unsafe { (*fd).memid };
        // A dispinterface's typeinfo begins with the seven IUnknown and
        // IDispatch members oleaut32 synthesises in front of the real ones,
        // all flagged FRESTRICTED. They are not events, and finding
        // "QueryInterface" in a list of event names is how a client would
        // discover that the hard way.
        let restricted = unsafe { (*fd).wFuncFlags }.0 & FUNCFLAG_FRESTRICTED.0 != 0;
        if !restricted {
            let mut names = [BSTR::new()];
            let mut got = 0u32;
            if unsafe { source.GetNames(memid, &mut names, &mut got) }.is_ok() && got > 0 {
                out.insert(memid, names[0].to_string());
            }
        }
        unsafe { source.ReleaseFuncDesc(fd) };
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::dispatch::invoke_member;
    use crate::value::variant_to_dispatch;
    use std::sync::mpsc::channel;
    use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

    /// The name of a COM object's own type -- "_Worksheet", "Range" -- read
    /// from its typeinfo's `MEMBERID_NIL` documentation.
    ///
    /// The only way to tell one event argument from another that does not
    /// depend on the arguments already being in the order under test: asking
    /// for a property only one of them has would prove the same thing, but
    /// asking the object what it *is* cannot be argued with.
    fn type_name(v: &VARIANT) -> String {
        let Some(disp) = variant_to_dispatch(v) else { return "<not an object>".to_string() };
        unsafe {
            let Ok(ti) = disp.GetTypeInfo(0, LOCALE_USER_DEFAULT) else {
                return "<no typeinfo>".to_string();
            };
            let mut name = BSTR::new();
            // MEMBERID_NIL (-1) asks for the type's own documentation rather
            // than a member's.
            if ti.GetDocumentation(-1, Some(&mut name), None, std::ptr::null_mut(), None).is_err() {
                return "<no documentation>".to_string();
            }
            name.to_string()
        }
    }

    /// Advising on a live Excel Application and changing a cell must call
    /// back into the sink. This is the whole point of Task 2's pump: without
    /// it the Advise succeeds and nothing ever arrives.
    ///
    /// It also pins the argument ORDER, which is why it changes a cell rather
    /// than settling for the first event Excel happens to raise. The typelib
    /// declares `SheetChange(Sh, Target)`, so argument one must be the
    /// Worksheet and argument two the Range. Counting the arguments would not
    /// catch a swap -- `(Range, Worksheet)` is two arguments too, and reads
    /// perfectly well -- so this names their types instead. It is the
    /// assertion that decided how [`ordered_args`] reads a DISPPARAMS, and it
    /// is the only one that sees the shape a live Excel actually sends;
    /// `test_invoke_puts_the_arguments_in_declaration_order` covers the shapes
    /// Excel cannot be made to produce.
    #[test]
    fn test_advising_excel_delivers_a_sheet_change() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let app = crate::dispatch::create_instance("Excel.Application").expect("create Excel");
        // Without this, `Quit` on a workbook with unsaved changes puts up a
        // save prompt no one can answer and leaves EXCEL.EXE running for
        // every later test to trip over.
        invoke_member(&app, "DisplayAlerts=", vec![VARIANT::from(false)], vec![])
            .expect("DisplayAlerts=false");
        let (tx, rx) = channel();
        let mut advised = advise(&app, 1, tx).expect("advise");

        // A name is the point of the typeinfo lookup -- a DISPID alone would
        // make the client's on('SheetChange') impossible to satisfy.
        assert!(
            advised.names.values().any(|n| n == "SheetChange"),
            "the source interface's names must resolve; got {:?}",
            advised.names.values().take(5).collect::<Vec<_>>()
        );
        // The FRESTRICTED skip in `dispid_names`, pinned. A dispinterface's
        // typeinfo begins with the seven IUnknown/IDispatch members oleaut32
        // synthesises in front of the real ones, and without the skip they
        // arrive here as event names -- while every other assertion in this
        // test still passes, which is exactly why this one has to exist.
        // Measured on this Excel: 36 members, of which the first seven are
        // restricted and 29 are real events.
        assert!(
            !advised.names.values().any(|n| n == "QueryInterface"),
            "QueryInterface is not an event: the FRESTRICTED members oleaut32 synthesises in \
             front of a dispinterface must be skipped, got {:?}",
            advised.names.values().collect::<Vec<_>>()
        );

        let books = invoke_member(&app, "Workbooks", vec![], vec![]).expect("Workbooks");
        let books = variant_to_dispatch(&books).expect("Workbooks dispatch");
        let book = invoke_member(&books, "Add", vec![], vec![]).expect("Workbooks.Add");
        let book = variant_to_dispatch(&book).expect("Workbook dispatch");
        let sheet = invoke_member(&book, "ActiveSheet", vec![], vec![]).expect("ActiveSheet");
        let sheet = variant_to_dispatch(&sheet).expect("ActiveSheet dispatch");
        let cell = invoke_member(&sheet, "Range", vec![VARIANT::from("A1")], vec![])
            .expect("Range(A1)");
        let cell = variant_to_dispatch(&cell).expect("Range dispatch");
        // The change that raises SheetChange.
        invoke_member(&cell, "Value=", vec![VARIANT::from(42i32)], vec![]).expect("Range.Value=");

        // Pump until the event arrives or we give up. This test has no
        // session thread, so it drives the pump by hand -- but it BLOCKS on
        // the message queue rather than sleeping and re-checking, because a
        // sleep-and-poll loop is the thing this whole design rejects. The
        // waker here is only there to satisfy wait_timeout's signature; the
        // wake that matters comes from the message queue.
        let waker = crate::pump::Waker::new().expect("Waker::new");
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(20);
        let mut seen: Vec<i32> = Vec::new();
        let mut sheet_change: Option<RawEvent> = None;
        while std::time::Instant::now() < deadline && sheet_change.is_none() {
            crate::pump::wait_timeout(waker.handle(), 250);
            crate::pump::drain_messages();
            while let Ok(ev) = rx.try_recv() {
                seen.push(ev.dispid);
                if advised.names.get(&ev.dispid).map(String::as_str) == Some("SheetChange") {
                    sheet_change = Some(ev);
                    break;
                }
            }
        }

        // Everything that has to be read off a live COM object is read HERE,
        // before the teardown below: an event argument is an IDispatch, and
        // once Excel has quit there is nothing left to ask.
        let arg_types: Vec<String> = sheet_change
            .as_ref()
            .map(|ev| ev.args.iter().map(type_name).collect())
            .unwrap_or_default();
        let names = advised.names.clone();

        // Drop order is the teardown, not an afterthought: every COM pointer
        // -- the event's argument VARIANTs, the connection point, Excel itself
        // -- has to be released while the apartment still exists. Releasing
        // one after CoUninitialize is a call into an apartment that is gone.
        drop(sheet_change);
        advised.unadvise().expect("Unadvise must succeed");
        drop(advised);

        // The sink must now be FREED, not merely disconnected -- and from
        // outside, nothing can see that directly: the object owns itself
        // through its refcount, so a missing `Release`, an extra `AddRef` in
        // `QueryInterface` and a box that is never dropped all look identical
        // from here. All three have one observable consequence, though. The
        // sink owns the `Sender`; freeing the sink drops it, and a dropped
        // `Sender` is the one thing a `Receiver` can detect. So pump until the
        // channel disconnects: Excel's release of its marshalled reference
        // comes back over RPC and only arrives if the queue is drained.
        let mut sink_freed = false;
        let free_deadline = std::time::Instant::now() + std::time::Duration::from_secs(20);
        while std::time::Instant::now() < free_deadline && !sink_freed {
            crate::pump::wait_timeout(waker.handle(), 250);
            crate::pump::drain_messages();
            loop {
                match rx.try_recv() {
                    // A late event: drop its VARIANTs here, while the
                    // apartment that owns them is still alive.
                    Ok(_) => continue,
                    Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                        sink_freed = true;
                        break;
                    }
                    Err(std::sync::mpsc::TryRecvError::Empty) => break,
                }
            }
        }

        drop(cell);
        drop(sheet);
        drop(book);
        drop(books);
        let _ = invoke_member(&app, "Quit", vec![], vec![]);
        drop(app);
        unsafe { CoUninitialize() };

        assert!(
            sink_freed,
            "the sink was never freed: after Unadvise and dropping the subscription its \
             refcount must reach 0, which drops the Sender and disconnects this channel"
        );
        assert!(!seen.is_empty(), "no event arrived within 20s");
        for dispid in &seen {
            assert!(names.contains_key(dispid), "the event's DISPID must be a known name");
        }
        assert_eq!(
            arg_types.len(),
            2,
            "SheetChange(Sh, Target) has two arguments; got {arg_types:?} from {seen:?}"
        );
        assert!(
            arg_types[0].contains("Worksheet"),
            "SheetChange's first argument is the sheet; the sink delivered {arg_types:?}"
        );
        assert!(
            arg_types[1].contains("Range"),
            "SheetChange's second argument is the changed range, not {:?}",
            arg_types[1]
        );
    }

    /// The bug this pins: a worksheet ActiveX control's MSForms extender
    /// (`OLEObject.Object`) has no reachable `IProvideClassInfo2` under Wine
    /// and answers `EnumConnectionPoints` with `E_NOTIMPL`, so before the
    /// coclass lookup `advise` failed on it outright (measured 2026-09-04,
    /// spec M1). Its coclass typeinfo names `CommandButtonEvents`; advising
    /// that delivers `Click`. Setting `Value = true` on a CommandButton is how
    /// a click is raised without a UI. Same teardown discipline as the
    /// SheetChange test above, for the same reasons.
    #[test]
    fn test_advising_a_worksheet_activex_control_delivers_click() {
        let _guard = crate::dispatch::lock_excel_for_test();
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let app = crate::dispatch::create_instance("Excel.Application").expect("create Excel");
        invoke_member(&app, "DisplayAlerts=", vec![VARIANT::from(false)], vec![])
            .expect("DisplayAlerts=false");

        let books = invoke_member(&app, "Workbooks", vec![], vec![]).expect("Workbooks");
        let books = variant_to_dispatch(&books).expect("Workbooks dispatch");
        let book = invoke_member(&books, "Add", vec![], vec![]).expect("Workbooks.Add");
        let book = variant_to_dispatch(&book).expect("Workbook dispatch");
        let sheet = invoke_member(&book, "ActiveSheet", vec![], vec![]).expect("ActiveSheet");
        let sheet = variant_to_dispatch(&sheet).expect("ActiveSheet dispatch");
        let objects = invoke_member(&sheet, "OLEObjects", vec![], vec![]).expect("OLEObjects");
        let objects = variant_to_dispatch(&objects).expect("OLEObjects dispatch");
        // All five arguments, named: Excel 11 answers 0x800A03EC to an Add
        // that leaves Width and Height out.
        let named = |name: &str, v: VARIANT| (name.to_string(), v);
        let host = invoke_member(
            &objects,
            "Add",
            vec![],
            vec![
                named("ClassType", VARIANT::from("Forms.CommandButton.1")),
                named("Left", VARIANT::from(10.0f64)),
                named("Top", VARIANT::from(10.0f64)),
                named("Width", VARIANT::from(80.0f64)),
                named("Height", VARIANT::from(24.0f64)),
            ],
        )
        .expect("OLEObjects.Add(Forms.CommandButton.1)");
        let host = variant_to_dispatch(&host).expect("OLEObject dispatch");
        // The extender, not the OLEObject host (CONNECT_E_NOCONNECTION) and
        // not the extender's own `.Object` (no IConnectionPointContainer).
        let control = invoke_member(&host, "Object", vec![], vec![]).expect("OLEObject.Object");
        let control = variant_to_dispatch(&control).expect("MSForms control dispatch");

        let (tx, rx) = channel();
        let mut advised = advise(&control, 7, tx).expect("advise on the MSForms extender");
        assert!(
            advised.names.values().any(|n| n == "Click"),
            "CommandButtonEvents' names must resolve; got {:?}",
            advised.names.values().take(5).collect::<Vec<_>>()
        );

        invoke_member(&control, "Value=", vec![VARIANT::from(true)], vec![])
            .expect("CommandButton.Value=");

        let waker = crate::pump::Waker::new().expect("Waker::new");
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(20);
        let mut seen: Vec<i32> = Vec::new();
        let mut click: Option<RawEvent> = None;
        while std::time::Instant::now() < deadline && click.is_none() {
            crate::pump::wait_timeout(waker.handle(), 250);
            crate::pump::drain_messages();
            while let Ok(ev) = rx.try_recv() {
                seen.push(ev.dispid);
                if advised.names.get(&ev.dispid).map(String::as_str) == Some("Click") {
                    click = Some(ev);
                    break;
                }
            }
        }
        let click_arity = click.as_ref().map(|ev| ev.args.len());
        let names = advised.names.clone();

        drop(click);
        advised.unadvise().expect("Unadvise must succeed");
        drop(advised);

        let mut sink_freed = false;
        let free_deadline = std::time::Instant::now() + std::time::Duration::from_secs(20);
        while std::time::Instant::now() < free_deadline && !sink_freed {
            crate::pump::wait_timeout(waker.handle(), 250);
            crate::pump::drain_messages();
            loop {
                match rx.try_recv() {
                    Ok(_) => continue,
                    Err(std::sync::mpsc::TryRecvError::Disconnected) => {
                        sink_freed = true;
                        break;
                    }
                    Err(std::sync::mpsc::TryRecvError::Empty) => break,
                }
            }
        }

        drop(control);
        drop(host);
        drop(objects);
        drop(sheet);
        drop(book);
        drop(books);
        let _ = invoke_member(&app, "Quit", vec![], vec![]);
        drop(app);
        unsafe { CoUninitialize() };

        assert!(
            sink_freed,
            "the sink was never freed: after Unadvise and dropping the subscription its \
             refcount must reach 0, which drops the Sender and disconnects this channel"
        );
        assert_eq!(
            click_arity,
            Some(0),
            "Click() has no arguments and must arrive within 20s; events seen: {seen:?}, \
             names: {names:?}"
        );
    }

    // ------------------------------------------------------------------
    // A hand-written COM event source.
    //
    // `source_iid`'s headline promise is that it never guesses, and Excel
    // cannot test that promise: Excel's Application offers exactly ONE
    // connection point, so the branch that matters -- several of them, and no
    // `IProvideClassInfo2` to say which is the default -- is unreachable from
    // every real object this project drives. Rewriting that arm to
    // `GetConnectionInterface()` on the first one, the precise mistake the
    // module exists to refuse, leaves a suite that only knows Excel entirely
    // green. The same goes for an object with no connection points and for one
    // that is not a connection point container at all.
    //
    // Hence a fake source: in-process, no Excel, no marshalling. What it is
    // entitled to simplify as a test double, said out loud so nobody reads it
    // as a statement about COM: `QueryInterface` for
    // `IConnectionPointContainer` hands back a SEPARATE object rather than a
    // second vtable on this one (real COM demands one stable `IUnknown`
    // identity; nothing under test asks for it), and the `IDispatch` members
    // are `E_NOTIMPL` because `source_iid` calls none of them -- it only ever
    // casts.
    // ------------------------------------------------------------------

    use std::cell::{Cell, RefCell};
    use std::rc::Rc;
    use std::sync::mpsc::TryRecvError;
    use windows::Win32::Foundation::{CO_E_NOTINITIALIZED, S_FALSE};
    use windows::Win32::System::Com::{
        IConnectionPointContainer_Vtbl, IConnectionPoint_Vtbl, IEnumConnectionPoints_Vtbl,
        COINIT_MULTITHREADED,
    };
    use windows::Win32::System::Ole::{
        IProvideClassInfo_Vtbl, LoadRegTypeLib, CONNECT_E_CANNOTCONNECT, CONNECT_E_NOCONNECTION,
    };

    /// Two IIDs that are nobody's real interface, so a test that gets one back
    /// knows exactly where it came from.
    const FAKE_SOURCE_A: GUID = GUID::from_u128(0x11111111_2222_3333_4444_555555555555);
    const FAKE_SOURCE_B: GUID = GUID::from_u128(0x66666666_7777_8888_9999_aaaaaaaaaaaa);
    /// The cookie the fake connection point hands out. Not 0 and not 1: a
    /// cookie the code under test happened to invent would pass by accident.
    const FAKE_COOKIE: u32 = 0x5EED;

    /// What the fake connection point saw, shared with the test that made it.
    #[derive(Default)]
    struct FakeLog {
        advise_calls: u32,
        unadvise_calls: u32,
        /// How often the container was asked to enumerate its points.
        enum_calls: u32,
        /// The sink's answer to `QueryInterface(source_iid)`, asked the way
        /// ATL's `IConnectionPointImpl::Advise` asks it.
        sink_answered_source_iid: bool,
    }
    type Log = Rc<RefCell<FakeLog>>;

    fn log() -> Log {
        Rc::new(RefCell::new(FakeLog::default()))
    }

    // The four IDispatch members the fake source does not implement. Shared by
    // its vtable; `source_iid` and `advise` between them call none of these,
    // and `dispid_names` is documented to return an empty map when the
    // typeinfo cannot be read -- which is the path E_NOTIMPL sends it down.
    unsafe extern "system" fn fake_type_info_count(_this: *mut c_void, count: *mut u32) -> HRESULT {
        if !count.is_null() {
            unsafe { *count = 0 };
        }
        S_OK
    }
    unsafe extern "system" fn fake_type_info(
        _this: *mut c_void,
        _index: u32,
        _lcid: u32,
        out: *mut *mut c_void,
    ) -> HRESULT {
        if !out.is_null() {
            unsafe { *out = std::ptr::null_mut() };
        }
        E_NOTIMPL
    }
    unsafe extern "system" fn fake_ids_of_names(
        _this: *mut c_void,
        _riid: *const GUID,
        _names: *const PCWSTR,
        _count: u32,
        _lcid: u32,
        _ids: *mut i32,
    ) -> HRESULT {
        E_NOTIMPL
    }
    #[allow(clippy::too_many_arguments)]
    unsafe extern "system" fn fake_dispatch_invoke(
        _this: *mut c_void,
        _dispid: i32,
        _riid: *const GUID,
        _lcid: u32,
        _flags: DISPATCH_FLAGS,
        _params: *const DISPPARAMS,
        _result: *mut std::mem::MaybeUninit<VARIANT>,
        _excep: *mut EXCEPINFO,
        _arg_err: *mut u32,
    ) -> HRESULT {
        E_NOTIMPL
    }

    #[repr(C)]
    struct FakeSource {
        vtable: &'static IDispatch_Vtbl,
        refcount: AtomicU32,
        /// One connection point per IID here: none, one, or several.
        points: Vec<GUID>,
        /// Whether it admits to being an `IConnectionPointContainer` at all.
        offers_container: bool,
        /// `None`: no `IProvideClassInfo` at all -- the Excel this project
        /// measured. `Some(None)`: the interface, but `GetClassInfo` fails.
        /// `Some(Some(ti))`: a coclass typeinfo to hand out.
        class_info: Option<Option<ITypeInfo>>,
        log: Log,
    }

    static FAKE_SOURCE_VTABLE: IDispatch_Vtbl = IDispatch_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: FakeSource::query_interface,
            AddRef: FakeSource::add_ref,
            Release: FakeSource::release,
        },
        GetTypeInfoCount: fake_type_info_count,
        GetTypeInfo: fake_type_info,
        GetIDsOfNames: fake_ids_of_names,
        Invoke: fake_dispatch_invoke,
    };

    impl FakeSource {
        fn new(points: Vec<GUID>, offers_container: bool, log: Log) -> IDispatch {
            Self::build(points, offers_container, None, log)
        }

        /// A source that also answers `IProvideClassInfo`; see `class_info`.
        fn with_class_info(
            points: Vec<GUID>,
            offers_container: bool,
            class_info: Option<ITypeInfo>,
            log: Log,
        ) -> IDispatch {
            Self::build(points, offers_container, Some(class_info), log)
        }

        fn build(
            points: Vec<GUID>,
            offers_container: bool,
            class_info: Option<Option<ITypeInfo>>,
            log: Log,
        ) -> IDispatch {
            let boxed = Box::new(FakeSource {
                vtable: &FAKE_SOURCE_VTABLE,
                refcount: AtomicU32::new(1),
                points,
                offers_container,
                class_info,
                log,
            });
            unsafe { IDispatch::from_raw(Box::into_raw(boxed) as *mut c_void) }
        }

        unsafe extern "system" fn query_interface(
            this: *mut c_void,
            iid: *const GUID,
            interface: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || interface.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeSource) };
            if iid == IUnknown::IID || iid == IDispatch::IID {
                obj.refcount.fetch_add(1, Ordering::Relaxed);
                unsafe { *interface = this };
                return S_OK;
            }
            // Never IProvideClassInfo2 -- exactly like every out-of-process
            // object under this Wine, which has no proxy for it; that is what
            // puts `source_iid` onto the coclass and connection point routes.
            if iid == IProvideClassInfo::IID {
                if let Some(class_info) = &obj.class_info {
                    unsafe { *interface = FakeProvideClassInfo::new(class_info.clone()) };
                    return S_OK;
                }
            }
            if iid == IConnectionPointContainer::IID && obj.offers_container {
                unsafe { *interface = FakeCpc::new(obj.points.clone(), obj.log.clone()) };
                return S_OK;
            }
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }

        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            let obj = unsafe { &*(this as *const FakeSource) };
            obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
        }

        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            let ptr = this as *mut FakeSource;
            let remaining =
                unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
            if remaining == 0 {
                drop(unsafe { Box::from_raw(ptr) });
            }
            remaining
        }
    }

    /// `IProvideClassInfo` as a separate object, like `FakeCpc`: hands out the
    /// typeinfo it was given, or fails the way an object with the interface
    /// but no usable class info does.
    #[repr(C)]
    struct FakeProvideClassInfo {
        vtable: &'static IProvideClassInfo_Vtbl,
        refcount: AtomicU32,
        class_info: Option<ITypeInfo>,
    }

    static FAKE_PCI_VTABLE: IProvideClassInfo_Vtbl = IProvideClassInfo_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: FakeProvideClassInfo::query_interface,
            AddRef: FakeProvideClassInfo::add_ref,
            Release: FakeProvideClassInfo::release,
        },
        GetClassInfo: FakeProvideClassInfo::get_class_info,
    };

    impl FakeProvideClassInfo {
        fn new(class_info: Option<ITypeInfo>) -> *mut c_void {
            let boxed = Box::new(FakeProvideClassInfo {
                vtable: &FAKE_PCI_VTABLE,
                refcount: AtomicU32::new(1),
                class_info,
            });
            Box::into_raw(boxed) as *mut c_void
        }

        unsafe extern "system" fn query_interface(
            this: *mut c_void,
            iid: *const GUID,
            interface: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || interface.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeProvideClassInfo) };
            if iid == IUnknown::IID || iid == IProvideClassInfo::IID {
                obj.refcount.fetch_add(1, Ordering::Relaxed);
                unsafe { *interface = this };
                return S_OK;
            }
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }

        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            let obj = unsafe { &*(this as *const FakeProvideClassInfo) };
            obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
        }

        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            let ptr = this as *mut FakeProvideClassInfo;
            let remaining =
                unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
            if remaining == 0 {
                drop(unsafe { Box::from_raw(ptr) });
            }
            remaining
        }

        unsafe extern "system" fn get_class_info(
            this: *mut c_void,
            out: *mut *mut c_void,
        ) -> HRESULT {
            if out.is_null() {
                return E_POINTER;
            }
            let obj = unsafe { &*(this as *const FakeProvideClassInfo) };
            match &obj.class_info {
                // The out-param owns one reference: hand over the clone's.
                Some(ti) => {
                    unsafe { *out = ti.clone().into_raw() };
                    S_OK
                }
                None => {
                    unsafe { *out = std::ptr::null_mut() };
                    E_NOTIMPL
                }
            }
        }
    }

    #[repr(C)]
    struct FakeCpc {
        vtable: &'static IConnectionPointContainer_Vtbl,
        refcount: AtomicU32,
        points: Vec<GUID>,
        log: Log,
    }

    static FAKE_CPC_VTABLE: IConnectionPointContainer_Vtbl = IConnectionPointContainer_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: FakeCpc::query_interface,
            AddRef: FakeCpc::add_ref,
            Release: FakeCpc::release,
        },
        EnumConnectionPoints: FakeCpc::enum_connection_points,
        FindConnectionPoint: FakeCpc::find_connection_point,
    };

    impl FakeCpc {
        fn new(points: Vec<GUID>, log: Log) -> *mut c_void {
            let boxed = Box::new(FakeCpc {
                vtable: &FAKE_CPC_VTABLE,
                refcount: AtomicU32::new(1),
                points,
                log,
            });
            Box::into_raw(boxed) as *mut c_void
        }

        unsafe extern "system" fn query_interface(
            this: *mut c_void,
            iid: *const GUID,
            interface: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || interface.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeCpc) };
            if iid == IUnknown::IID || iid == IConnectionPointContainer::IID {
                obj.refcount.fetch_add(1, Ordering::Relaxed);
                unsafe { *interface = this };
                return S_OK;
            }
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }

        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            let obj = unsafe { &*(this as *const FakeCpc) };
            obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
        }

        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            let ptr = this as *mut FakeCpc;
            let remaining =
                unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
            if remaining == 0 {
                drop(unsafe { Box::from_raw(ptr) });
            }
            remaining
        }

        unsafe extern "system" fn enum_connection_points(
            this: *mut c_void,
            out: *mut *mut c_void,
        ) -> HRESULT {
            if out.is_null() {
                return E_POINTER;
            }
            let obj = unsafe { &*(this as *const FakeCpc) };
            obj.log.borrow_mut().enum_calls += 1;
            unsafe { *out = FakeEnum::new(obj.points.clone(), obj.log.clone()) };
            S_OK
        }

        unsafe extern "system" fn find_connection_point(
            this: *mut c_void,
            iid: *const GUID,
            out: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || out.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeCpc) };
            if obj.points.contains(&iid) {
                unsafe { *out = FakeCp::new(iid, obj.log.clone()) };
                S_OK
            } else {
                unsafe { *out = std::ptr::null_mut() };
                CONNECT_E_NOCONNECTION
            }
        }
    }

    #[repr(C)]
    struct FakeEnum {
        vtable: &'static IEnumConnectionPoints_Vtbl,
        refcount: AtomicU32,
        points: Vec<GUID>,
        position: Cell<usize>,
        log: Log,
    }

    static FAKE_ENUM_VTABLE: IEnumConnectionPoints_Vtbl = IEnumConnectionPoints_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: FakeEnum::query_interface,
            AddRef: FakeEnum::add_ref,
            Release: FakeEnum::release,
        },
        Next: FakeEnum::next,
        Skip: FakeEnum::skip,
        Reset: FakeEnum::reset,
        Clone: FakeEnum::clone_enum,
    };

    impl FakeEnum {
        fn new(points: Vec<GUID>, log: Log) -> *mut c_void {
            let boxed = Box::new(FakeEnum {
                vtable: &FAKE_ENUM_VTABLE,
                refcount: AtomicU32::new(1),
                points,
                position: Cell::new(0),
                log,
            });
            Box::into_raw(boxed) as *mut c_void
        }

        unsafe extern "system" fn query_interface(
            this: *mut c_void,
            iid: *const GUID,
            interface: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || interface.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeEnum) };
            if iid == IUnknown::IID || iid == windows::Win32::System::Com::IEnumConnectionPoints::IID
            {
                obj.refcount.fetch_add(1, Ordering::Relaxed);
                unsafe { *interface = this };
                return S_OK;
            }
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }

        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            let obj = unsafe { &*(this as *const FakeEnum) };
            obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
        }

        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            let ptr = this as *mut FakeEnum;
            let remaining =
                unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
            if remaining == 0 {
                drop(unsafe { Box::from_raw(ptr) });
            }
            remaining
        }

        unsafe extern "system" fn next(
            this: *mut c_void,
            count: u32,
            out: *mut *mut c_void,
            fetched: *mut u32,
        ) -> HRESULT {
            if out.is_null() && count > 0 {
                return E_POINTER;
            }
            let obj = unsafe { &*(this as *const FakeEnum) };
            let mut n = 0usize;
            while n < count as usize && obj.position.get() < obj.points.len() {
                let iid = obj.points[obj.position.get()];
                obj.position.set(obj.position.get() + 1);
                unsafe { *out.add(n) = FakeCp::new(iid, obj.log.clone()) };
                n += 1;
            }
            if !fetched.is_null() {
                unsafe { *fetched = n as u32 };
            }
            // S_FALSE for a short read, which is what a real enumerator does
            // and what `source_iid`'s loop has to cope with.
            if n == count as usize {
                S_OK
            } else {
                S_FALSE
            }
        }

        unsafe extern "system" fn skip(_this: *mut c_void, _count: u32) -> HRESULT {
            E_NOTIMPL
        }
        unsafe extern "system" fn reset(this: *mut c_void) -> HRESULT {
            let obj = unsafe { &*(this as *const FakeEnum) };
            obj.position.set(0);
            S_OK
        }
        unsafe extern "system" fn clone_enum(_this: *mut c_void, out: *mut *mut c_void) -> HRESULT {
            if !out.is_null() {
                unsafe { *out = std::ptr::null_mut() };
            }
            E_NOTIMPL
        }
    }

    #[repr(C)]
    struct FakeCp {
        vtable: &'static IConnectionPoint_Vtbl,
        refcount: AtomicU32,
        iid: GUID,
        /// The sink, held exactly as a real connection point holds it: an
        /// owning reference, released on `Unadvise`.
        sink: RefCell<Option<IUnknown>>,
        log: Log,
    }

    static FAKE_CP_VTABLE: IConnectionPoint_Vtbl = IConnectionPoint_Vtbl {
        base__: IUnknown_Vtbl {
            QueryInterface: FakeCp::query_interface,
            AddRef: FakeCp::add_ref,
            Release: FakeCp::release,
        },
        GetConnectionInterface: FakeCp::get_connection_interface,
        GetConnectionPointContainer: FakeCp::get_connection_point_container,
        Advise: FakeCp::advise,
        Unadvise: FakeCp::unadvise,
        EnumConnections: FakeCp::enum_connections,
    };

    impl FakeCp {
        fn new(iid: GUID, log: Log) -> *mut c_void {
            let boxed = Box::new(FakeCp {
                vtable: &FAKE_CP_VTABLE,
                refcount: AtomicU32::new(1),
                iid,
                sink: RefCell::new(None),
                log,
            });
            Box::into_raw(boxed) as *mut c_void
        }

        unsafe extern "system" fn query_interface(
            this: *mut c_void,
            iid: *const GUID,
            interface: *mut *mut c_void,
        ) -> HRESULT {
            if iid.is_null() || interface.is_null() {
                return E_POINTER;
            }
            let iid = unsafe { std::ptr::read_unaligned(iid) };
            let obj = unsafe { &*(this as *const FakeCp) };
            if iid == IUnknown::IID || iid == IConnectionPoint::IID {
                obj.refcount.fetch_add(1, Ordering::Relaxed);
                unsafe { *interface = this };
                return S_OK;
            }
            unsafe { *interface = std::ptr::null_mut() };
            E_NOINTERFACE
        }

        unsafe extern "system" fn add_ref(this: *mut c_void) -> u32 {
            let obj = unsafe { &*(this as *const FakeCp) };
            obj.refcount.fetch_add(1, Ordering::Relaxed).wrapping_add(1)
        }

        unsafe extern "system" fn release(this: *mut c_void) -> u32 {
            let ptr = this as *mut FakeCp;
            let remaining =
                unsafe { (*ptr).refcount.fetch_sub(1, Ordering::AcqRel) }.wrapping_sub(1);
            if remaining == 0 {
                drop(unsafe { Box::from_raw(ptr) });
            }
            remaining
        }

        unsafe extern "system" fn get_connection_interface(
            this: *mut c_void,
            out: *mut GUID,
        ) -> HRESULT {
            if out.is_null() {
                return E_POINTER;
            }
            let obj = unsafe { &*(this as *const FakeCp) };
            unsafe { *out = obj.iid };
            S_OK
        }

        unsafe extern "system" fn get_connection_point_container(
            _this: *mut c_void,
            out: *mut *mut c_void,
        ) -> HRESULT {
            if !out.is_null() {
                unsafe { *out = std::ptr::null_mut() };
            }
            E_NOTIMPL
        }

        /// ATL's `IConnectionPointImpl::Advise` asks the sink for the source
        /// dispinterface itself and refuses the connection if it says no; this
        /// fake does the same, which is what makes the `source_iid` arm of the
        /// sink's `QueryInterface` load-bearing rather than decorative. Excel's
        /// own connection point settles for `IDispatch`, so nothing driven by
        /// Excel can tell whether that arm is there.
        unsafe extern "system" fn advise(
            this: *mut c_void,
            sink: *mut c_void,
            cookie: *mut u32,
        ) -> HRESULT {
            if sink.is_null() || cookie.is_null() {
                return E_POINTER;
            }
            let obj = unsafe { &*(this as *const FakeCp) };
            obj.log.borrow_mut().advise_calls += 1;

            let unknown = unsafe { IUnknown::from_raw_borrowed(&sink) }.expect("sink is non-null");
            let mut typed: *mut c_void = std::ptr::null_mut();
            let answered = unsafe { unknown.query(&obj.iid, &mut typed) }.is_ok();
            obj.log.borrow_mut().sink_answered_source_iid = answered;
            if !answered {
                unsafe { *cookie = 0 };
                return CONNECT_E_CANNOTCONNECT;
            }
            // That QueryInterface AddRef'd; the reference below is the one
            // this connection point keeps.
            drop(unsafe { IUnknown::from_raw(typed) });

            *obj.sink.borrow_mut() = Some(unknown.clone());
            unsafe { *cookie = FAKE_COOKIE };
            S_OK
        }

        unsafe extern "system" fn unadvise(this: *mut c_void, cookie: u32) -> HRESULT {
            let obj = unsafe { &*(this as *const FakeCp) };
            obj.log.borrow_mut().unadvise_calls += 1;
            if cookie != FAKE_COOKIE {
                return CONNECT_E_NOCONNECTION;
            }
            match obj.sink.borrow_mut().take() {
                Some(_) => S_OK,
                None => CONNECT_E_NOCONNECTION,
            }
        }

        unsafe extern "system" fn enum_connections(
            _this: *mut c_void,
            out: *mut *mut c_void,
        ) -> HRESULT {
            if !out.is_null() {
                unsafe { *out = std::ptr::null_mut() };
            }
            E_NOTIMPL
        }
    }

    /// The refusal that is this module's headline safety property: several
    /// connection points, nothing saying which is the default, so there is no
    /// answer and `source_iid` must not invent one.
    #[test]
    fn test_source_iid_refuses_to_choose_between_two_connection_points() {
        let source = FakeSource::new(vec![FAKE_SOURCE_A, FAKE_SOURCE_B], true, log());
        let err = source_iid(&source)
            .expect_err("two connection points and no IProvideClassInfo2 must not be guessed at");
        assert_eq!(err.code, E_NOINTERFACE, "got {err}");
        assert!(
            err.to_string().contains("several connection points"),
            "the error must say why it refused, not just that it failed: {err}"
        );
    }

    /// The same object with one connection point, so the refusal above is a
    /// decision and not the fake simply never working.
    #[test]
    fn test_source_iid_takes_the_only_connection_point() {
        let source = FakeSource::new(vec![FAKE_SOURCE_A], true, log());
        let iid = source_iid(&source).expect("one connection point leaves no choice");
        assert_eq!(iid, FAKE_SOURCE_A);
    }

    #[test]
    fn test_source_iid_reports_an_object_with_no_connection_points() {
        let source = FakeSource::new(vec![], true, log());
        let err = source_iid(&source).expect_err("an empty container is not an event source");
        assert_eq!(err.code, E_NOINTERFACE, "got {err}");
        assert!(
            err.to_string().contains("no connection points"),
            "the error must name the reason: {err}"
        );
    }

    #[test]
    fn test_source_iid_reports_an_object_that_is_not_a_connection_point_container() {
        let source = FakeSource::new(vec![FAKE_SOURCE_A], false, log());
        let err = source_iid(&source)
            .expect_err("an object with no IConnectionPointContainer is not an event source");
        assert_eq!(err.code, E_NOINTERFACE, "got {err}");
        assert!(
            err.to_string().contains("IConnectionPointContainer"),
            "the error must name the missing interface: {err}"
        );
    }

    /// MSForms 2.0's type library, its `CommandButton` coclass -- the control
    /// every measurement in this project places -- and the events
    /// dispinterface that coclass marks as its default source.
    const MSFORMS_TYPELIB: GUID = GUID::from_u128(0x0D452EE1_E08F_101A_852E_02608C4D0BB4);
    const MSFORMS_COMMAND_BUTTON: GUID = GUID::from_u128(0xD7053240_CE69_11CD_A777_00DD01143C57);
    const COMMAND_BUTTON_EVENTS: GUID = GUID::from_u128(0x7B020EC1_AF6C_11CE_9F46_00AA00574A4F);

    /// A typeinfo straight from the registered MSForms library. No Excel:
    /// oleaut32 loads FM20.DLL's library into this process. The caller has
    /// initialised an apartment and drops the result before uninitialising it.
    fn msforms_typeinfo(guid: &GUID) -> ITypeInfo {
        let lib = unsafe { LoadRegTypeLib(&MSFORMS_TYPELIB, 2, 0, 0) }
            .expect("the MSForms 2.0 type library (FM20.DLL) is registered on this host");
        unsafe { lib.GetTypeInfoOfGuid(guid) }.expect("typeinfo by GUID")
    }

    /// The lookup reads what the type library says: MSForms' `CommandButton`
    /// implements two interfaces, and the one flagged FSOURCE|FDEFAULT is
    /// `CommandButtonEvents`. Measured before this code existed
    /// (`.superpowers/sdd/diag-msforms-events.md`): that is also the IID whose
    /// Advise delivers `Click`.
    #[test]
    fn test_coclass_default_source_names_the_command_button_events() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        {
            let coclass = msforms_typeinfo(&MSFORMS_COMMAND_BUTTON);
            assert_eq!(coclass_default_source(&coclass), Some(COMMAND_BUTTON_EVENTS));
        }
        unsafe { CoUninitialize() };
    }

    /// A typeinfo that lists no source interface -- here a dispinterface
    /// standing in for a coclass without events -- is no evidence, and the
    /// lookup says so rather than picking some implemented interface.
    #[test]
    fn test_coclass_default_source_is_none_without_a_source_interface() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        {
            let not_a_coclass = msforms_typeinfo(&COMMAND_BUTTON_EVENTS);
            assert_eq!(coclass_default_source(&not_a_coclass), None);
        }
        unsafe { CoUninitialize() };
    }

    /// With a coclass that names its default source, the connection points
    /// are never enumerated. That order is the fix for the UserForm extender,
    /// whose three points would otherwise end in the refusal above -- and for
    /// the worksheet extender, whose `EnumConnectionPoints` is `E_NOTIMPL`.
    #[test]
    fn test_source_iid_prefers_the_coclass_default_source_over_enumerating() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        {
            let log = log();
            let source = FakeSource::with_class_info(
                vec![FAKE_SOURCE_A, FAKE_SOURCE_B],
                true,
                Some(msforms_typeinfo(&MSFORMS_COMMAND_BUTTON)),
                log.clone(),
            );
            let iid = source_iid(&source).expect("the coclass typeinfo names the default source");
            assert_eq!(iid, COMMAND_BUTTON_EVENTS);
            assert_eq!(
                log.borrow().enum_calls,
                0,
                "the coclass answered; enumerating would have found two points and refused"
            );
        }
        unsafe { CoUninitialize() };
    }

    /// An `IProvideClassInfo` whose `GetClassInfo` fails is no evidence at
    /// all, so the connection points are enumerated exactly as before.
    #[test]
    fn test_source_iid_falls_back_to_the_connection_points_when_the_coclass_cannot_be_read() {
        let log = log();
        let source = FakeSource::with_class_info(vec![FAKE_SOURCE_A], true, None, log.clone());
        let iid = source_iid(&source).expect("one connection point leaves no choice");
        assert_eq!(iid, FAKE_SOURCE_A);
        assert_eq!(log.borrow().enum_calls, 1, "GetClassInfo failed, so the points were enumerated");
    }

    /// The sink's vtable, reached the way COM reaches it: through the object
    /// pointer's first word. The thunks are called through the vtable rather
    /// than by name on purpose -- a slot wired to the wrong function is
    /// exactly the mistake a hand-written vtable invites, and calling
    /// `EventSink::query_interface` directly would not notice.
    unsafe fn sink_vtable(raw: *mut c_void) -> &'static IDispatch_Vtbl {
        unsafe { &**(raw as *const &'static IDispatch_Vtbl) }
    }

    /// Every branch of the sink's `QueryInterface`, and the refcount
    /// transitions around it. No Excel: Excel's connection point settles for
    /// `IDispatch` and never asks for the source IID, so the arm that answers
    /// it could be deleted -- or made to return a wrong pointer -- without a
    /// single Excel-driven test noticing.
    #[test]
    fn test_the_sink_answers_exactly_three_iids_and_counts_its_references() {
        let (tx, rx) = channel();
        let sink = EventSink::new(tx, 7, FAKE_SOURCE_A);
        let raw = sink.as_raw();
        let vtable = unsafe { sink_vtable(raw) };

        for (label, iid) in [
            ("IUnknown", IUnknown::IID),
            ("IDispatch", IDispatch::IID),
            ("the source dispinterface", FAKE_SOURCE_A),
        ] {
            let mut out: *mut c_void = std::ptr::null_mut();
            let hr = unsafe { (vtable.base__.QueryInterface)(raw, &iid, &mut out) };
            assert!(hr.is_ok(), "QueryInterface({label}) must succeed, got {hr:?}");
            assert_eq!(
                out, raw,
                "QueryInterface({label}) must answer with this object, not something else"
            );
            let remaining = unsafe { (vtable.base__.Release)(raw) };
            assert_eq!(remaining, 1, "QueryInterface({label}) must AddRef exactly once");
        }

        // A junk IID: refused, and `*ppv` nulled -- poisoned first, so a
        // QueryInterface that leaves the caller's pointer alone is caught.
        let junk = GUID::from_u128(0xdeadbeef_dead_beef_dead_beefdeadbeef);
        let mut out: *mut c_void = 0x1 as *mut c_void;
        let hr = unsafe { (vtable.base__.QueryInterface)(raw, &junk, &mut out) };
        assert_eq!(hr, E_NOINTERFACE, "an unknown IID must be refused");
        assert!(out.is_null(), "a refused QueryInterface must null *ppv");

        // Null guards: both are E_POINTER, and neither may touch the refcount.
        let mut out2: *mut c_void = std::ptr::null_mut();
        assert_eq!(
            unsafe { (vtable.base__.QueryInterface)(raw, std::ptr::null(), &mut out2) },
            E_POINTER
        );
        assert_eq!(
            unsafe {
                (vtable.base__.QueryInterface)(raw, &IUnknown::IID, std::ptr::null_mut())
            },
            E_POINTER
        );

        assert_eq!(unsafe { (vtable.base__.AddRef)(raw) }, 2, "AddRef returns the new count");
        assert_eq!(unsafe { (vtable.base__.Release)(raw) }, 1, "Release returns the new count");

        // Still alive: the channel is open because the sink still holds the
        // Sender.
        assert!(matches!(rx.try_recv(), Err(TryRecvError::Empty)), "the sink is still alive");

        drop(sink);
        // The refcount reached 0, the box was freed, and freeing it dropped
        // the Sender. Nothing else can observe a leak from outside.
        assert!(
            matches!(rx.try_recv(), Err(TryRecvError::Disconnected)),
            "releasing the last reference must free the sink and with it the Sender"
        );
    }

    /// Every shape of DISPPARAMS `ordered_args` claims to understand, driven
    /// by hand through the vtable. The live Excel test covers exactly one of
    /// them (the fully-named one), and cannot be made to produce the others.
    #[test]
    fn test_invoke_puts_the_arguments_in_declaration_order() {
        let (tx, rx) = channel();
        let sink = EventSink::new(tx, 42, FAKE_SOURCE_A);
        let raw = sink.as_raw();
        let vtable = unsafe { sink_vtable(raw) };

        // Fire one callback and read back what the sink delivered. VARIANTs
        // are read through this crate's own converter, NOT through
        // `i32::try_from(&VARIANT)`: that routes via `propsys.dll`'s
        // `VariantToInt32`, which this Wine does not implement -- calling it
        // aborts the test binary into winedbg rather than failing.
        let fire = |dispid: i32, values: &[i32], named: &[i32]| -> Vec<i64> {
            let mut args: Vec<VARIANT> = values.iter().map(|v| VARIANT::from(*v)).collect();
            let mut named_ids: Vec<i32> = named.to_vec();
            let params = DISPPARAMS {
                rgvarg: args.as_mut_ptr(),
                rgdispidNamedArgs: if named_ids.is_empty() {
                    std::ptr::null_mut()
                } else {
                    named_ids.as_mut_ptr()
                },
                cArgs: values.len() as u32,
                cNamedArgs: named.len() as u32,
            };
            let hr = unsafe {
                (vtable.Invoke)(
                    raw,
                    dispid,
                    std::ptr::null(),
                    0,
                    DISPATCH_FLAGS(1),
                    &params,
                    std::ptr::null_mut(),
                    std::ptr::null_mut(),
                    std::ptr::null_mut(),
                )
            };
            assert_eq!(hr, S_OK, "Invoke must always succeed: failing back into a source helps nobody");
            let event = rx.try_recv().expect("the event must arrive, and without blocking");
            assert_eq!(event.dispid, dispid);
            assert_eq!(event.handle, 42);
            event
                .args
                .iter()
                .map(|v| {
                    crate::value::variant_to_json(v, |_| panic!("no object arguments here"))
                        .as_i64()
                        .expect("every argument here is an integer")
                })
                .collect()
        };

        // The shape a live Excel actually sends, measured: every argument
        // named, by position, in ascending order.
        assert_eq!(fire(1, &[11, 22], &[0, 1]), vec![11, 22]);

        // The same arguments, named the other way round. This is what proves
        // the sink READS the DISPIDs rather than trusting rgvarg's order --
        // delivering [11, 22] here would be the plausible-looking wrong
        // answer.
        assert_eq!(fire(2, &[11, 22], &[1, 0]), vec![22, 11]);

        // Fully positional, the shape an in-process ATL control's Fire_
        // method sends: last parameter first.
        assert_eq!(fire(3, &[11, 22, 33], &[]), vec![33, 22, 11]);

        // Mixed: one named argument at position 2, and two positionals
        // filling the front, last parameter first.
        assert_eq!(fire(4, &[33, 11, 22], &[2]), vec![22, 11, 33]);

        // Named DISPIDs that are not parameter positions (DISPID_PROPERTYPUT
        // here) mean nothing can be reordered honestly, so nothing is: the
        // arguments arrive exactly as they were sent.
        assert_eq!(fire(5, &[11, 22], &[-3]), vec![11, 22]);
        // ...and so does the same position claimed twice.
        assert_eq!(fire(6, &[11, 22], &[1, 1]), vec![11, 22]);

        // A null DISPPARAMS is not a crash and not an argument.
        let hr = unsafe {
            (vtable.Invoke)(
                raw,
                7,
                std::ptr::null(),
                0,
                DISPATCH_FLAGS(1),
                std::ptr::null(),
                std::ptr::null_mut(),
                std::ptr::null_mut(),
                std::ptr::null_mut(),
            )
        };
        assert_eq!(hr, S_OK);
        let event = rx.try_recv().expect("a callback with no DISPPARAMS is still a callback");
        assert!(event.args.is_empty());
    }

    /// The whole subscription lifecycle without Excel: `advise` puts a sink on
    /// a connection point that asks for the source IID the way ATL does, and
    /// dropping the `Advised` unadvises -- which releases the last reference to
    /// the sink, which drops the `Sender`.
    #[test]
    fn test_advise_answers_the_source_iid_and_dropping_the_subscription_unadvises() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let log = log();
        let source = FakeSource::new(vec![FAKE_SOURCE_A], true, log.clone());
        let (tx, rx) = channel();

        // This fake has no readable typeinfo at all (its `GetTypeInfo` is
        // E_NOTIMPL), so it is also the "cannot tell" half of
        // `refuse_non_dispinterface_typekind`: advising must still SUCCEED
        // here, because an object whose type library says nothing is the
        // ordinary case, not a suspect one.
        let advised = advise(&source, 3, tx).expect("advise on a one-connection-point source");
        assert_eq!(log.borrow().advise_calls, 1);
        assert!(
            log.borrow().sink_answered_source_iid,
            "the sink must answer QueryInterface for the source dispinterface: a connection \
             point is entitled to ask, and ATL's does"
        );
        // No typeinfo on this fake, so no names -- the documented best-effort
        // path, not an error.
        assert!(advised.names.is_empty());
        assert!(matches!(rx.try_recv(), Err(TryRecvError::Empty)), "the sink is alive");

        drop(advised);
        assert_eq!(
            log.borrow().unadvise_calls,
            1,
            "dropping an Advised must unadvise; leaving that to the call site is what this \
             type exists to prevent"
        );
        assert!(
            matches!(rx.try_recv(), Err(TryRecvError::Disconnected)),
            "unadvising must release the source's reference to the sink, freeing it"
        );

        drop(source);
        unsafe { CoUninitialize() };
    }

    /// An explicit `unadvise` surfaces its HRESULT and spends the cookie, so
    /// the `Drop` that follows cannot unadvise a second time with a number
    /// that may since have been handed to somebody else.
    #[test]
    fn test_unadvising_twice_unadvises_once() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let log = log();
        let source = FakeSource::new(vec![FAKE_SOURCE_A], true, log.clone());
        let (tx, _rx) = channel();

        let mut advised = advise(&source, 4, tx).expect("advise");
        advised.unadvise().expect("Unadvise must report success, not swallow it");
        assert_eq!(log.borrow().unadvise_calls, 1);
        advised.unadvise().expect("a second unadvise is a no-op, not an error");
        assert_eq!(log.borrow().unadvise_calls, 1);
        drop(advised);
        assert_eq!(
            log.borrow().unadvise_calls,
            1,
            "Drop must not spend a cookie that has already been handed back"
        );

        drop(source);
        unsafe { CoUninitialize() };
    }

    /// The sink's soundness rests on every callback arriving on one apartment
    /// thread. On an MTA thread COM delivers them concurrently, and `Invoke`
    /// would then be racing a `Sender` that is not `Sync`. `advise` refuses
    /// there rather than documenting it and hoping.
    #[test]
    fn test_advise_refuses_a_multi_threaded_apartment() {
        // On its own thread: CoInitializeEx fixes a thread's apartment for the
        // rest of its life, and every other test's thread must stay STA.
        std::thread::spawn(|| {
            unsafe { CoInitializeEx(None, COINIT_MULTITHREADED).ok().unwrap() };
            let log = log();
            let source = FakeSource::new(vec![FAKE_SOURCE_A], true, log.clone());
            let (tx, _rx) = channel();
            let err = advise(&source, 5, tx).expect_err("an MTA thread must not get a sink");
            assert_eq!(err.code, RPC_E_WRONG_THREAD, "got {err}");
            assert_eq!(
                log.borrow().advise_calls,
                0,
                "it must refuse before any sink exists to be raced"
            );
            drop(source);
            unsafe { CoUninitialize() };
        })
        .join()
        .expect("the MTA thread must not panic");
    }

    /// ...and a thread with no apartment at all is refused too, rather than
    /// having its apartment type guessed.
    #[test]
    fn test_advise_refuses_a_thread_with_no_apartment() {
        std::thread::spawn(|| {
            let log = log();
            let source = FakeSource::new(vec![FAKE_SOURCE_A], true, log.clone());
            let (tx, _rx) = channel();
            let err = advise(&source, 6, tx)
                .expect_err("a thread that never called CoInitializeEx must not get a sink");
            // Wine is entitled to answer either way here: CO_E_NOTINITIALIZED
            // for "no apartment", or an implicit MTA. Both are refusals.
            assert!(
                err.code == CO_E_NOTINITIALIZED || err.code == RPC_E_WRONG_THREAD,
                "expected a refusal naming the apartment, got {err}"
            );
            assert_eq!(log.borrow().advise_calls, 0);
            drop(source);
        })
        .join()
        .expect("the uninitialised thread must not panic");
    }

    /// The memory-safety refusal: a source whose only connection point is
    /// `IPropertyNotifySink` must not be advised on.
    ///
    /// A Range is exactly this object (measured -- see
    /// [`UNIMPLEMENTABLE_SOURCE_IIDS`]), and `source_iid` is right to take the
    /// only connection point on offer; it is `advise` that has to refuse,
    /// because the sink would then be answering `QI(IPropertyNotifySink)` with
    /// an `IDispatch` vtable whose slot 3 takes a `*mut u32` where the caller
    /// passes a DISPID by value. No Excel needed to pin it: the IID is the
    /// whole of the decision.
    #[test]
    fn test_advise_refuses_a_property_notify_sink_connection_point() {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
        let log = log();
        let source = FakeSource::new(vec![IPropertyNotifySink::IID], true, log.clone());
        let (tx, rx) = channel();

        let err = advise(&source, 8, tx)
            .expect_err("IPropertyNotifySink is a vtable interface: this sink cannot implement it");
        assert_eq!(err.code, E_NOINTERFACE, "got {err}");
        assert!(
            err.to_string().contains("IPropertyNotifySink"),
            "the refusal must name the interface it refused, or nobody can act on it: {err}"
        );
        assert_eq!(
            log.borrow().advise_calls,
            0,
            "it must refuse before the source is ever handed a sink"
        );
        assert!(
            matches!(rx.try_recv(), Err(TryRecvError::Disconnected)),
            "a refused advise must not leave a sink holding the Sender"
        );

        // The same fake with an ordinary source dispinterface still works, so
        // the refusal above is a decision about the IID and not the fake
        // failing to advise at all.
        let (tx2, _rx2) = channel();
        let ok = FakeSource::new(vec![FAKE_SOURCE_A], true, log.clone());
        drop(advise(&ok, 9, tx2).expect("an ordinary source dispinterface is still advised on"));

        drop(ok);
        drop(source);
        unsafe { CoUninitialize() };
    }

    /// The typekind rule's PERMISSIVE form, which is the half that is easy to
    /// get backwards. Driven directly rather than through a type library,
    /// because the decision *is* the policy: "cannot tell" and "told, and it
    /// is not a dispinterface" must be different outcomes.
    ///
    /// Inverting the `None` arm would refuse every source with no readable
    /// typeinfo -- which is the documented, required path for an object whose
    /// events can only be named `DISPID_<n>` -- and no Excel-driven test would
    /// notice, because Excel's Application does have a type library.
    #[test]
    fn test_the_typekind_rule_refuses_only_what_it_could_actually_read() {
        use windows::Win32::System::Com::{TKIND_COCLASS, TKIND_INTERFACE};

        // Told, and it is not a dispinterface: refuse, naming the typekind.
        let err = refuse_non_dispinterface_typekind(Some(TKIND_INTERFACE), &FAKE_SOURCE_A)
            .expect_err("a TKIND_INTERFACE source is a vtable interface, not a dispinterface");
        assert_eq!(err.code, E_NOINTERFACE, "got {err}");
        assert!(
            err.to_string().contains(&TKIND_INTERFACE.0.to_string()),
            "the refusal must say what it was told: {err}"
        );
        assert!(
            refuse_non_dispinterface_typekind(Some(TKIND_COCLASS), &FAKE_SOURCE_A).is_err(),
            "any typekind that is not TKIND_DISPATCH is refused, not just TKIND_INTERFACE"
        );

        // Told, and it IS a dispinterface: proceed. This is every real Excel
        // source interface.
        assert!(
            refuse_non_dispinterface_typekind(Some(TKIND_DISPATCH), &FAKE_SOURCE_A).is_ok(),
            "a dispinterface source is exactly what this sink implements"
        );

        // Cannot tell: PROCEED. An object with no readable typeinfo must keep
        // delivering events -- named `DISPID_<n>` -- rather than being refused
        // on a suspicion the type library never voiced.
        assert!(
            refuse_non_dispinterface_typekind(None, &FAKE_SOURCE_A).is_ok(),
            "an unreadable typeinfo is not evidence: 'cannot tell' must not refuse, or every \
             source without a type library stops working"
        );
    }
}
