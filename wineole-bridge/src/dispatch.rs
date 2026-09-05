use windows::core::{Interface, BSTR, GUID, HRESULT, PCWSTR};
use windows::Win32::Foundation::{DISP_E_EXCEPTION, E_INVALIDARG};
use windows::Win32::System::Com::{
    CoCreateInstance, CLSIDFromProgID, IDispatch, DISPPARAMS, EXCEPINFO, DISPATCH_FLAGS,
    DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT, CLSCTX_LOCAL_SERVER,
};
use windows::Win32::System::Com::{ITypeInfo, ITypeLib, TYPEATTR, VARDESC, TKIND_ENUM};
use windows::Win32::System::Ole::GetActiveObject;
use windows::Win32::Foundation::E_FAIL;
use windows::core::VARIANT;

pub(crate) const LOCALE_USER_DEFAULT: u32 = 0x0400;
const DISPID_PROPERTYPUT: i32 = -3;
const DISPID_VALUE: i32 = 0;

/// A COM failure: the HRESULT, plus whatever human-readable detail we managed
/// to recover for it.
///
/// `windows::core::Error` cannot carry this itself in this environment:
///
/// - `Error::message()` on a bare HRESULT is routinely **empty**, because
///   Wine's `FormatMessage` cannot resolve most automation HRESULTs. That is
///   how a bad ProgID used to reach the client as `WIN32OLERuntimeError: ""`.
/// - `Error::new(code, text)` — the obvious way to attach the text we *do*
///   have — stores it via `RoOriginateError`, a WinRT entry point that does
///   nothing here, so the text is silently discarded and `message()` still
///   comes back empty (verified against this Wine build).
///
/// So keep the code and the text side by side, and always render the code:
/// even with no message at all, the caller learns *which* HRESULT failed.
#[derive(Debug, Clone)]
pub struct ComError {
    pub code: HRESULT,
    pub detail: Option<String>,
}

pub type ComResult<T> = std::result::Result<T, ComError>;

impl ComError {
    pub fn new(code: HRESULT, detail: impl Into<String>) -> Self {
        let detail = detail.into();
        ComError {
            code,
            detail: if detail.trim().is_empty() { None } else { Some(detail) },
        }
    }
}

impl From<windows::core::Error> for ComError {
    fn from(e: windows::core::Error) -> Self {
        ComError::new(e.code(), e.message())
    }
}

impl std::fmt::Display for ComError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        let code = self.code.0 as u32;
        match &self.detail {
            Some(detail) => write!(f, "{} (0x{:08X})", detail, code),
            None => write!(f, "COM error (0x{:08X})", code),
        }
    }
}

fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

fn get_dispids(disp: &IDispatch, names: &[&str]) -> ComResult<Vec<i32>> {
    let wides: Vec<Vec<u16>> = names.iter().map(|n| to_wide(n)).collect();
    let ptrs: Vec<PCWSTR> = wides.iter().map(|w| PCWSTR(w.as_ptr())).collect();
    let mut dispids = vec![0i32; names.len()];
    unsafe {
        disp.GetIDsOfNames(
            &GUID::zeroed(),
            ptrs.as_ptr(),
            names.len() as u32,
            LOCALE_USER_DEFAULT,
            dispids.as_mut_ptr(),
        )?;
    }
    Ok(dispids)
}

fn raw_invoke(
    disp: &IDispatch,
    dispid: i32,
    flags: DISPATCH_FLAGS,
    mut positional: Vec<VARIANT>,
    named: Vec<(i32, VARIANT)>,
) -> ComResult<VARIANT> {
    // OUTBOUND arguments are reversed; INBOUND ones, in sink.rs's `Invoke`,
    // are not. That looks like the project holding two opposite beliefs, so:
    // both are measured, and they are a pair rather than a contradiction.
    //
    // Out: DISPPARAMS::rgvarg is documented last-to-first, and Excel's own
    // Invoke reads it that way. Measured against a live Excel by
    // `test_positional_arguments_reach_excel_in_order` below -- Cells(2, 3)
    // must be $C$2, and without this line it is $B$3.
    //
    // In: the callback does not come from a caller like this one at all, it
    // comes out of Wine's RPC stub -- and measured, that stub sends every
    // event argument as a NAMED one, identified by its parameter position
    // (SheetChange arrives with cArgs=2, cNamedArgs=2, rgdispidNamedArgs=[0,
    // 1]). So no positional reversal applies to it, and `sink::ordered_args`
    // -- which is this function read backwards, and says so -- places those
    // arguments by the positions their DISPIDs name.
    positional.reverse();
    let mut rgvarg: Vec<VARIANT> = Vec::new();
    let mut rgdispid_named: Vec<i32> = Vec::new();
    for (id, v) in named {
        rgdispid_named.push(id);
        rgvarg.push(v);
    }
    rgvarg.extend(positional);

    let params = DISPPARAMS {
        rgvarg: if rgvarg.is_empty() { std::ptr::null_mut() } else { rgvarg.as_mut_ptr() },
        rgdispidNamedArgs: if rgdispid_named.is_empty() { std::ptr::null_mut() } else { rgdispid_named.as_mut_ptr() },
        cArgs: rgvarg.len() as u32,
        cNamedArgs: rgdispid_named.len() as u32,
    };

    let mut result = VARIANT::default();
    let mut excepinfo = EXCEPINFO::default();
    let mut arg_err: u32 = 0;
    let invoked = unsafe {
        disp.Invoke(
            dispid,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            flags,
            &params,
            Some(&mut result),
            Some(&mut excepinfo),
            Some(&mut arg_err),
        )
    };

    // Per the IDispatch::Invoke contract the callee fills EXCEPINFO's three
    // BSTRs whenever it returns DISP_E_EXCEPTION, and is permitted to touch
    // the structure on other calls too; ownership of those strings transfers
    // to us either way. They are `ManuallyDrop<BSTR>`, so nothing frees them
    // unless we do — take all three unconditionally (which frees them when the
    // owned `BSTR`s go out of scope at the end of this function) rather than
    // only on the exception path, where a non-exception call that still
    // allocated would leak.
    let source = unsafe { take_excepinfo_bstr(&mut excepinfo.bstrSource) };
    let description = unsafe { take_excepinfo_bstr(&mut excepinfo.bstrDescription) };
    let _help_file = unsafe { take_excepinfo_bstr(&mut excepinfo.bstrHelpFile) };

    match invoked {
        Ok(()) => Ok(result),
        Err(e) if e.code() == DISP_E_EXCEPTION => {
            // DISP_E_EXCEPTION on its own says only "the callee raised
            // something"; the automation server's own description — e.g.
            // Excel's "Application-defined or object-defined error" — is the
            // real error, and lives in EXCEPINFO. Its `scode`/`wCode` is the
            // server-specific code (0x800A03EC for that Excel error), which is
            // far more useful here than repeating DISP_E_EXCEPTION.
            let inner = if excepinfo.scode != 0 {
                HRESULT(excepinfo.scode)
            } else if excepinfo.wCode != 0 {
                HRESULT(excepinfo.wCode as i32)
            } else {
                DISP_E_EXCEPTION
            };
            match description.or(source) {
                Some(detail) => Err(ComError::new(inner, detail)),
                None => Err(ComError::from(e)),
            }
        }
        Err(e) => Err(ComError::from(e)),
    }
}

/// Take ownership of an `EXCEPINFO` BSTR field, returning its text (if any).
///
/// The returned `BSTR` is dropped by the caller's scope, which is what calls
/// `SysFreeString` — `ManuallyDrop` exists on these fields precisely because
/// the struct itself won't do it.
///
/// # Safety
/// The field must not be read again after this call (it is left holding a
/// bitwise copy of a freed pointer).
unsafe fn take_excepinfo_bstr(field: &mut std::mem::ManuallyDrop<BSTR>) -> Option<String> {
    let bstr = std::mem::ManuallyDrop::take(field);
    if bstr.is_empty() {
        None
    } else {
        Some(bstr.to_string())
    }
}

pub fn create_instance(progid: &str) -> ComResult<IDispatch> {
    let wide = to_wide(progid);
    let clsid = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr()))? };
    unsafe { Ok(CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)?) }
}

pub fn get_active_object(progid: &str) -> ComResult<IDispatch> {
    let wide = to_wide(progid);
    let clsid = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr()))? };
    // GetActiveObject (declared in windows::Win32::System::Ole, not ::Com, in
    // this crate version) returns HRESULT via an out-param (`ppunk: *mut
    // Option<IUnknown>`), not as the function's `Result<T>` payload — unlike
    // the brief's assumed `GetActiveObject(&clsid, None) -> Result<IUnknown>`
    // shape. See windows-0.58.0/src/Windows/Win32/System/Ole/mod.rs:104.
    let mut punk: Option<windows::core::IUnknown> = None;
    unsafe {
        GetActiveObject(&clsid, None, &mut punk)?;
    }
    let unknown = punk.ok_or_else(|| {
        ComError::new(
            windows::Win32::Foundation::E_FAIL,
            format!("GetActiveObject({}) succeeded but returned no object", progid),
        )
    })?;
    Ok(unknown.cast()?)
}

/// GetActiveObject-or-CreateObject: attach to an already-running instance if
/// one exists, otherwise spawn a new one. Returns the object plus whether it
/// was freshly created (`true`) or an existing instance was attached to
/// (`false`).
///
/// Any `get_active_object` failure OTHER than "nothing is currently running"
/// (`MK_E_UNAVAILABLE`) is a real error and must propagate as-is -- e.g. a
/// malformed ProgID must fail the same way a bare `connect` would, not
/// silently fall through to `create_instance` and mask the problem.
pub fn connect_or_create(progid: &str) -> ComResult<(IDispatch, bool)> {
    match get_active_object(progid) {
        Ok(disp) => Ok((disp, false)),
        Err(e) if e.code == windows::Win32::Foundation::MK_E_UNAVAILABLE => {
            let disp = create_instance(progid)?;
            Ok((disp, true))
        }
        Err(e) => Err(e),
    }
}

/// Invoke a member by name against `disp`.
///
/// - `name == ""` invokes the default member (`DISPID_VALUE`) as
///   `DISPATCH_METHOD | DISPATCH_PROPERTYGET` (e.g. a collection indexer).
/// - `name` ending in `"="` is a property-set: the bare name (without the
///   `=`) is resolved and the single value in `positional[0]` is passed as
///   the required `DISPID_PROPERTYPUT` named argument (`named` is ignored).
/// - otherwise, `name` (plus every key of `named`) is resolved via one
///   `GetIDsOfNames` call and invoked as `DISPATCH_METHOD | DISPATCH_PROPERTYGET`.
pub fn invoke_member(
    disp: &IDispatch,
    name: &str,
    positional: Vec<VARIANT>,
    named: Vec<(String, VARIANT)>,
) -> ComResult<VARIANT> {
    if name.is_empty() {
        return raw_invoke(disp, DISPID_VALUE, DISPATCH_METHOD | DISPATCH_PROPERTYGET, positional, vec![]);
    }
    if let Some(bare) = name.strip_suffix('=') {
        let ids = get_dispids(disp, &[bare])?;
        let value = positional.into_iter().next().ok_or_else(|| {
            ComError::new(E_INVALIDARG, format!("property-set {} requires a value", bare))
        })?;
        return raw_invoke(disp, ids[0], DISPATCH_PROPERTYPUT, vec![], vec![(DISPID_PROPERTYPUT, value)]);
    }

    let mut names: Vec<&str> = vec![name];
    names.extend(named.iter().map(|(k, _)| k.as_str()));
    let ids = get_dispids(disp, &names)?;
    let named_pairs: Vec<(i32, VARIANT)> = named
        .into_iter()
        .zip(ids[1..].iter())
        .map(|((_, v), id)| (*id, v))
        .collect();
    raw_invoke(disp, ids[0], DISPATCH_METHOD | DISPATCH_PROPERTYGET, positional, named_pairs)
}

/// Enumerate every constant in `disp`'s type library — the equivalent of
/// `WIN32OLE.const_load`. Walks `GetTypeInfo(0)` → `GetContainingTypeLib` →
/// every `TKIND_ENUM` type info in that library → `GetVarDesc` for each of
/// its members. Returns raw `(name, VARIANT)` pairs; the caller (session.rs)
/// converts to JSON, since this module stays JSON-agnostic.
pub fn const_load(disp: &IDispatch) -> ComResult<Vec<(String, VARIANT)>> {
    unsafe {
        let type_info: ITypeInfo = disp.GetTypeInfo(0, LOCALE_USER_DEFAULT)?;

        let mut type_lib: Option<ITypeLib> = None;
        let mut lib_index: u32 = 0;
        type_info.GetContainingTypeLib(&mut type_lib, &mut lib_index)?;
        let type_lib = type_lib.ok_or_else(|| {
            ComError::new(E_FAIL, "GetContainingTypeLib returned no type library")
        })?;

        let mut constants = Vec::new();
        let info_count = type_lib.GetTypeInfoCount();
        for i in 0..info_count {
            let ti: ITypeInfo = match type_lib.GetTypeInfo(i) {
                Ok(t) => t,
                Err(_) => continue, // a handful of type infos can be unreadable; skip, don't abort the whole walk
            };

            let attr: *mut TYPEATTR = match ti.GetTypeAttr() {
                Ok(a) => a,
                Err(_) => continue,
            };

            if (*attr).typekind != TKIND_ENUM {
                ti.ReleaseTypeAttr(attr);
                continue;
            }

            let var_count = (*attr).cVars;
            for v in 0..var_count {
                if let Ok(vardesc) = ti.GetVarDesc(v as u32) {
                    if let Some(pair) = read_enum_constant(&ti, vardesc) {
                        constants.push(pair);
                    }
                    ti.ReleaseVarDesc(vardesc);
                }
            }

            ti.ReleaseTypeAttr(attr);
        }

        Ok(constants)
    }
}

/// Read one enum member's name and value out of a `VARDESC`. Returns `None`
/// (rather than failing the whole `const_load` call) if this one member's
/// name can't be resolved — a malformed single entry shouldn't hide every
/// other constant in the library.
///
/// # Safety
/// `vardesc` must be a valid pointer obtained from `ti.GetVarDesc` and not
/// yet released.
unsafe fn read_enum_constant(ti: &ITypeInfo, vardesc: *mut VARDESC) -> Option<(String, VARIANT)> {
    let memid = (*vardesc).memid;
    let mut names: [windows::core::BSTR; 1] = [windows::core::BSTR::new()];
    let mut name_count: u32 = 0;
    if ti.GetNames(memid, &mut names, &mut name_count).is_err() || name_count == 0 {
        return None;
    }
    let name = names[0].to_string();

    // VARDESC_0::lpvarValue is only meaningful when varkind == VAR_CONST,
    // which every member of a TKIND_ENUM type info is by definition — clone
    // (not move) the VARIANT it points to, since VARDESC still owns it until
    // ReleaseVarDesc runs.
    let value_ptr = (*vardesc).Anonymous.lpvarValue;
    if value_ptr.is_null() {
        return None;
    }
    let value = (*value_ptr).clone();

    Some((name, value))
}

/// Serialises every test in this crate that creates or connects to
/// `Excel.Application` — they live in `dispatch`, `session`, `server` and
/// `sink`, which is why this sits at module scope rather than inside
/// `mod tests` where its Word counterpart (`WORD_ROT_LOCK`) lives.
///
/// **This is a symptomatic fix. The cause is not understood.** Read the
/// whole of this comment before changing anything here.
///
/// ## What is observed
///
/// `cargo test` runs the whole suite as concurrent threads in one process.
/// Past some concurrency threshold, `create_instance("Excel.Application")`
/// starts failing with `E_NOINTERFACE` (0x80004002) while Wine logs
///
/// ```text
/// err:ole:com_get_class_object no class object
///   {00024500-0000-0000-c000-000000000046} could be created for context 0x4
/// ```
///
/// (that CLSID is `Excel.Application`; context 0x4 is `CLSCTX_LOCAL_SERVER`).
/// Measured, `cargo test --target x86_64-pc-windows-gnu`:
///
/// | tree | result |
/// | --- | --- |
/// | this suite, parallel | 64 passed / 2 failed — twice, the same two |
/// | the suite as it stood before this fix, untouched, parallel | 64 passed / 1 failed — already failing |
/// | that same earlier suite plus ONE trivial extra Excel test (create, `Quit`, done — no sink, no events) | 64 passed / **2** failed — the same two |
/// | `--test-threads=1` | 66 passed / 0 failed — twice |
///
/// The third row is the load-bearing measurement. A test that does nothing
/// but create Excel and quit it moved the failure count from one to two, and
/// the tests that broke were not the new one. So the trigger is *how many
/// tests drive Excel at once*, not what any individual test does — which is
/// why the lock is taken by all of them and not just by the two that happen
/// to fail today.
///
/// ## What the cause is not
///
/// Four hypotheses were measured, and **all four are disproven**. They are
/// recorded here so the next person does not spend another afternoon
/// re-measuring them:
///
/// 1. *"The tests share one Excel instance, and one test's cleanup breaks
///    another's."* No. `CoCreateInstance("Excel.Application")` spawns a
///    **new** EXCEL.EXE every time; measured 1 → 2 → 3 → 4 live processes on
///    successive creates.
/// 2. *"`Quit` bypasses COM refcounting and kills other clients'
///    references."* No. With client A and client B both holding references,
///    `A.Quit` leaves the process alive and B fully usable.
/// 3. *"There is a limit of about three concurrent Excel processes."* No.
///    Eight simultaneous creates from eight threads all succeeded.
/// 4. *"A create fails when it races another Excel's teardown."* No. Twelve
///    rounds of create-racing-a-`Quit`: twelve successes.
///
/// Hypothesis 4 is why this lock does **not** carry a
/// `wait_for_excel_gone_from_rot` counterpart to the Word lock's drain: a
/// create racing a teardown was measured not to fail, and the five green
/// full-suite runs that shipped this change agree (`pgrep -cf 'EXCE[L].EXE'`
/// settled at 0 after every one of them). `WORD_ROT_LOCK` needs its
/// drain for a different and *understood* reason — a Word test asserts that
/// nothing matching `Word.Application` is running, which is only true once
/// WINWORD.EXE has actually left the Running Object Table. No Excel test
/// asserts the absence of Excel, so there is no equivalent invariant to
/// uphold here. If the failure ever returns, adding that drain is the first
/// thing to try.
///
/// ## Deliberately separate from `WORD_ROT_LOCK`
///
/// Two locks, not one, decided on the measurement rather than on taste. The
/// question is whether Word and Publisher tests running concurrently with
/// Excel tests can also trigger this: with the two locks kept separate they
/// do exactly that on every run, and five consecutive suites came back green
/// against a baseline that failed the same way every time. So they do not
/// interfere, and one lock would be buying nothing.
///
/// They are also answers to different questions. `WORD_ROT_LOCK` upholds a
/// specific, understood invariant (a Word test asserts nothing matching
/// `Word.Application` is running) and is paired with a Word-specific ROT
/// drain; this one suppresses an unexplained creation failure and has no
/// drain. Folding them together would put three more tests on this chain,
/// and would leave the Word lock's carefully-argued doc comment describing
/// only half of what it guards. If Excel ever *does* start failing while
/// only a Word or Publisher test is in flight, merging them is the obvious
/// next move — the evidence just does not point there today.
///
/// ## Not `--test-threads=1`
///
/// Serialising the *whole* suite is forbidden here in writing
/// (`.superpowers/sdd/task-1.5-brief.md`): the tests that do not touch
/// Office are meant to keep running in parallel, and a previous attempt at
/// `--test-threads=1` (`5ed1c24`) was reverted for that reason (`b3e522e`).
/// Only the Office-driving tests take a lock; everything else still runs in
/// parallel. Measured on the development host: green in 18.9–19.1s with this
/// lock, against 22.8s for the same suite under `--test-threads=1`. (The
/// *failing* parallel baseline took 30.5s — this fix is not a slowdown; a
/// pile-up of Excel launches that ends in `E_NOINTERFACE` costs more than
/// starting them one at a time.)
#[cfg(test)]
static EXCEL_TEST_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

/// Take [`EXCEL_TEST_LOCK`] for the rest of the caller's scope. Every test
/// that creates or connects to `Excel.Application` must open with
///
/// ```ignore
/// let _guard = crate::dispatch::lock_excel_for_test();
/// ```
///
/// — bound to a named `_guard`, never to a bare `_`, which would drop the
/// guard immediately and serialise nothing.
///
/// Recovering a poisoned lock (rather than unwrapping) is deliberate: one
/// panicking Excel test must not cascade into a failure in every other Excel
/// test, which would bury the real one in noise. The mutex guards no data,
/// so there is no invariant a panic could have left broken.
#[cfg(test)]
pub(crate) fn lock_excel_for_test() -> std::sync::MutexGuard<'static, ()> {
    EXCEL_TEST_LOCK.lock().unwrap_or_else(|e| e.into_inner())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::value::{dispatch_from_variant, variant_from_dispatch, variant_to_json};
    use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};

    // Serializes the two Word.Application-based tests below against each
    // other. One of them (`test_connect_or_create_creates_when_word_is_not_running`)
    // necessarily makes a Word.Application exist (however briefly) via its
    // `create_instance` fallback; the other asserts nothing matching
    // Word.Application is running. cargo test runs tests as concurrent
    // threads within one process by default, so without this lock the two
    // can interleave and produce a false failure depending on scheduling.
    // Poisoning is not a concern here: if one test panics while holding the
    // lock, letting the other fail loudly (rather than silently racing) is
    // the correct outcome, so recovering the guts of a poisoned lock is
    // intentional.
    //
    // The lock alone is NOT sufficient, though: it only serializes the two
    // tests' *execution*, not Word's actual removal from the Running Object
    // Table. `Quit` is asynchronous -- WINWORD.EXE is still alive (and still
    // registered in the ROT) at the moment `invoke_member(&word, "Quit",
    // ...)` returns, and the creating test's own `word: IDispatch` handle
    // keeps a reference to it alive until it is dropped. If the lock is
    // released before Word has actually finished exiting, the "nothing is
    // running" test can acquire the lock, call `get_active_object`, and find
    // Word still there -- an intermittent, scheduling-dependent failure.
    // `wait_for_word_gone_from_rot` below closes that gap by polling
    // `get_active_object` until it genuinely fails, before the lock guard is
    // dropped.
    //
    // This lock covers Word only. The suite's Excel-driving tests are
    // serialized separately, by `super::EXCEL_TEST_LOCK` / the
    // `super::lock_excel_for_test()` helper at module scope above, for an
    // unrelated and (unlike this one) *unexplained* reason. See that
    // static's doc comment for why the two are kept apart.
    static WORD_ROT_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    /// `HRESULT_FROM_WIN32(RPC_S_CALL_FAILED)`. The `windows` crate has no
    /// named constant for this one. Empirically, under Wine, the first poll
    /// of `GetActiveObject("Word.Application")` immediately after `Quit` +
    /// dropping the last live handle reliably returns this -- Word's RPC
    /// endpoint is in the middle of being torn down -- and then the very
    /// next poll (50ms later) returns `MK_E_UNAVAILABLE` as expected. It is
    /// a transient of normal shutdown, not a genuine failure; do not
    /// "simplify" this back into the general error arm below, or
    /// `wait_for_word_gone_from_rot` will fail deterministically on every
    /// run under Wine.
    const RPC_CALL_FAILED: windows::core::HRESULT = windows::core::HRESULT(0x800706BE_u32 as _);

    /// Poll `get_active_object("Word.Application")` until it stops finding an
    /// instance (i.e. until Word has actually left the ROT), rather than
    /// trusting that a call to `Quit` has taken effect by the time it returns.
    ///
    /// Panics if Word is still visible after `timeout`, so a real regression
    /// here fails loudly and immediately instead of silently destabilizing
    /// whichever test happens to run next. Also panics immediately -- without
    /// waiting out the timeout -- on any error other than `MK_E_UNAVAILABLE`
    /// or the `RPC_CALL_FAILED` shutdown transient, since no amount of
    /// waiting will fix a genuine COM failure and misreporting it as "Word
    /// was still visible" would send a future debugger chasing a race that
    /// never happened.
    fn wait_for_word_gone_from_rot(timeout: std::time::Duration, poll_interval: std::time::Duration) {
        let deadline = std::time::Instant::now() + timeout;
        loop {
            let last_observation = match get_active_object("Word.Application") {
                Err(e) if e.code == windows::Win32::Foundation::MK_E_UNAVAILABLE => return,
                Err(e) if e.code == RPC_CALL_FAILED => {
                    // Word's RPC endpoint tearing down mid-shutdown -- treat
                    // exactly like "still registered" and keep polling.
                    format!(
                        "GetActiveObject returned RPC_S_CALL_FAILED (0x800706BE) -- \
                         Word's RPC endpoint is tearing down: {:?}",
                        e
                    )
                }
                Err(e) => {
                    // Anything else is a genuine, unexpected COM failure --
                    // waiting cannot fix it, so fail fast with the real
                    // error instead of spinning for the full timeout and
                    // then blaming a race that never occurred.
                    panic!(
                        "unexpected error from GetActiveObject while polling for Word to exit: {:?}",
                        e
                    );
                }
                Ok(_) => "GetActiveObject still returned a live Word.Application instance".to_string(),
            };

            if std::time::Instant::now() >= deadline {
                panic!(
                    "Word.Application never left the Running Object Table after Quit \
                     (timeout {:?}); last observation: {}",
                    timeout, last_observation
                );
            }
            std::thread::sleep(poll_interval);
        }
    }

    #[test]
    fn test_full_round_trip_against_real_excel() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        // This test additionally *needs* to be the only Excel client running:
        // its `get_active_object("Excel.Application")` below would otherwise
        // be free to attach to some other test's instance.
        let _guard = lock_excel_for_test();

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");

        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).expect("Visible=false");
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]).expect("DisplayAlerts=false");

        let workbooks_v = invoke_member(&xl, "Workbooks", vec![], vec![]).expect("Workbooks");
        let workbooks = dispatch_from_variant(&workbooks_v);
        invoke_member(&workbooks, "Add", vec![], vec![]).expect("Workbooks.Add");

        let sheets_v = invoke_member(&xl, "Worksheets", vec![], vec![]).expect("Worksheets");
        let sheets = dispatch_from_variant(&sheets_v);

        let first_sheet_v = invoke_member(&sheets, "", vec![VARIANT::from(1i32)], vec![]).expect("Worksheets(1)");
        let first_sheet = dispatch_from_variant(&first_sheet_v);

        invoke_member(
            &sheets,
            "Add",
            vec![],
            vec![("After".to_string(), variant_from_dispatch(first_sheet))],
        )
        .expect("Worksheets.Add(After: ...)");

        let sheets2_v = invoke_member(&xl, "Worksheets", vec![], vec![]).unwrap();
        let sheets2 = dispatch_from_variant(&sheets2_v);
        let count_v = invoke_member(&sheets2, "Count", vec![], vec![]).unwrap();
        let count_json = variant_to_json(&count_v, |_| panic!("Count should not be an object"));
        assert_eq!(count_json, serde_json::json!(4));

        // get_active_object (the `connect` RPC path) is not covered by anything
        // above — Excel registers itself in the Running Object Table once
        // instantiated via automation, so a second, independent handle obtained
        // via GetActiveObject should observe the same live instance.
        let xl2 = get_active_object("Excel.Application").expect("GetActiveObject(Excel.Application)");
        let version_v = invoke_member(&xl2, "Version", vec![], vec![]).expect("Version via connected instance");
        let version_json = variant_to_json(&version_v, |_| panic!("Version should not be an object"));
        assert_eq!(version_json, serde_json::json!("11.0"));

        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");

        unsafe {
            CoUninitialize();
        }
    }

    #[test]
    fn test_com_errors_carry_a_usable_message() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        let _guard = lock_excel_for_test();

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");
        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).expect("Visible=false");
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![])
            .expect("DisplayAlerts=false");

        // 1. A member that does not exist fails at GetIDsOfNames with
        //    DISP_E_UNKNOWNNAME (0x80020006) — the HRESULT is what identifies
        //    the failure, since Wine's FormatMessage often yields "".
        let unknown = invoke_member(&xl, "NoSuchMemberAtAll", vec![], vec![]).unwrap_err();
        assert_eq!(
            unknown.code.0 as u32,
            0x8002_0006,
            "expected DISP_E_UNKNOWNNAME, got {:?}",
            unknown
        );
        assert!(
            unknown.to_string().contains("0x80020006"),
            "the HRESULT must survive into the wire message, got {:?}",
            unknown.to_string()
        );

        // 2. A real member called in a way the server rejects raises a COM
        //    exception; the description in EXCEPINFO is the only human-readable
        //    part of it, and it used to be discarded entirely.
        let workbooks_v = invoke_member(&xl, "Workbooks", vec![], vec![]).expect("Workbooks");
        let workbooks = dispatch_from_variant(&workbooks_v);
        let open_err = invoke_member(
            &workbooks,
            "Open",
            vec![VARIANT::from("C:\\definitely-does-not-exist-wineole.xls")],
            vec![],
        )
        .unwrap_err();
        eprintln!("Workbooks.Open(bad path) -> {}", open_err);
        let rendered = open_err.to_string();
        assert!(
            open_err.detail.is_some(),
            "EXCEPINFO's description must reach the caller; got {:?}",
            rendered
        );
        assert!(
            rendered.contains("(0x"),
            "the wire message must name an HRESULT, got {:?}",
            rendered
        );
        // Excel reports its automation failures as scode 0x800A03EC
        // ("Application-defined or object-defined error", VBA error 1004),
        // which is far more specific than the DISP_E_EXCEPTION wrapper.
        assert_eq!(
            open_err.code.0 as u32,
            0x800A_03EC,
            "expected Excel's own scode, got {:?}",
            rendered
        );

        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");

        unsafe {
            CoUninitialize();
        }
    }

    #[test]
    fn test_const_load_returns_real_excel_constants() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        let _guard = lock_excel_for_test();

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");
        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).expect("Visible=false");

        let constants = const_load(&xl).expect("const_load should succeed against Excel.Application");
        assert!(
            !constants.is_empty(),
            "const_load returned no constants at all — the type library walk found nothing"
        );

        let as_json: std::collections::HashMap<String, i64> = constants
            .iter()
            .filter_map(|(name, v)| {
                let json = crate::value::variant_to_json(v, |_| {
                    panic!("an enum constant should never be an object reference")
                });
                json.as_i64().map(|n| (name.clone(), n))
            })
            .collect();

        // xlUp and xlDown are well-known, stable XlDirection constants from
        // Excel's own type library — their values are documented and have not
        // changed across Excel versions.
        assert_eq!(as_json.get("xlUp").copied(), Some(-4162), "xlUp missing or wrong: {:?}", as_json.get("xlUp"));
        assert_eq!(as_json.get("xlDown").copied(), Some(-4121), "xlDown missing or wrong: {:?}", as_json.get("xlDown"));

        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");
        unsafe {
            CoUninitialize();
        }
    }

    #[test]
    fn test_connect_or_create_creates_when_word_is_not_running() {
        // Word.Application, not Excel.Application -- see this task's header
        // note on why. Held for the whole test: this briefly makes a real
        // Word.Application exist, which would otherwise race
        // test_get_active_object_fails_with_mk_e_unavailable_when_word_is_not_running's
        // "nothing is running" assertion below.
        let _guard = WORD_ROT_LOCK.lock().unwrap_or_else(|e| e.into_inner());

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let (word, created) = connect_or_create("Word.Application")
            .expect("connect_or_create should create a fresh Word instance");
        assert!(created, "connect_or_create must report created=true when nothing was running");

        invoke_member(&word, "Quit", vec![], vec![]).expect("Quit");

        // `word` itself is a live reference to the Word.Application COM
        // object; as long as it is held, Word cannot finish exiting. Drop it
        // before polling, and poll (rather than assuming `Quit` has already
        // taken effect) since `Quit` is asynchronous -- see WORD_ROT_LOCK's
        // doc comment above.
        drop(word);
        wait_for_word_gone_from_rot(
            std::time::Duration::from_secs(10),
            std::time::Duration::from_millis(50),
        );

        unsafe {
            CoUninitialize();
        }
    }

    #[test]
    fn test_connect_or_create_attaches_when_publisher_is_already_running() {
        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        // Publisher.Application in this Wine/Office install does not expose
        // a working "Visible" property (DISP_E_UNKNOWNNAME) -- empirically
        // verified while diagnosing this fix. "ScreenUpdating" is a mutable
        // boolean property it does support, so use that instead to prove
        // shared live-instance identity below.
        let publisher = create_instance("Publisher.Application").expect("CreateObject(Publisher.Application)");
        invoke_member(&publisher, "ScreenUpdating=", vec![VARIANT::from(false)], vec![]).expect("ScreenUpdating=false");

        let (publisher2, created) = connect_or_create("Publisher.Application")
            .expect("connect_or_create should attach to the running instance");
        assert!(!created, "connect_or_create must report created=false when something was already running");

        // Prove it's genuinely the same live instance, not a coincidentally
        // similar second one: `publisher` was set to ScreenUpdating=false
        // above (a non-default state -- ScreenUpdating defaults to true on
        // Office Application objects), so an independent second instance
        // would read back true here and be caught; only the same live
        // instance reads back false.
        let screen_updating_v = invoke_member(&publisher2, "ScreenUpdating", vec![], vec![]).expect("ScreenUpdating via publisher2");
        let screen_updating_json = variant_to_json(&screen_updating_v, |_| panic!("ScreenUpdating should not be an object"));
        assert_eq!(screen_updating_json, serde_json::json!(false), "publisher2 must observe the same live Publisher instance publisher mutated");

        invoke_member(&publisher, "Quit", vec![], vec![]).expect("Quit");

        unsafe {
            CoUninitialize();
        }
    }

    /// The unit round trip in value.rs cannot catch a consistently swapped
    /// dimension order -- writing and reading with the same wrong convention
    /// round trips fine. Excel is the arbiter: after writing a 2x3 array to
    /// A1:C2, cell B1 must be the first row's second element.
    #[test]
    fn test_two_dimensional_array_lands_in_excel_row_major() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        let _guard = lock_excel_for_test();

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");
        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).expect("Visible=false");
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![])
            .expect("DisplayAlerts=false");

        let workbooks_v = invoke_member(&xl, "Workbooks", vec![], vec![]).expect("Workbooks");
        let workbooks = dispatch_from_variant(&workbooks_v);
        invoke_member(&workbooks, "Add", vec![], vec![]).expect("Workbooks.Add");

        let sheets_v = invoke_member(&xl, "Worksheets", vec![], vec![]).expect("Worksheets");
        let sheets = dispatch_from_variant(&sheets_v);
        let sheet_v = invoke_member(&sheets, "", vec![VARIANT::from(1i32)], vec![]).expect("Worksheets(1)");
        let sheet = dispatch_from_variant(&sheet_v);

        // 2 rows x 3 cols. Values encode their own position so a transpose
        // is unmistakable.
        let payload = serde_json::json!([[11, 12, 13], [21, 22, 23]]);
        let payload_variant = crate::value::json_to_variant(&payload, |_| {
            panic!("no $ole_ref in this payload")
        })
        .expect("json_to_variant should build a 2-D SAFEARRAY");

        let range_v = invoke_member(&sheet, "Range", vec![VARIANT::from("A1:C2")], vec![])
            .expect("Range(A1:C2)");
        let range = dispatch_from_variant(&range_v);
        invoke_member(&range, "Value=", vec![payload_variant], vec![]).expect("Range.Value = 2-D array");

        // B1 is row 1, column 2 -- it must hold 12, not 21.
        let b1_v = invoke_member(&sheet, "Range", vec![VARIANT::from("B1")], vec![]).expect("Range(B1)");
        let b1 = dispatch_from_variant(&b1_v);
        let b1_value = invoke_member(&b1, "Value", vec![], vec![]).expect("B1.Value");
        let b1_json = crate::value::variant_to_json(&b1_value, |_| panic!("B1 should be a scalar"));
        assert_eq!(
            b1_json,
            serde_json::json!(12.0),
            "B1 must be the first row's second element -- getting 21 means the \
             SAFEARRAY dimensions are transposed"
        );

        // ...and read the whole range back as one array.
        let read_v = invoke_member(&range, "Value", vec![], vec![]).expect("Range.Value");
        let read_json = crate::value::variant_to_json(&read_v, |_| panic!("range should be an array"));
        assert_eq!(
            read_json,
            serde_json::json!([[11.0, 12.0, 13.0], [21.0, 22.0, 23.0]]),
            "the bulk read must come back in the same shape it was written"
        );

        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");

        // Drop every COM handle before CoUninitialize -- letting them drop
        // afterward can leak the underlying process (see session.rs's note
        // on ordering).
        drop(b1);
        drop(b1_v);
        drop(range);
        drop(range_v);
        drop(sheet);
        drop(sheet_v);
        drop(sheets);
        drop(sheets_v);
        drop(workbooks);
        drop(workbooks_v);
        drop(xl);
        unsafe {
            CoUninitialize();
        }
    }

    /// The outbound half of the argument-order pair. Nothing else in this
    /// repo passes more than ONE positional argument to Excel, so
    /// `raw_invoke`'s `positional.reverse()` -- the opposite of what the sink
    /// does on the way in -- has never been exercised against a live server;
    /// a one-argument call cannot tell the two directions apart.
    ///
    /// Excel is the arbiter. `Cells(row, column)` and `Offset(rows, columns)`
    /// both take two numbers whose meanings differ, and `Address` reports back
    /// which cell was meant, so a swap is unmistakable: Cells(2, 3) is $C$2,
    /// and a reversed call would answer $B$3.
    #[test]
    fn test_positional_arguments_reach_excel_in_order() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        let _guard = lock_excel_for_test();

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }
        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");
        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).expect("Visible=false");
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![])
            .expect("DisplayAlerts=false");

        let workbooks = dispatch_from_variant(&invoke_member(&xl, "Workbooks", vec![], vec![]).unwrap());
        invoke_member(&workbooks, "Add", vec![], vec![]).expect("Workbooks.Add");
        let sheets = dispatch_from_variant(&invoke_member(&xl, "Worksheets", vec![], vec![]).unwrap());
        let sheet = dispatch_from_variant(
            &invoke_member(&sheets, "", vec![VARIANT::from(1i32)], vec![]).unwrap(),
        );

        let address_of = |target: &IDispatch| -> String {
            let v = invoke_member(target, "Address", vec![], vec![]).expect("Range.Address");
            crate::value::variant_to_json(&v, |_| panic!("Address is a string"))
                .as_str()
                .expect("Address is a string")
                .to_string()
        };

        // Cells(row, column): row 2, column 3 is C2.
        let c2 = dispatch_from_variant(
            &invoke_member(
                &sheet,
                "Cells",
                vec![VARIANT::from(2i32), VARIANT::from(3i32)],
                vec![],
            )
            .expect("Cells(2, 3)"),
        );
        let c2_address = address_of(&c2);

        // ...and the other way round, so a test that passed on symmetry
        // (a call that reads the same reversed) cannot exist here.
        let a5 = dispatch_from_variant(
            &invoke_member(
                &sheet,
                "Cells",
                vec![VARIANT::from(5i32), VARIANT::from(1i32)],
                vec![],
            )
            .expect("Cells(5, 1)"),
        );
        let a5_address = address_of(&a5);

        // A different method with the same shape: Offset(rows, columns) from
        // A1 by one row and two columns is C2.
        let a1 = dispatch_from_variant(
            &invoke_member(&sheet, "Range", vec![VARIANT::from("A1")], vec![]).expect("Range(A1)"),
        );
        let offset = dispatch_from_variant(
            &invoke_member(
                &a1,
                "Offset",
                vec![VARIANT::from(1i32), VARIANT::from(2i32)],
                vec![],
            )
            .expect("Range(A1).Offset(1, 2)"),
        );
        let offset_address = address_of(&offset);

        drop(offset);
        drop(a1);
        drop(a5);
        drop(c2);
        drop(sheet);
        drop(sheets);
        drop(workbooks);
        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");
        drop(xl);
        unsafe {
            CoUninitialize();
        }

        assert_eq!(
            c2_address, "$C$2",
            "Cells(2, 3) is row 2, column 3; $B$3 means the positional arguments reached \
             Excel in the wrong order"
        );
        assert_eq!(
            a5_address, "$A$5",
            "Cells(5, 1) is row 5, column 1; $E$1 means the positional arguments reached \
             Excel in the wrong order"
        );
        assert_eq!(
            offset_address, "$C$2",
            "Range(A1).Offset(1, 2) moves one row down and two columns right; $B$3 means \
             the positional arguments reached Excel in the wrong order"
        );
    }

    #[test]
    fn test_get_active_object_fails_with_mk_e_unavailable_when_word_is_not_running() {
        // The empirical check this whole feature's fallback condition
        // depends on: confirms Wine's COM implementation actually surfaces
        // the same HRESULT real Windows documents for "nothing is currently
        // running", rather than trusting the documented Windows API
        // contract to hold under Wine without checking (see ComError's own
        // doc comment on Wine's FormatMessage for a prior case where it
        // didn't). Held for the whole test -- see the lock's own doc
        // comment above for why this must not overlap with
        // test_connect_or_create_creates_when_word_is_not_running.
        let _guard = WORD_ROT_LOCK.lock().unwrap_or_else(|e| e.into_inner());

        unsafe {
            CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap();
        }

        let err = get_active_object("Word.Application")
            .expect_err("GetActiveObject must fail when nothing matching the ProgID is running");
        assert_eq!(
            err.code,
            windows::Win32::Foundation::MK_E_UNAVAILABLE,
            "expected MK_E_UNAVAILABLE (0x800401E3), got {:?}",
            err
        );

        unsafe {
            CoUninitialize();
        }
    }

    /// Writing a date and reading a date back proves nothing by itself:
    /// Excel turns a date-shaped string in a General cell into a real date
    /// on its own, so that assertion passes even for the broken path where a
    /// Time was serialised as a string. This writes the same value both ways
    /// and requires the two to be distinguishable.
    ///
    /// The string half uses Ruby's actual `Time#to_s` shape (with a UTC
    /// offset), not a bare "YYYY-MM-DD HH:MM:SS" -- Step 1's probe found
    /// that Excel silently auto-parses the latter into a real date too
    /// (matching Value2, NumberFormat and all), which would make this test
    /// pass even for the broken path it exists to catch. The offset suffix
    /// is what the design spec's own example uses, and it is what actually
    /// leaves Excel unable to recognise the string as a date, so the two
    /// paths only diverge with it present.
    #[test]
    fn test_a_tagged_date_reaches_excel_as_a_date_not_as_text() {
        // Held for the whole test -- see EXCEL_TEST_LOCK's doc comment.
        let _guard = lock_excel_for_test();

        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap(); }
        let xl = create_instance("Excel.Application").expect("CreateObject(Excel.Application)");
        invoke_member(&xl, "Visible=", vec![VARIANT::from(false)], vec![]).unwrap();
        invoke_member(&xl, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]).unwrap();
        let workbooks = dispatch_from_variant(&invoke_member(&xl, "Workbooks", vec![], vec![]).unwrap());
        invoke_member(&workbooks, "Add", vec![], vec![]).unwrap();
        let sheets = dispatch_from_variant(&invoke_member(&xl, "Worksheets", vec![], vec![]).unwrap());
        let sheet = dispatch_from_variant(
            &invoke_member(&sheets, "", vec![VARIANT::from(1i32)], vec![]).unwrap(),
        );

        let cell = |addr: &str| {
            dispatch_from_variant(
                &invoke_member(&sheet, "Range", vec![VARIANT::from(addr)], vec![])
                    .unwrap_or_else(|e| panic!("Range({addr}): {e}")),
            )
        };

        // A1 via the tagged path this task exists to prove.
        let tagged = crate::value::json_to_variant(
            &serde_json::json!({"$type": "time", "iso8601": "2026-08-31T09:30:00"}),
            |_| panic!("no $ole_ref in this payload"),
        )
        .expect("the tag must convert");
        invoke_member(&cell("A1"), "Value=", vec![tagged], vec![]).expect("A1 = date");

        // A2 as text, the way the broken path used to send it: Ruby's
        // JSON.generate turns a Time into Time#to_s, e.g.
        // "2026-08-31 09:30:00 +0900" (see the design spec's own example).
        invoke_member(
            &cell("A2"),
            "Value=",
            vec![VARIANT::from("2026-08-31 09:30:00 +0900")],
            vec![],
        )
        .expect("A2 = string");

        // Discriminator (Step 1): Range.Value2. Measured during
        // implementation: it returns a plain number for a genuine VT_DATE
        // cell and a string for a text cell that Excel did not itself
        // recognise as a date -- confirmed on this Wine + Excel build with
        // the exact string above. Both must actually be present for the
        // assertion to mean anything.
        let a1_value2 = invoke_member(&cell("A1"), "Value2", vec![], vec![]).expect("A1.Value2");
        let a2_value2 = invoke_member(&cell("A2"), "Value2", vec![], vec![]).expect("A2.Value2");
        let a1_value2_json = crate::value::variant_to_json(&a1_value2, |_| panic!("A1.Value2 should be a scalar"));
        let a2_value2_json = crate::value::variant_to_json(&a2_value2, |_| panic!("A2.Value2 should be a scalar"));
        assert!(
            a1_value2_json.is_number(),
            "A1 (tagged date) must read back as a numeric serial via Value2, got {a1_value2_json:?}"
        );
        assert!(
            a2_value2_json.is_string(),
            "A2 (broken-path string) must read back as text via Value2 -- if it doesn't, Excel \
             auto-converted it and this test cannot tell the two paths apart; got {a2_value2_json:?}"
        );
        drop(a2_value2);
        drop(a1_value2);

        // The value itself must also survive.
        let a1 = invoke_member(&cell("A1"), "Value", vec![], vec![]).expect("A1.Value");
        assert_eq!(
            crate::value::variant_to_json(&a1, |_| panic!("A1 should be a scalar")),
            serde_json::json!({"$type": "time", "iso8601": "2026-08-31T09:30:00"}),
            "the date must read back unchanged"
        );

        drop(a1);
        drop(sheet);
        drop(sheets);
        drop(workbooks);
        invoke_member(&xl, "Quit", vec![], vec![]).expect("Quit");
        drop(xl);
        unsafe { CoUninitialize(); }
    }
}
