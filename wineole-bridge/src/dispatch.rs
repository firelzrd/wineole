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

const LOCALE_USER_DEFAULT: u32 = 0x0400;
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
    static WORD_ROT_LOCK: std::sync::Mutex<()> = std::sync::Mutex::new(());

    #[test]
    fn test_full_round_trip_against_real_excel() {
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
}
