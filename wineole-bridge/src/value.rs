use serde_json::Value;
use std::ffi::c_void;
use windows::core::VARIANT;
use windows::Win32::System::Com::IDispatch;

const VT_NULL: u16 = 1;
const VT_I2: u16 = 2;
const VT_I4: u16 = 3;
const VT_R4: u16 = 4;
const VT_R8: u16 = 5;
const VT_DATE: u16 = 7;
const VT_BSTR: u16 = 8;
const VT_DISPATCH: u16 = 9;
const VT_BOOL: u16 = 11;

pub fn json_to_variant<F>(v: &Value, mut resolve_ref: F) -> windows::core::Result<VARIANT>
where
    F: FnMut(u64) -> windows::core::Result<IDispatch>,
{
    match v {
        Value::Null => Ok(VARIANT::new()),
        Value::Bool(b) => Ok(VARIANT::from(*b)),
        Value::Number(n) => {
            match n.as_i64() {
                // Only integers that actually fit in an i32 may become VT_I4:
                // `i as i32` silently wraps for anything wider (9_012_345_678
                // would arrive at COM as garbage), so anything out of range
                // falls through to the same VT_R8 encoding used for genuinely
                // fractional numbers. A double loses precision above 2^53, but
                // that is a documented, uniform loss rather than a wrap.
                Some(i) if i32::try_from(i).is_ok() => Ok(VARIANT::from(i as i32)),
                Some(i) => Ok(VARIANT::from(i as f64)),
                None => Ok(VARIANT::from(n.as_f64().unwrap_or(0.0))),
            }
        }
        Value::String(s) => Ok(VARIANT::from(s.as_str())),
        Value::Object(map) => {
            let id = map
                .get("$ole_ref")
                .and_then(|v| v.as_u64())
                .ok_or_else(|| windows::core::Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))?;
            let disp = resolve_ref(id)?;
            Ok(variant_from_dispatch(disp))
        }
        Value::Array(_) => Err(windows::core::Error::from_hresult(windows::Win32::Foundation::E_NOTIMPL)),
    }
}

pub fn variant_to_json<F>(v: &VARIANT, mut register_ref: F) -> Value
where
    F: FnMut(IDispatch) -> u64,
{
    unsafe {
        let raw = v.as_raw();
        match raw.Anonymous.Anonymous.vt {
            0 => Value::Null, // VT_EMPTY
            vt if vt == VT_NULL => Value::Null,
            vt if vt == VT_I4 => Value::from(raw.Anonymous.Anonymous.Anonymous.lVal),
            vt if vt == VT_I2 => Value::from(raw.Anonymous.Anonymous.Anonymous.iVal as i32),
            vt if vt == VT_R4 => Value::from(raw.Anonymous.Anonymous.Anonymous.fltVal as f64),
            vt if vt == VT_R8 => Value::from(raw.Anonymous.Anonymous.Anonymous.dblVal),
            vt if vt == VT_DATE => serde_json::json!({
                "$type": "time",
                "iso8601": ole_date_to_iso8601(raw.Anonymous.Anonymous.Anonymous.date),
            }),
            vt if vt == VT_BOOL => Value::from(raw.Anonymous.Anonymous.Anonymous.boolVal != 0),
            vt if vt == VT_BSTR => Value::from(read_bstr(raw.Anonymous.Anonymous.Anonymous.bstrVal)),
            vt if vt == VT_DISPATCH => {
                let ptr = raw.Anonymous.Anonymous.Anonymous.pdispVal;
                if ptr.is_null() {
                    Value::Null
                } else {
                    let ptr_raw = ptr as *const c_void;
                    let ptr_mut = ptr_raw as *mut c_void;
                    let disp = <IDispatch as windows::core::Interface>::from_raw_borrowed(&ptr_mut)
                        .expect("null pdispVal")
                        .clone();
                    let id = register_ref(disp);
                    serde_json::json!({"$ole_ref": id})
                }
            }
            // Anything we don't know how to represent is reported as such
            // rather than silently collapsing to `null`, which is
            // indistinguishable from a genuine empty/null VARIANT and hides
            // the fact that data was lost.
            vt => serde_json::json!({"$unsupported_vt": vt}),
        }
    }
}

/// Convert an OLE Automation `DATE` (`VT_DATE`) into an ISO 8601 timestamp.
///
/// An OLE Automation date is an `f64` counting days since 1899-12-30 00:00:00.
/// The fractional part is the time of day and is always measured *forward*
/// from midnight, even for negative (pre-1899) dates — hence the `.abs()`.
///
/// The rendered form is `YYYY-MM-DDTHH:MM:SS` with **no** timezone designator,
/// because a `VT_DATE` carries no timezone: it is a local wall-clock time as
/// the automation server sees it (an Excel date cell, typically). Sub-second
/// precision is not represented; the value is rounded to the nearest second.
fn ole_date_to_iso8601(date: f64) -> String {
    // Days from the Unix epoch (1970-01-01) back to the OLE epoch (1899-12-30).
    const OLE_EPOCH_UNIX_DAYS: i64 = -25569;

    let mut days = date.trunc() as i64;
    let mut secs = (date.fract().abs() * 86_400.0).round() as i64;
    if secs >= 86_400 {
        secs -= 86_400;
        days += 1;
    }
    let (year, month, day) = civil_from_days(days + OLE_EPOCH_UNIX_DAYS);
    format!(
        "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
        year,
        month,
        day,
        secs / 3600,
        (secs % 3600) / 60,
        secs % 60
    )
}

/// Days-since-1970-01-01 → (year, month, day), via Howard Hinnant's
/// `civil_from_days` algorithm (proleptic Gregorian, valid far beyond any
/// range a VT_DATE can express).
fn civil_from_days(unix_days: i64) -> (i64, u32, u32) {
    let z = unix_days + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097; // [0, 146096]
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365; // [0, 399]
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100); // [0, 365]
    let mp = (5 * doy + 2) / 153; // [0, 11]
    let d = (doy - (153 * mp + 2) / 5 + 1) as u32; // [1, 31]
    let m = (if mp < 10 { mp + 3 } else { mp - 9 }) as u32; // [1, 12]
    (if m <= 2 { y + 1 } else { y }, m, d)
}

pub fn variant_from_dispatch(disp: IDispatch) -> VARIANT {
    unsafe {
        let raw = windows::core::imp::VARIANT {
            Anonymous: windows::core::imp::VARIANT_0 {
                Anonymous: windows::core::imp::VARIANT_0_0 {
                    vt: VT_DISPATCH,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: windows::core::imp::VARIANT_0_0_0 {
                        pdispVal: windows::core::Interface::into_raw(disp),
                    },
                },
            },
        };
        VARIANT::from_raw(raw)
    }
}

/// Test-only helper: the production paths (`session.rs`) never unwrap an
/// `IDispatch` out of a VARIANT by hand — `variant_to_json`'s `register_ref`
/// closure does that as part of marshaling. Only `dispatch.rs`'s real-Excel
/// tests need it, so it is gated to avoid a dead-code warning in the binary.
#[cfg(test)]
pub fn dispatch_from_variant(v: &VARIANT) -> IDispatch {
    unsafe {
        let raw = v.as_raw();
        assert_eq!(raw.Anonymous.Anonymous.vt, VT_DISPATCH, "expected VT_DISPATCH");
        let ptr = raw.Anonymous.Anonymous.Anonymous.pdispVal;
        let ptr_raw = ptr as *const c_void;
        let ptr_mut = ptr_raw as *mut c_void;
        <IDispatch as windows::core::Interface>::from_raw_borrowed(&ptr_mut).expect("null pdispVal").clone()
    }
}

fn read_bstr(ptr: *const u16) -> String {
    if ptr.is_null() {
        return String::new();
    }
    unsafe {
        // BSTR layout: a 4-byte length-in-bytes prefix immediately before the
        // pointer, followed by UTF-16 code units (no reliance on the
        // propsys-backed TryFrom<&VARIANT> for BSTR, which is unsafe to use here).
        let len_bytes = *(ptr as *const u32).offset(-1);
        let len_chars = (len_bytes / 2) as usize;
        let slice = std::slice::from_raw_parts(ptr, len_chars);
        String::from_utf16_lossy(slice)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn no_ref_expected(_id: u64) -> windows::core::Result<IDispatch> {
        panic!("no $ole_ref expected in this test")
    }

    #[test]
    fn test_json_to_variant_primitives_round_trip() {
        let v = json_to_variant(&serde_json::json!(true), no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), serde_json::json!(true));

        let v = json_to_variant(&serde_json::json!(42), no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), serde_json::json!(42));

        let v = json_to_variant(&serde_json::json!(3.5), no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), serde_json::json!(3.5));

        let v = json_to_variant(&serde_json::json!("hello"), no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), serde_json::json!("hello"));

        let v = json_to_variant(&serde_json::Value::Null, no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), serde_json::Value::Null);
    }

    #[test]
    fn test_json_to_variant_ole_ref_calls_resolver() {
        let mut resolved_id = None;
        let resolver = |id: u64| -> windows::core::Result<IDispatch> {
            resolved_id = Some(id);
            Err(windows::core::Error::from_hresult(windows::Win32::Foundation::E_NOTIMPL))
        };
        let _ = json_to_variant(&serde_json::json!({"$ole_ref": 7}), resolver);
        assert_eq!(resolved_id, Some(7));
    }

    #[test]
    fn test_json_to_variant_out_of_i32_range_integer_becomes_a_double() {
        // 9_012_345_678 truncated to i32 used to come out as 422_411_086.
        let v = json_to_variant(&serde_json::json!(9_012_345_678i64), no_ref_expected).unwrap();
        unsafe {
            assert_eq!(v.as_raw().Anonymous.Anonymous.vt, VT_R8);
        }
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!(9_012_345_678.0f64)
        );

        // ...and the negative side of the range, too.
        let v = json_to_variant(&serde_json::json!(-9_012_345_678i64), no_ref_expected).unwrap();
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!(-9_012_345_678.0f64)
        );

        // Values that do fit are still VT_I4 (unchanged behavior).
        let v = json_to_variant(&serde_json::json!(i32::MAX), no_ref_expected).unwrap();
        unsafe {
            assert_eq!(v.as_raw().Anonymous.Anonymous.vt, VT_I4);
        }
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!(i32::MAX)
        );
    }

    /// Build a VARIANT directly through the raw union (the only safe way to
    /// read/write VARIANTs in this crate — see the propsys constraint).
    fn raw_variant(vt: u16, fill: impl FnOnce(&mut windows::core::imp::VARIANT_0_0_0)) -> VARIANT {
        unsafe {
            let mut payload = windows::core::imp::VARIANT_0_0_0 { llVal: 0 };
            fill(&mut payload);
            VARIANT::from_raw(windows::core::imp::VARIANT {
                Anonymous: windows::core::imp::VARIANT_0 {
                    Anonymous: windows::core::imp::VARIANT_0_0 {
                        vt,
                        wReserved1: 0,
                        wReserved2: 0,
                        wReserved3: 0,
                        Anonymous: payload,
                    },
                },
            })
        }
    }

    #[test]
    fn test_variant_to_json_reads_vt_r4() {
        let v = raw_variant(VT_R4, |p| p.fltVal = 2.5f32);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!(2.5)
        );
    }

    #[test]
    fn test_variant_to_json_reads_vt_date_as_tagged_iso8601() {
        // OLE epoch itself.
        let v = raw_variant(VT_DATE, |p| p.date = 0.0);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!({"$type": "time", "iso8601": "1899-12-30T00:00:00"})
        );

        // 2.0 == 1900-01-01; the .5 is midday.
        let v = raw_variant(VT_DATE, |p| p.date = 2.5);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!({"$type": "time", "iso8601": "1900-01-01T12:00:00"})
        );

        // 25569.0 == the Unix epoch, a well-known OLE date landmark.
        let v = raw_variant(VT_DATE, |p| p.date = 25_569.0);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!({"$type": "time", "iso8601": "1970-01-01T00:00:00"})
        );

        // A realistic Excel date cell: 2026-08-26 09:30:00.
        let v = raw_variant(VT_DATE, |p| p.date = 46_260.0 + 0.39583333333);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!({"$type": "time", "iso8601": "2026-08-26T09:30:00"})
        );
    }

    #[test]
    fn test_variant_to_json_flags_unsupported_vt_instead_of_null() {
        // VT_UI1 (17) is not handled; it must be visibly reported, not silently
        // returned as `null` (which a caller cannot tell from a real nil).
        let v = raw_variant(17, |p| p.bVal = 3);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!({"$unsupported_vt": 17})
        );

        // ...while VT_NULL genuinely is nil.
        let v = raw_variant(VT_NULL, |p| p.llVal = 0);
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::Value::Null
        );
    }
}
