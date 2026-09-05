use serde_json::Value;
use std::ffi::c_void;
use windows::core::VARIANT;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Com::{SAFEARRAY, SAFEARRAYBOUND};
use windows::Win32::System::Ole::{
    SafeArrayCreate, SafeArrayDestroy, SafeArrayGetDim, SafeArrayGetLBound, SafeArrayGetUBound,
    SafeArrayGetElement, SafeArrayGetVartype, SafeArrayPutElement,
};
use windows::Win32::System::Variant::VARENUM;
use crate::dispatch::{ComError, ComResult};

const VT_EMPTY: u16 = 0;
const VT_NULL: u16 = 1;
const VT_I2: u16 = 2;
const VT_I4: u16 = 3;
const VT_R4: u16 = 4;
const VT_R8: u16 = 5;
const VT_DATE: u16 = 7;
const VT_BSTR: u16 = 8;
const VT_DISPATCH: u16 = 9;
const VT_BOOL: u16 = 11;
const VT_VARIANT: u16 = 12;
const VT_ARRAY: u16 = 0x2000;
const VT_BYREF: u16 = 0x4000;

// Days from the Unix epoch (1970-01-01) back to the OLE epoch (1899-12-30).
const OLE_EPOCH_UNIX_DAYS: i64 = -25569;

pub fn json_to_variant<F>(v: &Value, mut resolve_ref: F) -> ComResult<VARIANT>
where
    F: FnMut(u64) -> windows::core::Result<IDispatch>,
{
    json_to_variant_ref(v, &mut resolve_ref)
}

/// The actual conversion, taking `resolve_ref` by mutable reference rather
/// than by value.
///
/// The mirror image of `variant_to_json_ref`, and it exists for the same
/// reason: the SAFEARRAY path recurses back into "convert this JSON value"
/// once per element, and routing that recursion through the public by-value
/// `json_to_variant<F>` would instantiate a fresh closure type at every level
/// (`F`, then the closure `|id| resolve_ref(id)` wrapping it, then one
/// wrapping *that*, ...) -- polymorphic recursion that never terminates at
/// compile time. Rejecting nested arrays at runtime does not help;
/// monomorphization is static and cannot see the check. Recursing through
/// this by-reference helper keeps `F` one concrete type at every call site.
fn json_to_variant_ref<F>(v: &Value, resolve_ref: &mut F) -> ComResult<VARIANT>
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
            // The receive side emits this tag for VT_DATE; accepting it here
            // is what makes a date survive a read-modify-write instead of
            // going back as a string.
            if map.get("$type").and_then(|t| t.as_str()) == Some("time") {
                let iso = map
                    .get("iso8601")
                    .and_then(|v| v.as_str())
                    .ok_or_else(|| {
                        ComError::new(
                            windows::Win32::Foundation::E_INVALIDARG,
                            "a time argument needs an iso8601 string",
                        )
                    })?;
                let date = ole_date_from_iso8601(iso).map_err(|detail| {
                    ComError::new(windows::Win32::Foundation::E_INVALIDARG, detail)
                })?;
                return Ok(variant_from_ole_date(date));
            }

            let id = map
                .get("$ole_ref")
                .and_then(|v| v.as_u64())
                .ok_or_else(|| {
                    ComError::new(
                        windows::Win32::Foundation::E_INVALIDARG,
                        "object argument must be an $ole_ref reference",
                    )
                })?;
            let disp = resolve_ref(id).map_err(ComError::from)?;
            Ok(variant_from_dispatch(disp))
        }
        Value::Array(items) => json_array_to_variant(items, resolve_ref),
    }
}

pub fn variant_to_json<F>(v: &VARIANT, mut register_ref: F) -> Value
where
    F: FnMut(IDispatch) -> u64,
{
    variant_to_json_ref(v, &mut register_ref)
}

/// The actual conversion, taking `register_ref` by mutable reference rather
/// than by value.
///
/// This split exists only to dodge a Rust monomorphization trap: the
/// SAFEARRAY path recurses back into "convert this VARIANT" once per
/// element, and if that recursive call went through the public by-value
/// `variant_to_json<F>` entry point, each recursion would instantiate a new,
/// one-layer-more-wrapped closure type (`F`, then `&mut F`, then
/// `&mut &mut F`, ...) -- polymorphic recursion that never terminates at
/// compile time (`reached the recursion limit while instantiating
/// variant_to_json::<&mut &mut ...>`). Recursing through this by-reference
/// helper instead keeps `F` the same concrete type at every call site, so
/// there is exactly one instantiation regardless of how deep a SAFEARRAY's
/// elements nest.
fn variant_to_json_ref<F>(v: &VARIANT, register_ref: &mut F) -> Value
where
    F: FnMut(IDispatch) -> u64,
{
    unsafe {
        let raw = v.as_raw();
        let vt = raw.Anonymous.Anonymous.vt;

        if vt & VT_ARRAY != 0 {
            // A BYREF array is a SAFEARRAY** rather than a SAFEARRAY*.
            // Excel's Range.Value never returns one, so rather than write
            // deref code that nothing exercises, report it -- the module's
            // standing rule is to never silently drop data.
            if vt & VT_BYREF != 0 {
                return serde_json::json!({"$unsupported_vt": vt});
            }
            let psa = raw.Anonymous.Anonymous.Anonymous.parray as *const SAFEARRAY;
            return safearray_to_json(psa, vt, register_ref);
        }

        match vt {
            vt if vt == VT_EMPTY => Value::Null,
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
            vt if vt == VT_DISPATCH => match variant_to_dispatch(v) {
                // A null pdispVal is COM's `Nothing`, which really is JSON
                // null -- nothing was lost, so this is not `$unsupported_vt`.
                None => Value::Null,
                Some(disp) => serde_json::json!({"$ole_ref": register_ref(disp)}),
            },
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

/// (year, month, day) → days since 1970-01-01, the inverse of
/// `civil_from_days`. Howard Hinnant's `days_from_civil`, the companion of
/// the algorithm above; keeping the pair together is what makes the two
/// directions verifiably each other's inverse.
fn days_from_civil(y: i64, m: u32, d: u32) -> i64 {
    let y = if m <= 2 { y - 1 } else { y };
    let era = if y >= 0 { y } else { y - 399 } / 400;
    let yoe = y - era * 400; // [0, 399]
    let mp = (if m > 2 { m - 3 } else { m + 9 }) as i64; // [0, 11]
    let doy = (153 * mp + 2) / 5 + d as i64 - 1; // [0, 365]
    let doe = yoe * 365 + yoe / 4 - yoe / 100 + doy; // [0, 146096]
    era * 146_097 + doe - 719_468
}

/// Parse the exact shape `ole_date_to_iso8601` emits -- `YYYY-MM-DDTHH:MM:SS`,
/// no timezone, no sub-second -- into an OLE Automation `DATE`.
///
/// Deliberately strict and hand-written: accepting a looser set of formats
/// would mean guessing at what the caller meant, and this crate has no date
/// library and should not gain one for nineteen fixed characters.
///
/// The fractional part is always measured *forward* from midnight, even for
/// pre-1899 dates, which is why a negative day count subtracts it -- the same
/// convention `ole_date_to_iso8601`'s `.abs()` reflects in the other
/// direction.
fn ole_date_from_iso8601(s: &str) -> Result<f64, String> {
    let b = s.as_bytes();
    if b.len() != 19 || b[4] != b'-' || b[7] != b'-' || b[10] != b'T' || b[13] != b':' || b[16] != b':' {
        return Err(format!("expected YYYY-MM-DDTHH:MM:SS, got {s:?}"));
    }
    // Every field must be plain ASCII digits, except a leading `-` in the
    // year: `ole_date_to_iso8601` formats the year with `{:04}`, which emits
    // that shape for a negative (pre-1899) year, and round-trip symmetry
    // means accepting it back. `parse::<i64>` alone also accepts a leading
    // `+` or `-` in *any* field, so without this check
    // "+026-08-31T09:30:00", "2026-+8-31T09:30:00" and
    // "2026-08-31T+9:30:00" would all be accepted with whatever value
    // happened to result -- silently, since none of them panic.
    let field = |range: std::ops::Range<usize>, allow_leading_minus: bool| -> Result<i64, String> {
        let all_digits = s.as_bytes()[range.clone()].iter().enumerate().all(|(i, c)| {
            c.is_ascii_digit() || (allow_leading_minus && i == 0 && *c == b'-')
        });
        if !all_digits {
            return Err(format!("non-numeric field in {s:?}"));
        }
        s[range].parse::<i64>().map_err(|_| format!("non-numeric field in {s:?}"))
    };
    let year = field(0..4, true)?;
    let month = field(5..7, false)?;
    let day = field(8..10, false)?;
    let hour = field(11..13, false)?;
    let minute = field(14..16, false)?;
    let second = field(17..19, false)?;

    if !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return Err(format!("month or day out of range in {s:?}"));
    }
    if !(0..24).contains(&hour) || !(0..60).contains(&minute) || !(0..60).contains(&second) {
        return Err(format!("time out of range in {s:?}"));
    }

    // Catches 2026-02-30 and friends: a date that does not exist comes back
    // from the round trip as a different date.
    let unix_days = days_from_civil(year, month as u32, day as u32);
    if civil_from_days(unix_days) != (year, month as u32, day as u32) {
        return Err(format!("no such date: {s:?}"));
    }

    let ole_days = unix_days - OLE_EPOCH_UNIX_DAYS;
    let frac = (hour * 3600 + minute * 60 + second) as f64 / 86_400.0;
    Ok(if ole_days < 0 { ole_days as f64 - frac } else { ole_days as f64 + frac })
}

/// Wrap an OLE Automation `DATE` in a `VT_DATE` VARIANT. `VARIANT::from(f64)`
/// would produce `VT_R8`, which Excel stores as a plain number rather than a
/// date, so the union has to be built by hand.
fn variant_from_ole_date(date: f64) -> VARIANT {
    unsafe {
        VARIANT::from_raw(windows::core::imp::VARIANT {
            Anonymous: windows::core::imp::VARIANT_0 {
                Anonymous: windows::core::imp::VARIANT_0_0 {
                    vt: VT_DATE,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: windows::core::imp::VARIANT_0_0_0 { date },
                },
            },
        })
    }
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

/// The `IDispatch` inside a VARIANT, or `None` if it holds anything else
/// (including a `VT_DISPATCH` whose pointer is null, which is COM's `Nothing`).
///
/// The returned interface is a *new* reference — `from_raw_borrowed` does not
/// take ownership and the clone AddRefs — so it stays valid after the source
/// VARIANT is dropped. That is what makes it usable on an event argument,
/// whose VARIANT belongs to the caller.
///
/// This is the crate's one raw unwrap of a dispatch pointer out of a VARIANT;
/// `variant_to_json_ref` and `dispatch_from_variant` both go through it rather
/// than repeating the union access with their own null and vt handling.
pub fn variant_to_dispatch(v: &VARIANT) -> Option<IDispatch> {
    unsafe {
        let raw = v.as_raw();
        if raw.Anonymous.Anonymous.vt != VT_DISPATCH {
            return None;
        }
        let ptr = raw.Anonymous.Anonymous.Anonymous.pdispVal;
        if ptr.is_null() {
            return None;
        }
        let ptr_raw = ptr as *const c_void;
        let ptr_mut = ptr_raw as *mut c_void;
        <IDispatch as windows::core::Interface>::from_raw_borrowed(&ptr_mut).cloned()
    }
}

/// Test-only helper: the production paths (`session.rs`) never unwrap an
/// `IDispatch` out of a VARIANT by hand — `variant_to_json`'s `register_ref`
/// closure does that as part of marshaling. Only `dispatch.rs`'s real-Excel
/// tests need it, so it is gated to avoid a dead-code warning in the binary.
#[cfg(test)]
pub fn dispatch_from_variant(v: &VARIANT) -> IDispatch {
    variant_to_dispatch(v).expect("expected a non-null VT_DISPATCH")
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

/// Wrap an already-built SAFEARRAY in a VARIANT that owns it.
///
/// # Safety
///
/// This **takes ownership** of `psa`. The returned VARIANT's Drop calls
/// VariantClear, which calls SafeArrayDestroy on the pointer -- so the caller
/// must own the array outright and must not destroy it, nor hand the same
/// pointer to this function twice, nor keep using it after the VARIANT drops.
/// `psa` must be a valid, non-null SAFEARRAY created with `elem_vt` elements;
/// a mismatched `elem_vt` makes VariantClear free the elements with the wrong
/// rules.
unsafe fn variant_from_safearray(
    psa: *mut windows::core::imp::SAFEARRAY,
    elem_vt: u16,
) -> VARIANT {
    unsafe {
        VARIANT::from_raw(windows::core::imp::VARIANT {
            Anonymous: windows::core::imp::VARIANT_0 {
                Anonymous: windows::core::imp::VARIANT_0_0 {
                    vt: VT_ARRAY | elem_vt,
                    wReserved1: 0,
                    wReserved2: 0,
                    wReserved3: 0,
                    Anonymous: windows::core::imp::VARIANT_0_0_0 { parray: psa },
                },
            },
        })
    }
}

fn invalid_array(detail: &str) -> ComError {
    ComError::new(windows::Win32::Foundation::E_INVALIDARG, detail)
}

/// Build a SAFEARRAY(VT_VARIANT) from a JSON array.
///
/// A flat array of scalars becomes one dimension; an array whose first
/// element is an array becomes two (rows x columns). Anything else is
/// rejected with a message rather than coerced: a ragged array has no
/// correct padding, a scalar next to a row has no shape at all, and a deeper
/// nesting has no Excel meaning.
///
/// Dimension order is `bounds[0] = rows`, `bounds[1] = columns`, matching
/// what `safearray_to_json` reads back on the receive side.
fn json_array_to_variant<F>(items: &[Value], resolve_ref: &mut F) -> ComResult<VARIANT>
where
    F: FnMut(u64) -> windows::core::Result<IDispatch>,
{
    let two_d = matches!(items.first(), Some(Value::Array(_)));

    let bounds: Vec<SAFEARRAYBOUND> = if two_d {
        let cols = match &items[0] {
            Value::Array(row) => row.len(),
            _ => unreachable!("two_d implies the first element is an array"),
        };
        for (i, item) in items.iter().enumerate() {
            match item {
                Value::Array(row) => {
                    if row.len() != cols {
                        return Err(invalid_array(&format!(
                            "ragged array: row 0 has {} elements but row {} has {}",
                            cols,
                            i,
                            row.len()
                        )));
                    }
                    if row.iter().any(|e| matches!(e, Value::Array(_))) {
                        return Err(invalid_array(
                            "arrays of more than two dimensions are not supported",
                        ));
                    }
                }
                _ => {
                    return Err(invalid_array(
                        "array mixes rows and scalars: every element must be a row, or none",
                    ))
                }
            }
        }
        vec![
            SAFEARRAYBOUND { cElements: items.len() as u32, lLbound: 0 },
            SAFEARRAYBOUND { cElements: cols as u32, lLbound: 0 },
        ]
    } else {
        if items.iter().any(|e| matches!(e, Value::Array(_))) {
            return Err(invalid_array(
                "array mixes scalars and rows: every element must be a row, or none",
            ));
        }
        vec![SAFEARRAYBOUND { cElements: items.len() as u32, lLbound: 0 }]
    };

    unsafe {
        let psa = SafeArrayCreate(VARENUM(VT_VARIANT), bounds.len() as u32, bounds.as_ptr());
        if psa.is_null() {
            // Not `invalid_array`: the argument shape was fine (validation
            // above already passed it) -- SafeArrayCreate itself failed to
            // allocate, which is an out-of-memory condition, not a malformed
            // argument. Reporting E_INVALIDARG here would tell a client "your
            // data is wrong" when the real story is "the bridge ran out of
            // memory".
            return Err(ComError::new(
                windows::Win32::Foundation::E_OUTOFMEMORY,
                "SafeArrayCreate failed",
            ));
        }

        // From here on the array is owned locally: every early return has to
        // destroy it, and only the VARIANT built at the very end may take it
        // over. An empty array (or an empty row) simply stores nothing --
        // a dimension with zero elements has no valid index at all, not even
        // its lower bound, so a put here would fail with DISP_E_BADINDEX.
        if two_d {
            for (r, row) in items.iter().enumerate() {
                let row = match row {
                    Value::Array(row) => row,
                    _ => unreachable!("validated above"),
                };
                for (c, cell) in row.iter().enumerate() {
                    if let Err(e) = put_json_element(psa, &[r as i32, c as i32], cell, resolve_ref)
                    {
                        let _ = SafeArrayDestroy(psa);
                        return Err(e);
                    }
                }
            }
        } else {
            for (i, item) in items.iter().enumerate() {
                if let Err(e) = put_json_element(psa, &[i as i32], item, resolve_ref) {
                    let _ = SafeArrayDestroy(psa);
                    return Err(e);
                }
            }
        }

        // Hands ownership of `psa` over to the VARIANT, whose Drop destroys
        // it; nothing below may touch the pointer again.
        Ok(variant_from_safearray(psa as *mut _, VT_VARIANT))
    }
}

/// Convert one JSON value and store it into an already-created SAFEARRAY.
///
/// A free function rather than a closure inside `json_array_to_variant`: the
/// array-filling loops need it in two shapes (1-D and 2-D indices), and
/// keeping it out here means it does not depend on an `unsafe` block
/// lexically enclosing a closure body.
///
/// The local VARIANT owns its converted value. `SafeArrayPutElement` copies
/// the value into the array (VariantCopy semantics for a VT_VARIANT array),
/// and the VARIANT's Drop clears our copy on the way out, so nothing leaks
/// whether this succeeds or fails.
///
/// # Safety
///
/// `psa` must be a valid, live SAFEARRAY whose element type is VT_VARIANT,
/// and `indices` must name one element per dimension, within bounds.
unsafe fn put_json_element<F>(
    psa: *mut SAFEARRAY,
    indices: &[i32],
    value: &Value,
    resolve_ref: &mut F,
) -> ComResult<()>
where
    F: FnMut(u64) -> windows::core::Result<IDispatch>,
{
    // Straight back into the by-reference entry point, not the public
    // by-value one: see `json_to_variant_ref` for why that matters.
    let elem = json_to_variant_ref(value, resolve_ref)?;
    SafeArrayPutElement(psa, indices.as_ptr(), &elem as *const VARIANT as *const c_void)
        .map_err(ComError::from)
}

/// The element types a *typed* SAFEARRAY (VT_ARRAY|VT_R8 and friends) may be
/// read out of, one bare value at a time.
///
/// This is deliberately an allowlist rather than a size check, and it must
/// stay in step with the scalar match in `variant_to_json_ref`: it lists
/// exactly the types that match handles, so a type admitted here is always
/// one the match can render. Two rules for anyone editing it:
///
/// * Teaching the scalar match a new type does **not** automatically teach
///   the array path -- add the type here too, or arrays of it will keep
///   reporting `$unsupported_vt` while scalars of it convert.
/// * Never widen this to "anything". Every type listed here occupies at most
///   8 bytes, which is what bounds `SafeArrayGetElement`'s write into the
///   VARIANT payload union at the call site -- 16 bytes on x86_64 and
///   aarch64, but only 8 on the i686 target this project also ships, where
///   the margin is therefore zero, not eight. VT_RECORD in particular
///   has an element size of the record, not of a union member, and would
///   overrun the destination. (VT_DECIMAL is absent for a second reason: a
///   DECIMAL overlays the whole VARIANT, not the payload union, so a VARIANT
///   built this way would be malformed even if it fitted.)
///
/// Strictly, the write is bounded by `psa->cbElements`, and this list bounds
/// it only because `SafeArrayCreate`/`SafeArrayCreateEx` derive `cbElements`
/// from the vartype -- true of every array reachable through the SAFEARRAY
/// API. The one realistic divergence, a FADF_RECORD array whose `cbElements`
/// is the record size, is exactly what `SafeArrayGetVartype` reports as
/// VT_RECORD, which this list refuses. (VT_EMPTY and VT_NULL can't actually
/// be a real array's element type -- `SafeArrayCreate` rejects them -- they
/// are listed only to mirror the scalar match.)
const TYPED_ELEMENT_ALLOWLIST: [u16; 10] = [
    VT_EMPTY, VT_NULL, VT_I2, VT_I4, VT_R4, VT_R8, VT_DATE, VT_BSTR, VT_DISPATCH, VT_BOOL,
];

/// Read one element out of a SAFEARRAY and convert it with the same scalar
/// rules the rest of this module uses.
///
/// For a VT_VARIANT array the element *is* a VARIANT, so it can be read
/// straight into one. For a typed array (VT_ARRAY|VT_R8 and friends, which
/// some automation servers return) the element is a bare value, so it is
/// read into a raw union payload and then labelled with the element type --
/// which lands it back on the existing scalar match.
///
/// Either way the local VARIANT owns its copy: SafeArrayGetElement hands
/// back a copy (VariantCopy semantics for VT_VARIANT, a fresh BSTR for a
/// string), and the VARIANT's Drop calls VariantClear on it.
///
/// The typed path is guarded by `TYPED_ELEMENT_ALLOWLIST`; see its comment for
/// why the guard has to come before the read.
unsafe fn safearray_element_to_json<F>(
    psa: *const SAFEARRAY,
    indices: &[i32],
    elem_vt: u16,
    register_ref: &mut F,
) -> Value
where
    F: FnMut(IDispatch) -> u64,
{
    if elem_vt == VT_VARIANT {
        let mut elem = VARIANT::new();
        if SafeArrayGetElement(psa, indices.as_ptr(), &mut elem as *mut VARIANT as *mut c_void)
            .is_err()
        {
            // Not `null`: a caller could not tell a dropped element from a
            // genuinely empty cell, and this module reports data loss.
            return serde_json::json!({"$unsupported_vt": elem_vt});
        }
        // Reborrow rather than move: this helper is called once per element.
        return variant_to_json_ref(&elem, register_ref);
    }

    // Refuse the element *before* reading it: SafeArrayGetElement copies
    // `psa->cbElements` bytes into the destination, and the destination below
    // is a 16-byte VARIANT payload union. For a FADF_RECORD array cbElements
    // is the record size and the copy runs through IRecordInfo::RecordCopy --
    // arbitrarily many bytes into a stack slot -- after which labelling the
    // result VT_RECORD would have VariantClear read `pRecInfo` out of raw
    // record bytes and make a virtual call through it.
    if !TYPED_ELEMENT_ALLOWLIST.contains(&elem_vt) {
        return serde_json::json!({"$unsupported_vt": elem_vt});
    }

    let mut payload = windows::core::imp::VARIANT_0_0_0 { llVal: 0 };
    if SafeArrayGetElement(
        psa,
        indices.as_ptr(),
        &mut payload as *mut windows::core::imp::VARIANT_0_0_0 as *mut c_void,
    )
    .is_err()
    {
        return serde_json::json!({"$unsupported_vt": elem_vt});
    }
    let elem = VARIANT::from_raw(windows::core::imp::VARIANT {
        Anonymous: windows::core::imp::VARIANT_0 {
            Anonymous: windows::core::imp::VARIANT_0_0 {
                vt: elem_vt,
                wReserved1: 0,
                wReserved2: 0,
                wReserved3: 0,
                Anonymous: payload,
            },
        },
    });
    variant_to_json_ref(&elem, register_ref)
}

/// Convert a SAFEARRAY to JSON: a flat array for one dimension, an array of
/// rows for two. Three or more dimensions are reported as unsupported rather
/// than flattened, matching this module's policy of never hiding data loss.
unsafe fn safearray_to_json<F>(
    psa: *const SAFEARRAY,
    vt: u16,
    register_ref: &mut F,
) -> Value
where
    F: FnMut(IDispatch) -> u64,
{
    if psa.is_null() {
        return Value::Null;
    }
    let unsupported = || serde_json::json!({"$unsupported_vt": vt});

    // Report rather than guess. This fails only for an array carrying none of
    // the type-describing FADF_* flags, and the old fallback -- assume
    // VT_VARIANT -- was the worst guess available: it makes
    // SafeArrayGetElement memcpy raw element bytes over a real VARIANT, whose
    // first two bytes are then reinterpreted as `vt`, after which
    // VariantClear on that garbage tag can SysFreeString an arbitrary pointer.
    let elem_vt = match SafeArrayGetVartype(psa) {
        Ok(v) => v.0,
        Err(_) => return unsupported(),
    };

    match SafeArrayGetDim(psa) {
        1 => {
            let (lb, ub) = match (SafeArrayGetLBound(psa, 1), SafeArrayGetUBound(psa, 1)) {
                (Ok(lb), Ok(ub)) => (lb, ub),
                _ => return unsupported(),
            };
            let mut out = Vec::new();
            for i in lb..=ub {
                out.push(safearray_element_to_json(psa, &[i], elem_vt, register_ref));
            }
            Value::Array(out)
        }
        2 => {
            let (r_lb, r_ub) = match (SafeArrayGetLBound(psa, 1), SafeArrayGetUBound(psa, 1)) {
                (Ok(lb), Ok(ub)) => (lb, ub),
                _ => return unsupported(),
            };
            let (c_lb, c_ub) = match (SafeArrayGetLBound(psa, 2), SafeArrayGetUBound(psa, 2)) {
                (Ok(lb), Ok(ub)) => (lb, ub),
                _ => return unsupported(),
            };
            let mut rows = Vec::new();
            for r in r_lb..=r_ub {
                let mut row = Vec::new();
                for c in c_lb..=c_ub {
                    row.push(safearray_element_to_json(psa, &[r, c], elem_vt, register_ref));
                }
                rows.push(Value::Array(row));
            }
            Value::Array(rows)
        }
        _ => unsupported(),
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

    #[test]
    fn test_json_to_variant_round_trips_a_one_dimensional_array() {
        let v = json_to_variant(&serde_json::json!([1, "two", 3.5, null]), no_ref_expected).unwrap();
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([1, "two", 3.5, null])
        );
    }

    #[test]
    fn test_json_to_variant_round_trips_a_two_dimensional_array_without_transposing() {
        // 2 rows x 3 cols, deliberately non-square: a transposed round trip
        // would come back as 3 rows of 2 and fail loudly.
        let input = serde_json::json!([[1, 2, 3], [4, 5, 6]]);
        let v = json_to_variant(&input, no_ref_expected).unwrap();
        assert_eq!(variant_to_json(&v, |_| panic!("no dispatch expected")), input);
    }

    #[test]
    fn test_json_to_variant_round_trips_an_empty_array() {
        let v = json_to_variant(&serde_json::json!([]), no_ref_expected).unwrap();
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([])
        );
    }

    #[test]
    fn test_json_to_variant_round_trips_zero_column_two_dimensional_arrays() {
        // [[]] and [[],[]] are 2-D shapes with a zero-element column
        // dimension: they take the same two_d path as a real matrix, and
        // SafeArrayCreate is asked for a dimension with cElements == 0. That
        // is exactly the shape class behind the DISP_E_BADINDEX bug that hit
        // this project's own test helper earlier (a zero-element dimension
        // has no valid index, not even its lower bound, so a put into it
        // fails) -- worth pinning as a class, not just the 1-D empty array
        // above.
        //
        // On the way back, SafeArrayGetUBound for a zero-element dimension
        // reports LBound - 1 (0 - 1 == -1 here), so each row's inner
        // `0..=-1` column loop runs zero times and contributes an empty row;
        // the row count itself is unaffected. So one empty row survives as
        // one empty row, and two as two.
        let v = json_to_variant(&serde_json::json!([[]]), no_ref_expected).unwrap();
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([[]])
        );

        let v = json_to_variant(&serde_json::json!([[], []]), no_ref_expected).unwrap();
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([[], []])
        );
    }

    #[test]
    fn test_json_to_variant_rejects_a_ragged_array() {
        let err = json_to_variant(&serde_json::json!([[1, 2], [3]]), no_ref_expected)
            .expect_err("a ragged array must be rejected, not silently padded");
        let message = err.to_string();
        assert!(
            message.contains("ragged"),
            "the error must say what was wrong, got {message:?}"
        );
        assert_eq!(err.code, windows::Win32::Foundation::E_INVALIDARG);
    }

    #[test]
    fn test_json_to_variant_rejects_a_mix_of_scalars_and_arrays() {
        // Scalar first, then a row: the 1-D branch's rejection.
        let err = json_to_variant(&serde_json::json!([1, [2, 3]]), no_ref_expected)
            .expect_err("a scalar/array mix must be rejected");
        let message = err.to_string();
        assert!(
            message.contains("scalars and rows"),
            "the error must say what was wrong, got {message:?}"
        );
        assert_eq!(err.code, windows::Win32::Foundation::E_INVALIDARG);
    }

    #[test]
    fn test_json_to_variant_rejects_a_mix_of_arrays_and_scalars() {
        // Row first, then a scalar: the 2-D branch's rejection. Distinct from
        // the scalar-first case above -- both messages contain "mixes", so
        // asserting on the fuller phrase is what actually tells the two
        // branches apart.
        let err = json_to_variant(&serde_json::json!([[1, 2], 3]), no_ref_expected)
            .expect_err("a row/scalar mix must be rejected");
        let message = err.to_string();
        assert!(
            message.contains("rows and scalars"),
            "the error must say what was wrong, got {message:?}"
        );
        assert_eq!(err.code, windows::Win32::Foundation::E_INVALIDARG);
    }

    #[test]
    fn test_json_to_variant_rejects_a_three_dimensional_array() {
        let err = json_to_variant(&serde_json::json!([[[1, 2]], [[3, 4]]]), no_ref_expected)
            .expect_err("a 3-D array must be rejected");
        let message = err.to_string();
        assert!(
            message.contains("dimension"),
            "the error must say what was wrong, got {message:?}"
        );
        assert_eq!(err.code, windows::Win32::Foundation::E_INVALIDARG);
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

    /// Not in the scalar match, and so not in `TYPED_ELEMENT_ALLOWLIST` --
    /// but a perfectly legal SAFEARRAY element type, which makes it the
    /// cheapest way to build an array the allowlist has to refuse.
    const VT_UI1: u16 = 17;

    /// Store one element into a SAFEARRAY the way its element type requires.
    ///
    /// `SafeArrayPutElement`'s `pv` must point at whatever the array actually
    /// stores: a whole VARIANT for a VT_VARIANT array, but the *bare* value
    /// for a typed one. This helper used to pass `&VARIANT` unconditionally,
    /// which for a typed array silently wrote the VARIANT header (vt plus the
    /// three reserved words) into the element -- so a VT_R8 fixture would
    /// have held a reinterpreted header rather than the f64 it claimed.
    ///
    /// Element types this does not know how to unwrap panic rather than store
    /// something plausible-looking: a fixture that lies is worse than one that
    /// refuses.
    unsafe fn put_element(
        psa: *const SAFEARRAY,
        indices: &[i32],
        elem_vt: u16,
        elem: &VARIANT,
    ) {
        if elem_vt == VT_VARIANT {
            SafeArrayPutElement(psa, indices.as_ptr(), elem as *const VARIANT as *const c_void)
                .expect("SafeArrayPutElement");
            return;
        }

        let raw = elem.as_raw();
        let got = raw.Anonymous.Anonymous.vt;
        assert_eq!(
            got, elem_vt,
            "fill() returned a VT {got} VARIANT for a VT {elem_vt} array"
        );
        let payload = raw.Anonymous.Anonymous.Anonymous;
        match elem_vt {
            VT_R8 => {
                let value: f64 = payload.dblVal;
                SafeArrayPutElement(psa, indices.as_ptr(), &value as *const f64 as *const c_void)
            }
            VT_I4 => {
                let value: i32 = payload.lVal;
                SafeArrayPutElement(psa, indices.as_ptr(), &value as *const i32 as *const c_void)
            }
            VT_UI1 => {
                let value: u8 = payload.bVal;
                SafeArrayPutElement(psa, indices.as_ptr(), &value as *const u8 as *const c_void)
            }
            other => panic!(
                "make_safearray_variant cannot build a typed array of VT {other}: \
                 teach put_element how to unwrap that element type first"
            ),
        }
        .expect("SafeArrayPutElement");
    }

    /// Builds a real SAFEARRAY-carrying VARIANT for the receive-side tests.
    /// `dims` is one `(lower_bound, count)` pair per dimension, in the same
    /// order `SafeArrayGetLBound(psa, 1..)` reports them. `fill` is called
    /// with each index tuple and returns the element to store there.
    ///
    /// For a typed array (any `elem_vt` other than VT_VARIANT) `fill` must
    /// return a VARIANT of exactly that type: `put_element` unwraps the bare
    /// value out of it, because that is what the array stores.
    fn make_safearray_variant(
        elem_vt: u16,
        dims: &[(i32, u32)],
        fill: impl Fn(&[i32]) -> VARIANT,
    ) -> VARIANT {
        unsafe {
            let bounds: Vec<SAFEARRAYBOUND> = dims
                .iter()
                .map(|(lb, n)| SAFEARRAYBOUND { cElements: *n, lLbound: *lb })
                .collect();
            let psa = SafeArrayCreate(VARENUM(elem_vt), bounds.len() as u32, bounds.as_ptr());
            assert!(!psa.is_null(), "SafeArrayCreate returned null");

            // A dimension with zero elements has no valid indices at all --
            // not even the lower bound -- so there is nothing to fill.
            if dims.iter().any(|(_, n)| *n == 0) {
                return variant_from_safearray(psa as *mut _, elem_vt);
            }

            // Walk every index tuple in the cartesian product of the bounds.
            let mut idx: Vec<i32> = dims.iter().map(|(lb, _)| *lb).collect();
            loop {
                let elem = fill(&idx);
                put_element(psa, &idx, elem_vt, &elem);

                // Increment the right-most index, carrying leftwards.
                let mut d = dims.len();
                loop {
                    if d == 0 {
                        // Every index exhausted.
                        return variant_from_safearray(psa as *mut _, elem_vt);
                    }
                    d -= 1;
                    idx[d] += 1;
                    if idx[d] < dims[d].0 + dims[d].1 as i32 {
                        break;
                    }
                    idx[d] = dims[d].0;
                }
            }
        }
    }

    #[test]
    fn test_variant_to_json_reads_a_one_dimensional_safearray() {
        let v = make_safearray_variant(VT_VARIANT, &[(0, 3)], |idx| {
            VARIANT::from(idx[0] * 10)
        });
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([0, 10, 20])
        );
    }

    #[test]
    fn test_variant_to_json_reads_a_two_dimensional_safearray_row_major() {
        // 2 rows x 3 cols, deliberately non-square so a transposed result
        // is unmistakable rather than accidentally correct.
        let v = make_safearray_variant(VT_VARIANT, &[(1, 2), (1, 3)], |idx| {
            VARIANT::from(idx[0] * 100 + idx[1])
        });
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([[101, 102, 103], [201, 202, 203]])
        );
    }

    #[test]
    fn test_variant_to_json_honors_a_nonzero_lower_bound() {
        // Excel hands back 1-based arrays; the lower bound must be read,
        // not assumed.
        let v = make_safearray_variant(VT_VARIANT, &[(1, 2)], |idx| VARIANT::from(idx[0]));
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([1, 2])
        );
    }

    #[test]
    fn test_variant_to_json_reads_an_empty_safearray_as_an_empty_array() {
        let v = make_safearray_variant(VT_VARIANT, &[(0, 0)], |_| VARIANT::new());
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([])
        );
    }

    #[test]
    fn test_variant_to_json_reports_a_three_dimensional_safearray_as_unsupported() {
        // Honest reporting over silent data loss, matching the scalar
        // fallback's own policy.
        let v = make_safearray_variant(VT_VARIANT, &[(0, 2), (0, 2), (0, 2)], |_| {
            VARIANT::from(1i32)
        });
        let json = variant_to_json(&v, |_| panic!("no dispatch expected"));
        assert!(
            json.get("$unsupported_vt").is_some(),
            "a 3-D SAFEARRAY must be reported as unsupported, got {json}"
        );
    }

    #[test]
    fn test_variant_to_json_reads_mixed_element_types() {
        // What a real Excel range actually contains: numbers, strings,
        // dates and empty cells side by side.
        let v = make_safearray_variant(VT_VARIANT, &[(1, 4)], |idx| match idx[0] {
            1 => VARIANT::from(42i32),
            2 => VARIANT::from("hello"),
            3 => raw_variant(VT_DATE, |p| p.date = 25_569.0),
            _ => VARIANT::new(),
        });
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([
                42,
                "hello",
                {"$type": "time", "iso8601": "1970-01-01T00:00:00"},
                null
            ])
        );
    }

    #[test]
    fn test_variant_to_json_reads_a_typed_vt_r8_safearray() {
        // Not every automation server hands back VT_ARRAY|VT_VARIANT. A typed
        // array stores bare values, so the element path reads one into a raw
        // payload and labels it with the array's element type before the
        // scalar match can render it -- a branch no other test reaches.
        let v = make_safearray_variant(VT_R8, &[(0, 3)], |idx| {
            VARIANT::from(idx[0] as f64 + 1.5)
        });
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([1.5, 2.5, 3.5])
        );
    }

    #[test]
    fn test_variant_to_json_reports_typed_elements_outside_the_allowlist() {
        // What actually protects against a FADF_RECORD array is the
        // allowlist: SafeArrayGetElement copies cbElements bytes, which for a
        // record is the record size, into a 16-byte VARIANT payload union.
        // Building a real VT_RECORD array needs an IRecordInfo implementation
        // and is not worth it, so this pins the allowlist's behavior with
        // VT_UI1 instead: equally absent from the scalar match, equally
        // refused before any read happens, and constructible with nothing but
        // SafeArrayCreate.
        let v = make_safearray_variant(VT_UI1, &[(0, 2)], |idx| {
            raw_variant(VT_UI1, |p| p.bVal = idx[0] as u8 + 1)
        });
        assert_eq!(
            variant_to_json(&v, |_| panic!("no dispatch expected")),
            serde_json::json!([{"$unsupported_vt": 17}, {"$unsupported_vt": 17}])
        );
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

    /// The allowlist is the memory-safety guard for the typed-array element
    /// read: `SafeArrayGetElement` writes `cbElements` bytes into a 16-byte
    /// payload union, and every admitted type is at most 8 bytes, so the
    /// write is bounded by construction rather than by a size check.
    /// Widening it is therefore a memory-safety change, not a feature -- this
    /// test exists to make that edit fail loudly instead of passing silently.
    #[test]
    fn test_typed_element_allowlist_is_not_widened_without_review() {
        assert_eq!(
            TYPED_ELEMENT_ALLOWLIST,
            [VT_EMPTY, VT_NULL, VT_I2, VT_I4, VT_R4, VT_R8, VT_DATE, VT_BSTR, VT_DISPATCH, VT_BOOL]
        );
        // VT_RECORD's element size is the record's, not a union member's --
        // admitting it is exactly the overrun this guard exists to prevent.
        assert!(!TYPED_ELEMENT_ALLOWLIST.contains(&36));
        // A DECIMAL overlays the whole VARIANT rather than the payload union,
        // so a VARIANT built the typed path's way would be malformed even
        // though it fits.
        assert!(!TYPED_ELEMENT_ALLOWLIST.contains(&14));
    }

    /// The landmarks here are the same ones the receive-side test uses, read
    /// backwards -- if the two directions ever disagree, one of these fails.
    #[test]
    fn test_json_to_variant_converts_a_tagged_time_to_vt_date() {
        let cases = [
            ("1899-12-30T00:00:00", 0.0),          // the OLE epoch itself
            ("1900-01-01T12:00:00", 2.5),          // .5 is midday
            ("1970-01-01T00:00:00", 25_569.0),     // the Unix epoch
            ("2026-08-31T09:30:00", 46_265.0 + 0.395_833_333_333_333_3),
            // Leap-year landmark: 2024-02-29 at midnight. Derived by hand,
            // not read off the implementation -- see the task report for the
            // full days_from_civil(2024, 2, 29) working, which comes out to
            // OLE serial 45351.0 -- and cross-checked two independent ways:
            // (1) the well-known Excel serial for 2024-01-01 is 45292;
            // January has 31 days and February contributes 28 more days up
            // to but not including the 29th, so 2024-02-29 is
            // 45292 + 31 + 28 = 45351. (2) counting backward from the
            // 2026-08-31 landmark two lines up (46265): Aug 31 is
            // day-of-year 243 in a non-leap year (31+28+31+30+31+30+31 = 212
            // days through end of July, plus 31 more), i.e. 242 days after
            // Jan 1, so 2026-01-01 = 46265 - 242 = 46023; 2025-01-01 =
            // 46023 - 365 = 45658; 2024-01-01 = 45658 - 366 (2024 is a leap
            // year) = 45292, matching route (1) exactly. Both routes and the
            // direct days_from_civil hand-computation agree on 45351.
            ("2024-02-29T00:00:00", 45_351.0),
        ];
        for (iso, expected) in cases {
            let v = json_to_variant(
                &serde_json::json!({"$type": "time", "iso8601": iso}),
                no_ref_expected,
            )
            .unwrap_or_else(|e| panic!("{iso} should convert, got {e}"));
            unsafe {
                assert_eq!(v.as_raw().Anonymous.Anonymous.vt, VT_DATE, "{iso} must be VT_DATE");
                let got = v.as_raw().Anonymous.Anonymous.Anonymous.date;
                assert!(
                    (got - expected).abs() < 1e-9,
                    "{iso}: expected OLE date {expected}, got {got}"
                );
            }
        }
    }

    #[test]
    fn test_tagged_time_round_trips_through_both_directions() {
        for iso in ["1899-12-30T00:00:00", "1970-01-01T00:00:00", "2026-08-31T09:30:00"] {
            let v = json_to_variant(
                &serde_json::json!({"$type": "time", "iso8601": iso}),
                no_ref_expected,
            )
            .unwrap();
            assert_eq!(
                variant_to_json(&v, |_| panic!("no dispatch expected")),
                serde_json::json!({"$type": "time", "iso8601": iso}),
                "{iso} must survive a send-then-receive round trip unchanged"
            );
        }
    }

    /// A negative year is the one legitimate exception to "digits only":
    /// `ole_date_to_iso8601` formats the year with `{:04}`, which emits a
    /// leading `-` for a pre-1899 (proleptic Gregorian) year, and that shape
    /// must still parse back -- otherwise the digit-only tightening above
    /// would have broken round-trip symmetry instead of just narrowing the
    /// accepted input.
    #[test]
    fn test_a_negative_year_still_round_trips() {
        for iso in ["-001-03-15T00:00:00", "-026-08-31T09:30:00", "-100-01-01T00:00:00"] {
            let v = json_to_variant(
                &serde_json::json!({"$type": "time", "iso8601": iso}),
                no_ref_expected,
            )
            .unwrap_or_else(|e| panic!("{iso} should convert, got {e}"));
            assert_eq!(
                variant_to_json(&v, |_| panic!("no dispatch expected")),
                serde_json::json!({"$type": "time", "iso8601": iso}),
                "{iso}: a negative year must survive a send-then-receive round trip unchanged"
            );
        }
    }

    #[test]
    fn test_json_to_variant_rejects_a_malformed_iso8601_with_a_reason() {
        let bad = [
            ("2026-08-31", "too short"),
            ("2026-08-31 09:30:00", "space instead of T"),
            ("2026-13-01T00:00:00", "month 13"),
            ("2026-02-30T00:00:00", "no such day"),
            ("2026-08-31T25:00:00", "hour 25"),
            ("20xx-08-31T09:30:00", "non-numeric year"),
            ("2026-08-31T-1:30:00", "negative hour"),
            ("2026-08-31T09:-1:00", "negative minute"),
            ("2026-08-31T09:30:-1", "negative second"),
            ("2023-02-29T00:00:00", "not a leap year"),
            // A leading sign anywhere but the year: parse::<i64> alone
            // accepts these silently (measured before this test existed),
            // which is looser than the doc comment's "deliberately strict"
            // claims. Only the year may carry a leading `-`.
            ("+026-08-31T09:30:00", "leading + in year"),
            ("2026-+8-31T09:30:00", "leading + in month"),
            ("2026-08-31T+9:30:00", "leading + in hour"),
        ];
        for (iso, why) in bad {
            let err = json_to_variant(
                &serde_json::json!({"$type": "time", "iso8601": iso}),
                no_ref_expected,
            )
            .expect_err(&format!("{iso} ({why}) must be rejected"));
            let message = err.to_string();
            assert!(
                message.contains(iso),
                "the error should quote the offending value; got {message:?}"
            );
        }
    }

    #[test]
    fn test_an_object_with_an_unknown_type_tag_still_looks_for_an_ole_ref() {
        // Only "time" is special. Anything else keeps the pre-existing
        // behavior rather than becoming a second silent special case.
        let err = json_to_variant(
            &serde_json::json!({"$type": "something-else"}),
            no_ref_expected,
        )
        .expect_err("an unknown $type is not an $ole_ref and must be rejected");
        assert!(err.to_string().contains("$ole_ref"));
    }
}
