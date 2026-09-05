//! Cross-apartment identity for a COM instance.
//!
//! Two session threads live in two STA apartments. IUnknown pointer equality
//! is meaningless across that boundary (a proxy in one apartment is a
//! different pointer than the object in another -- a COM rule, not a Wine
//! quirk). Measured under Wine with Excel 2003: marshaling the root proxy
//! writes the SERVER's STDOBJREF, whose `(OXID, OID)` pair is equal across
//! apartments for the same EXCEL.EXE and differs between two of them. That
//! pair is the identity key. OID alone is not enough -- it was observed to be
//! `2` for two distinct Excels -- so the key is both fields together.
use windows::core::Interface;
use windows::Win32::System::Com::{IDispatch, IStream, MSHCTX_LOCAL, MSHLFLAGS_NORMAL, STREAM_SEEK_SET};
use windows::Win32::System::Com::Marshal::{CoMarshalInterface, CoUnmarshalInterface};
use windows::Win32::System::Com::StructuredStorage::CreateStreamOnHGlobal;
use windows::Win32::Foundation::HGLOBAL;
use crate::dispatch::{ComError, ComResult};

#[derive(Clone, Copy, PartialEq, Eq, Hash, Debug)]
pub struct InstanceKey {
    pub oxid: u64,
    pub oid: u64,
}

/// A per-session key that cannot collide with a real `(OXID, OID)`.
///
/// Real OIDs are small (Excel's was `2`), so `u64::MAX` in the OID slot is a
/// safe sentinel. Used when marshaling fails: the session is then treated as
/// the sole user of its own instance -- today's 1:1 behavior -- rather than
/// crashing.
pub fn synthetic_key(session_seq: u64) -> InstanceKey {
    InstanceKey { oxid: session_seq, oid: u64::MAX }
}

/// Marshal `disp` and read `(OXID, OID)` from the OBJREF the server wrote.
///
/// The packet is consumed with `CoUnmarshalInterface`, not
/// `CoReleaseMarshalData`: the latter answers `RPC_E_INVALID_OBJREF`
/// (0x8001011D) for a proxy-written packet under Wine. Leaving the packet
/// unconsumed would leak the marshaled reference (and so keep EXCEL.EXE
/// alive), so the unmarshal is not optional.
pub fn instance_key(disp: &IDispatch) -> ComResult<InstanceKey> {
    let unk: windows::core::IUnknown = disp.cast().map_err(ComError::from)?;
    let stream: IStream = unsafe { CreateStreamOnHGlobal(HGLOBAL::default(), true) }.map_err(ComError::from)?;
    unsafe {
        CoMarshalInterface(&stream, &IDispatch::IID, &unk, MSHCTX_LOCAL.0 as u32, None, MSHLFLAGS_NORMAL.0 as u32)
            .map_err(ComError::from)?;
        stream.Seek(0, STREAM_SEEK_SET, None).map_err(ComError::from)?;
    }
    let mut buf = vec![0u8; 512];
    let mut read = 0u32;
    unsafe {
        stream
            .Read(buf.as_mut_ptr() as *mut _, buf.len() as u32, Some(&mut read))
            .ok()
            .map_err(ComError::from)?;
    }
    buf.truncate(read as usize);
    // Consume the packet unconditionally, before any validation, so that no
    // early return below can skip it (see the doc comment on why this is
    // CoUnmarshalInterface, not CoReleaseMarshalData).
    unsafe {
        stream.Seek(0, STREAM_SEEK_SET, None).map_err(ComError::from)?;
        let _consumed: IDispatch = CoUnmarshalInterface(&stream).map_err(ComError::from)?;
    }
    if buf.len() < 48 {
        return Err(ComError::new(
            windows::Win32::Foundation::E_FAIL,
            format!("OBJREF too short ({} bytes) to read OXID/OID", buf.len()),
        ));
    }
    let u32_at = |o: usize| u32::from_le_bytes(buf[o..o + 4].try_into().unwrap());
    // OBJREF: sig(4) flags(4) iid(16) | STDOBJREF: flags(4) cPublicRefs(4) oxid@32(8) oid@40(8) ipid@48(16)
    if u32_at(0) != 0x574f_454d {
        return Err(ComError::new(
            windows::Win32::Foundation::E_FAIL,
            "OBJREF signature was not MEOW; cannot read identity".to_string(),
        ));
    }
    let u64_at = |o: usize| u64::from_le_bytes(buf[o..o + 8].try_into().unwrap());
    let oxid = u64_at(32);
    let oid = u64_at(40);
    Ok(InstanceKey { oxid, oid })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn synthetic_keys_are_unique_per_session_and_never_look_real() {
        let a = synthetic_key(1);
        let b = synthetic_key(2);
        assert_ne!(a, b);
        assert_eq!(a.oid, u64::MAX);
    }

    // Real Excel. Serialized against every other Excel test by the lock.
    #[test]
    fn instance_key_is_equal_across_apartments_and_differs_per_excel() {
        use crate::dispatch::{create_instance, get_active_object, invoke_member, lock_excel_for_test};
        use windows::core::VARIANT;
        use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED};
        use std::sync::mpsc::channel;

        let _guard = lock_excel_for_test();
        let (a_key_tx, a_key_rx) = channel::<InstanceKey>();
        let (b_done_tx, b_done_rx) = channel::<()>();
        let (a2_key_tx, a2_key_rx) = channel::<InstanceKey>();

        let a = std::thread::spawn(move || {
            unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
            let xl1 = create_instance("Excel.Application").expect("create #1");
            invoke_member(&xl1, "Visible=", vec![VARIANT::from(false)], vec![]).unwrap();
            invoke_member(&xl1, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]).unwrap();
            let k1 = instance_key(&xl1).expect("key #1");
            a_key_tx.send(k1).unwrap();
            b_done_rx.recv().unwrap();

            let xl2 = create_instance("Excel.Application").expect("create #2");
            invoke_member(&xl2, "Visible=", vec![VARIANT::from(false)], vec![]).unwrap();
            invoke_member(&xl2, "DisplayAlerts=", vec![VARIANT::from(false)], vec![]).unwrap();
            let k2 = instance_key(&xl2).expect("key #2");
            a2_key_tx.send(k2).unwrap();

            invoke_member(&xl2, "Quit", vec![], vec![]).unwrap();
            invoke_member(&xl1, "Quit", vec![], vec![]).unwrap();
            drop((xl1, xl2));
            unsafe { CoUninitialize() };
        });

        let b = std::thread::spawn(move || {
            unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok().unwrap() };
            let k1_a = a_key_rx.recv().unwrap();
            let xlb = get_active_object("Excel.Application").expect("GAO in B");
            let k1_b = instance_key(&xlb).expect("key from B");
            // Same Excel, different apartment: the key matches.
            assert_eq!(k1_a, k1_b, "identity key must be equal across apartments");
            drop(xlb);
            b_done_tx.send(()).unwrap();

            let k2 = a2_key_rx.recv().unwrap();
            // A second, distinct Excel has a different OXID (Wine sets its high
            // 32 bits to the server PID).
            assert_ne!(k1_a.oxid, k2.oxid, "two Excels must have different OXIDs");
            unsafe { CoUninitialize() };
        });

        a.join().unwrap();
        b.join().unwrap();
    }
}
