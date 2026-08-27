use std::collections::HashMap;
use std::sync::atomic::{AtomicU32, Ordering};
use std::sync::mpsc::{channel, Sender};
use serde_json::Value;
use windows::Win32::System::Com::{CoInitializeEx, CoUninitialize, COINIT_APARTMENTTHREADED, IDispatch};
use crate::dispatch::{ComError, ComResult};
use crate::{dispatch, value};

pub enum SessionCommand {
    Invoke {
        handle: u64,
        name: String,
        args: Vec<Value>,
        named: HashMap<String, Value>,
        reply: Sender<Result<(Value, Vec<u64>), (String, String)>>,
    },
    ConstLoad {
        handle: u64,
        reply: Sender<Result<Value, (String, String)>>,
    },
    Release {
        handle: u64,
        reply: Sender<()>,
    },
}

pub struct SessionHandle {
    pub sender: Sender<SessionCommand>,
    pub root_handle_id: u64,
    pub created: bool,
}

static SESSION_SEQ: AtomicU32 = AtomicU32::new(1);

enum SpawnMode {
    Create,
    Connect,
    ConnectOrCreate,
}

pub fn spawn_create(progid: String) -> ComResult<SessionHandle> {
    spawn(progid, SpawnMode::Create)
}

pub fn spawn_connect(progid: String) -> ComResult<SessionHandle> {
    spawn(progid, SpawnMode::Connect)
}

pub fn spawn_connect_or_create(progid: String) -> ComResult<SessionHandle> {
    spawn(progid, SpawnMode::ConnectOrCreate)
}

fn spawn(progid: String, mode: SpawnMode) -> ComResult<SessionHandle> {
    let session_seq = SESSION_SEQ.fetch_add(1, Ordering::SeqCst) as u64;
    let (cmd_tx, cmd_rx) = channel::<SessionCommand>();
    let (ready_tx, ready_rx) = channel::<ComResult<(u64, bool)>>();

    std::thread::spawn(move || {
        unsafe {
            if let Err(e) = CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok() {
                let _ = ready_tx.send(Err(ComError::from(e)));
                return;
            }
        }

        let root = match mode {
            SpawnMode::Create => dispatch::create_instance(&progid).map(|d| (d, true)),
            SpawnMode::Connect => dispatch::get_active_object(&progid).map(|d| (d, false)),
            SpawnMode::ConnectOrCreate => dispatch::connect_or_create(&progid),
        };

        let mut handles: HashMap<u64, IDispatch> = HashMap::new();
        let mut local_seq: u32 = 0;

        let (root_disp, created) = match root {
            Ok(pair) => pair,
            Err(e) => {
                let _ = ready_tx.send(Err(e));
                unsafe { CoUninitialize(); }
                return;
            }
        };

        let root_id = (session_seq << 32) | (local_seq as u64);
        local_seq += 1;
        handles.insert(root_id, root_disp);
        if ready_tx.send(Ok((root_id, created))).is_err() {
            handles.clear();
            unsafe { CoUninitialize(); }
            return;
        }

        // On the release that empties `handles`, don't reply on the spot: the
        // reply channel wakes the caller, who may immediately send another
        // command, and that send only fails once `cmd_rx` (moved into this
        // `for` loop's iterator) is actually dropped. That drop happens when
        // control leaves the loop, which is *after* a reply sent from inside
        // the loop body would already be in flight — a race the caller's
        // thread reliably wins, making the "further sends fail" test
        // assertion flaky/false. Deferring the reply until after the loop
        // (and thus after `cmd_rx` is dropped) makes the shutdown ordering a
        // real happens-before relationship instead of a race: the caller
        // can't observe the reply until the receiver is already gone.
        let mut shutdown_reply: Option<Sender<()>> = None;
        for cmd in cmd_rx {
            match cmd {
                SessionCommand::Invoke { handle, name, args, named, reply } => {
                    let outcome = run_invoke(&mut handles, &mut local_seq, session_seq, handle, &name, &args, &named);
                    let _ = reply.send(outcome);
                }
                SessionCommand::ConstLoad { handle, reply } => {
                    let outcome = run_const_load(&handles, handle);
                    let _ = reply.send(outcome);
                }
                SessionCommand::Release { handle, reply } => {
                    handles.remove(&handle);
                    if handles.is_empty() {
                        shutdown_reply = Some(reply);
                        break;
                    } else {
                        let _ = reply.send(());
                    }
                }
            }
        }

        // Drop (and thus Release) every surviving IDispatch proxy while the
        // apartment is still initialized. If the loop exited because all
        // senders were dropped (caller disconnected without releasing every
        // handle) rather than via the empty-`handles` Release path, `handles`
        // may still hold live COM pointers here; releasing them after
        // CoUninitialize would violate COM's contract and can leak the
        // underlying process (e.g. a lingering EXCEL.EXE).
        handles.clear();
        unsafe { CoUninitialize(); }
        if let Some(reply) = shutdown_reply {
            let _ = reply.send(());
        }
    });

    let (root_id, created) = ready_rx
        .recv()
        .map_err(|_| {
            ComError::new(
                windows::Win32::Foundation::E_FAIL,
                "session worker thread died before reporting readiness",
            )
        })??;
    Ok(SessionHandle { sender: cmd_tx, root_handle_id: root_id, created })
}

fn run_invoke(
    handles: &mut HashMap<u64, IDispatch>,
    local_seq: &mut u32,
    session_seq: u64,
    handle: u64,
    name: &str,
    args: &[Value],
    named: &HashMap<String, Value>,
) -> Result<(Value, Vec<u64>), (String, String)> {
    let disp = handles
        .get(&handle)
        .cloned()
        .ok_or_else(|| ("WineOLE::StaleReferenceError".to_string(), format!("unknown handle {}", handle)))?;

    // `ComError`'s Display always renders the HRESULT, so even the (common
    // under Wine) case of an empty OS message still tells the client which
    // HRESULT failed instead of arriving as `WIN32OLERuntimeError: ""`.
    let positional = args
        .iter()
        .map(|v| value::json_to_variant(v, |id| resolve(handles, id)))
        .collect::<windows::core::Result<Vec<_>>>()
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), ComError::from(e).to_string()))?;

    let named_pairs = named
        .iter()
        .map(|(k, v)| value::json_to_variant(v, |id| resolve(handles, id)).map(|variant| (k.clone(), variant)))
        .collect::<windows::core::Result<Vec<_>>>()
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), ComError::from(e).to_string()))?;

    let result_variant = dispatch::invoke_member(&disp, name, positional, named_pairs)
        .map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let mut new_ids = Vec::new();
    let json_result = value::variant_to_json(&result_variant, |new_disp| {
        let id = (session_seq << 32) | (*local_seq as u64);
        *local_seq += 1;
        handles.insert(id, new_disp);
        new_ids.push(id);
        id
    });

    Ok((json_result, new_ids))
}

fn run_const_load(handles: &HashMap<u64, IDispatch>, handle: u64) -> Result<Value, (String, String)> {
    let disp = handles
        .get(&handle)
        .cloned()
        .ok_or_else(|| ("WineOLE::StaleReferenceError".to_string(), format!("unknown handle {}", handle)))?;

    let pairs = dispatch::const_load(&disp).map_err(|e| ("WIN32OLERuntimeError".to_string(), e.to_string()))?;

    let mut map = serde_json::Map::new();
    for (name, variant) in pairs {
        // A constant is never an object reference — this closure existing to
        // satisfy variant_to_json's signature and panicking if it's ever
        // actually called is intentional: it would mean a type library
        // exposed an enum member holding a live COM object, which is not
        // meaningful and should surface loudly during development rather
        // than silently minting an orphan handle nobody will release.
        let value = value::variant_to_json(&variant, |_| {
            panic!("const_load: enum constant '{}' unexpectedly held an object reference", name)
        });
        map.insert(name, value);
    }
    Ok(Value::Object(map))
}

fn resolve(handles: &HashMap<u64, IDispatch>, id: u64) -> windows::core::Result<IDispatch> {
    handles
        .get(&id)
        .cloned()
        .ok_or_else(|| windows::core::Error::from_hresult(windows::Win32::Foundation::E_INVALIDARG))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_invoke_named_arg_and_release_shuts_down_session() {
        let session = spawn_create("Excel.Application".to_string()).expect("spawn_create");

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Visible=".to_string(),
                args: vec![serde_json::json!(false)],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        rx.recv().unwrap().expect("Visible=false should succeed");

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Workbooks".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        let (workbooks_json, new_ids) = rx.recv().unwrap().expect("Workbooks should succeed");
        assert_eq!(new_ids.len(), 1);
        let workbooks_handle = workbooks_json["$ole_ref"].as_u64().unwrap();

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: workbooks_handle,
                name: "Add".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        // Workbooks.Add returns the newly created Workbook, which run_invoke
        // registers as its own handle (new_ids.len() == 1) since the return
        // value is VT_DISPATCH. That handle must be released too, or
        // `handles` never becomes empty and the session never shuts down.
        let (_add_json, add_new_ids) = rx.recv().unwrap().expect("Workbooks.Add should succeed");
        assert_eq!(add_new_ids.len(), 1);
        let new_workbook_handle = add_new_ids[0];

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Quit".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        // Releasing every outstanding handle should let the worker thread exit.
        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: new_workbook_handle, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: workbooks_handle, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: session.root_handle_id, reply: tx }).unwrap();
        rx.recv().unwrap();

        // The worker thread has exited; further sends fail because the receiver is dropped.
        let (tx, _rx) = channel();
        let send_result = session.sender.send(SessionCommand::Release { handle: session.root_handle_id, reply: tx });
        assert!(send_result.is_err(), "worker thread should have exited after last handle released");
    }

    #[test]
    fn test_const_load_command_returns_excel_constants() {
        let session = spawn_create("Excel.Application".to_string()).expect("spawn_create");

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::ConstLoad { handle: session.root_handle_id, reply: tx })
            .unwrap();
        let constants = rx.recv().unwrap().expect("const_load should succeed");
        let obj = constants.as_object().expect("const_load result must be a JSON object");
        assert_eq!(obj.get("xlUp").and_then(|v| v.as_i64()), Some(-4162));

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Quit".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: session.root_handle_id, reply: tx }).unwrap();
        rx.recv().unwrap();
    }

    #[test]
    fn test_spawn_create_reports_created_true() {
        let session = spawn_create("Excel.Application".to_string()).expect("spawn_create");
        assert!(session.created, "spawn_create must report created=true");

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Quit".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: session.root_handle_id, reply: tx }).unwrap();
        rx.recv().unwrap();
    }

    #[test]
    fn test_spawn_connect_or_create_attaches_to_an_existing_session() {
        let creator = spawn_create("Excel.Application".to_string()).expect("spawn_create");

        let session = spawn_connect_or_create("Excel.Application".to_string()).expect("spawn_connect_or_create");
        assert!(!session.created, "spawn_connect_or_create must report created=false when something was already running");

        let (tx, rx) = channel();
        session
            .sender
            .send(SessionCommand::Invoke {
                handle: session.root_handle_id,
                name: "Quit".to_string(),
                args: vec![],
                named: HashMap::new(),
                reply: tx,
            })
            .unwrap();
        rx.recv().unwrap().expect("Quit should succeed");

        let (tx, rx) = channel();
        session.sender.send(SessionCommand::Release { handle: session.root_handle_id, reply: tx }).unwrap();
        rx.recv().unwrap();

        let (tx, rx) = channel();
        creator.sender.send(SessionCommand::Release { handle: creator.root_handle_id, reply: tx }).unwrap();
        rx.recv().unwrap();
    }
}
