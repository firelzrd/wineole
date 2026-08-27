use std::collections::HashMap;
use std::sync::mpsc::channel;
use std::sync::Mutex;
use serde_json::{json, Value};
use crate::protocol::{Request, Response};
use crate::session::{self, SessionCommand};

pub struct Server {
    pub token: Option<String>,
    routes: Mutex<HashMap<u64, std::sync::mpsc::Sender<SessionCommand>>>,
}

impl Server {
    pub fn new(token: Option<String>) -> Self {
        Server { token, routes: Mutex::new(HashMap::new()) }
    }

    /// Handle one request, discarding the record of any handles it minted.
    /// Only useful where nothing owns the connection's lifetime, i.e. tests —
    /// a real connection must use [`Server::handle_tracked`] so it can release
    /// what it was given when it goes away.
    #[cfg(test)]
    pub fn handle(&self, req: Request, peer_is_loopback: bool, authenticated: &mut bool) -> Response {
        let mut ignored = Vec::new();
        self.handle_tracked(req, peer_is_loopback, authenticated, &mut ignored)
    }

    /// Handle one request, appending every handle id newly registered on the
    /// caller's behalf to `new_handles`.
    ///
    /// The caller (a connection) accumulates these across its whole lifetime so
    /// it can hand them back to [`Server::release_all`] on teardown. Reporting
    /// the ids structurally — rather than having the connection re-parse the
    /// response JSON looking for `$ole_ref` — keeps the two in step no matter
    /// how the result is shaped or nested.
    pub fn handle_tracked(
        &self,
        req: Request,
        peer_is_loopback: bool,
        authenticated: &mut bool,
        new_handles: &mut Vec<u64>,
    ) -> Response {
        if req.method == "ping" {
            let ok = match &self.token {
                Some(t) if !peer_is_loopback => {
                    req.params.get("token").and_then(|v| v.as_str()) == Some(t.as_str())
                }
                _ => true,
            };
            *authenticated = ok;
            return if ok {
                Response::ok(req.id, json!({"pong": true}))
            } else {
                Response::err(req.id, "WineOLE::AuthError", "invalid or missing token")
            };
        }

        if !*authenticated {
            return Response::err(req.id, "WineOLE::AuthError", "call ping with a valid token first");
        }

        match req.method.as_str() {
            "create" | "connect" | "connect_or_create" => self.handle_create_or_connect(&req, new_handles),
            "invoke" => self.handle_invoke(&req, new_handles),
            "const_load" => self.handle_const_load(&req),
            "release" => self.handle_release(&req),
            other => Response::err(req.id, "WineOLE::ProtocolError", &format!("unknown method {}", other)),
        }
    }

    /// Release a batch of handles server-side, with no RPC wrapping — the
    /// connection-teardown counterpart of the `release` method.
    ///
    /// Unknown/already-released ids are ignored, and a session whose worker
    /// thread has already exited simply fails to accept the command; either
    /// way the route is dropped, so this also prunes routing entries that
    /// `handle_release` never saw.
    /// Released in reverse order of acquisition, so the objects obtained *from*
    /// a root are let go before the root itself — the order a well-behaved
    /// client would use, and the one COM servers are used to seeing.
    pub fn release_all(&self, handles: &[u64]) {
        for &handle in handles.iter().rev() {
            self.release_one(handle);
        }
    }

    fn release_one(&self, handle: u64) {
        let sender = self.routes.lock().unwrap().remove(&handle);
        if let Some(sender) = sender {
            let (reply_tx, reply_rx) = channel();
            if sender.send(SessionCommand::Release { handle, reply: reply_tx }).is_ok() {
                // Wait for the acknowledgement: on the release that empties a
                // session's handle table this only arrives once the worker has
                // dropped its receiver, uninitialized its apartment, and
                // released every IDispatch — i.e. once the automated process
                // has actually been let go.
                let _ = reply_rx.recv();
            }
        }
    }

    fn handle_create_or_connect(&self, req: &Request, new_handles: &mut Vec<u64>) -> Response {
        let class_name = match req.params.get("class_name").and_then(|v| v.as_str()) {
            Some(s) => s.to_string(),
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing class_name"),
        };
        let spawned = match req.method.as_str() {
            "create" => session::spawn_create(class_name),
            "connect" => session::spawn_connect(class_name),
            _ => session::spawn_connect_or_create(class_name),
        };
        match spawned {
            Ok(handle) => {
                let root_handle_id = handle.root_handle_id;
                self.routes.lock().unwrap().insert(root_handle_id, handle.sender);
                new_handles.push(root_handle_id);
                // "created" is only meaningful (and only sent) for
                // connect_or_create -- create/connect's caller already knows
                // the answer deterministically from which method it called,
                // so their wire response shape stays exactly as before.
                let mut result = json!({"$ole_ref": root_handle_id});
                if req.method == "connect_or_create" {
                    result["created"] = json!(handle.created);
                }
                Response::ok(req.id, result)
            }
            // `ComError`'s Display always names the HRESULT, so an
            // unresolvable-message failure (a bad ProgID, typically) no longer
            // reaches the client as `WIN32OLERuntimeError: ""`.
            Err(e) => Response::err(req.id, "WIN32OLERuntimeError", &e.to_string()),
        }
    }

    fn handle_invoke(&self, req: &Request, new_handles: &mut Vec<u64>) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        let name = req.params.get("name").and_then(|v| v.as_str()).unwrap_or("").to_string();
        let args: Vec<Value> = req.params.get("args").and_then(|v| v.as_array()).cloned().unwrap_or_default();
        let named: HashMap<String, Value> = req
            .params
            .get("named")
            .and_then(|v| v.as_object())
            .map(|m| m.iter().map(|(k, v)| (k.clone(), v.clone())).collect())
            .unwrap_or_default();

        let sender = {
            let routes = self.routes.lock().unwrap();
            match routes.get(&handle) {
                Some(s) => s.clone(),
                None => {
                    return Response::err(req.id, "WineOLE::StaleReferenceError", &format!("unknown handle {}", handle))
                }
            }
        };

        let (reply_tx, reply_rx) = channel();
        if sender.send(SessionCommand::Invoke { handle, name, args, named, reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        match reply_rx.recv() {
            Ok(Ok((json_result, new_ids))) => {
                let mut routes = self.routes.lock().unwrap();
                for id in new_ids {
                    routes.insert(id, sender.clone());
                    new_handles.push(id);
                }
                Response::ok(req.id, json_result)
            }
            Ok(Err((class, message))) => Response::err(req.id, &class, &message),
            Err(_) => Response::err(req.id, "WineOLE::ProtocolError", "session did not reply"),
        }
    }

    fn handle_const_load(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };

        let sender = {
            let routes = self.routes.lock().unwrap();
            match routes.get(&handle) {
                Some(s) => s.clone(),
                None => {
                    return Response::err(req.id, "WineOLE::StaleReferenceError", &format!("unknown handle {}", handle))
                }
            }
        };

        let (reply_tx, reply_rx) = channel();
        if sender.send(SessionCommand::ConstLoad { handle, reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        match reply_rx.recv() {
            Ok(Ok(value)) => Response::ok(req.id, value),
            Ok(Err((class, message))) => Response::err(req.id, &class, &message),
            Err(_) => Response::err(req.id, "WineOLE::ProtocolError", "session did not reply"),
        }
    }

    fn handle_release(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        self.release_one(handle);
        Response::ok(req.id, json!(null))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn req(id: u64, method: &str, params: Value) -> Request {
        serde_json::from_value(json!({"id": id, "method": method, "params": params})).unwrap()
    }

    #[test]
    fn test_ping_without_token_is_always_ok() {
        let server = Server::new(None);
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({})), false, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_ping_with_wrong_token_from_non_loopback_is_rejected() {
        let server = Server::new(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({"token": "wrong"})), false, &mut authenticated);
        assert!(resp.error.is_some());
        assert!(!authenticated);
    }

    #[test]
    fn test_ping_with_correct_token_from_non_loopback_is_accepted() {
        let server = Server::new(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({"token": "secret"})), false, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_ping_with_no_token_from_loopback_is_ok_even_when_token_configured() {
        let server = Server::new(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({})), true, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_invoke_before_authenticated_is_rejected() {
        let server = Server::new(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "invoke", json!({"handle": 1, "name": "Foo"})), false, &mut authenticated);
        assert_eq!(resp.error.as_ref().unwrap().class, "WineOLE::AuthError");
    }

    #[test]
    fn test_invoke_on_unknown_handle_is_stale_reference_error() {
        let server = Server::new(None);
        let mut authenticated = true;
        let resp = server.handle(req(1, "invoke", json!({"handle": 999, "name": "Foo"})), true, &mut authenticated);
        assert_eq!(resp.error.as_ref().unwrap().class, "WineOLE::StaleReferenceError");
    }

    #[test]
    fn test_create_invoke_release_against_real_excel() {
        let server = Server::new(None);
        let mut authenticated = true;

        let create_resp = server.handle(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
        );
        let handle = create_resp.result.unwrap()["$ole_ref"].as_u64().unwrap();

        let invoke_resp = server.handle(
            req(2, "invoke", json!({"handle": handle, "name": "Visible=", "args": [false]})),
            true,
            &mut authenticated,
        );
        assert!(invoke_resp.error.is_none(), "{:?}", invoke_resp.error);

        let quit_resp = server.handle(
            req(3, "invoke", json!({"handle": handle, "name": "Quit"})),
            true,
            &mut authenticated,
        );
        assert!(quit_resp.error.is_none(), "{:?}", quit_resp.error);

        let release_resp = server.handle(req(4, "release", json!({"handle": handle})), true, &mut authenticated);
        assert!(release_resp.error.is_none());
    }

    #[test]
    fn test_com_error_message_includes_the_hresult() {
        // A ProgID nobody has registered: `windows::core::Error::message()` is
        // empty for this HRESULT under Wine, so without the code appended the
        // client used to receive `WIN32OLERuntimeError: ""`.
        let server = Server::new(None);
        let mut authenticated = true;
        let resp = server.handle(
            req(1, "create", json!({"class_name": "NoSuch.ProgID.WineOLE"})),
            true,
            &mut authenticated,
        );
        let error = resp.error.expect("creating a bogus ProgID must fail");
        assert_eq!(error.class, "WIN32OLERuntimeError");
        assert!(
            error.message.contains("(0x"),
            "error message should name the HRESULT, got {:?}",
            error.message
        );
        assert!(
            error.message.len() > "(0x00000000)".len(),
            "error message should never be effectively empty, got {:?}",
            error.message
        );
    }

    #[test]
    fn test_connection_teardown_releases_every_handle_and_shuts_the_session_down() {
        let server = Server::new(None);
        let mut authenticated = true;
        // Stands in for a connection's accumulated handle set in main.rs.
        let mut owned: Vec<u64> = Vec::new();

        let create_resp = server.handle_tracked(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(create_resp.error.is_none(), "{:?}", create_resp.error);
        let root = create_resp.result.unwrap()["$ole_ref"].as_u64().unwrap();
        assert_eq!(owned, vec![root], "create must report its root handle");

        let hide = server.handle_tracked(
            req(2, "invoke", json!({"handle": root, "name": "Visible=", "args": [false]})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(hide.error.is_none(), "{:?}", hide.error);

        // An invoke returning an object mints a handle the client never
        // explicitly releases — exactly the leak this test pins down.
        let workbooks = server.handle_tracked(
            req(3, "invoke", json!({"handle": root, "name": "Workbooks"})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(workbooks.error.is_none(), "{:?}", workbooks.error);
        assert_eq!(owned.len(), 2, "the Workbooks object handle must be tracked too");

        // Keep a direct line to the session so we can observe its worker
        // thread exiting (the routing table is emptied by the teardown).
        let session_sender = server.routes.lock().unwrap().get(&root).unwrap().clone();

        let quit = server.handle_tracked(
            req(4, "invoke", json!({"handle": root, "name": "Quit"})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(quit.error.is_none(), "{:?}", quit.error);

        // The connection goes away without ever sending `release`.
        server.release_all(&owned);

        assert!(
            server.routes.lock().unwrap().is_empty(),
            "teardown must leave no stale routing entries behind"
        );

        // Sending to the session now fails, which is only possible once the
        // worker thread has dropped its receiver and exited — mirroring
        // session.rs's own shutdown assertion.
        let (tx, _rx) = channel();
        assert!(
            session_sender
                .send(SessionCommand::Release { handle: root, reply: tx })
                .is_err(),
            "the session worker thread should have exited on connection teardown"
        );
    }

    #[test]
    fn test_const_load_rpc_returns_excel_constants() {
        let server = Server::new(None);
        let mut authenticated = true;
        let mut owned = Vec::new();

        let create_resp = server.handle_tracked(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
        );
        let handle = create_resp.result.unwrap()["$ole_ref"].as_u64().unwrap();

        let cl_resp = server.handle_tracked(
            req(2, "const_load", json!({"handle": handle})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(cl_resp.error.is_none(), "{:?}", cl_resp.error);
        let consts = cl_resp.result.unwrap();
        assert_eq!(consts["xlDown"].as_i64(), Some(-4121));

        let quit = server.handle_tracked(
            req(3, "invoke", json!({"handle": handle, "name": "Quit"})),
            true,
            &mut authenticated,
            &mut owned,
        );
        assert!(quit.error.is_none());

        server.release_all(&owned);
    }

    #[test]
    fn test_connect_or_create_response_includes_created_flag_but_create_and_connect_do_not() {
        let server = Server::new(None);
        let mut authenticated = true;

        let create_resp = server.handle(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
        );
        let create_result = create_resp.result.expect("create should succeed");
        assert!(create_result.get("created").is_none(), "create's response must not include a created field");
        let handle = create_result["$ole_ref"].as_u64().unwrap();

        let coc_resp = server.handle(
            req(2, "connect_or_create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
        );
        let coc_result = coc_resp.result.expect("connect_or_create should succeed");
        assert_eq!(coc_result["created"], json!(false), "connect_or_create must report created=false when something was already running");
        let coc_handle = coc_result["$ole_ref"].as_u64().unwrap();

        let connect_resp = server.handle(
            req(3, "connect", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
        );
        let connect_result = connect_resp.result.expect("connect should succeed");
        assert!(connect_result.get("created").is_none(), "connect's response must not include a created field");
        let connect_handle = connect_result["$ole_ref"].as_u64().unwrap();

        for h in [handle, coc_handle, connect_handle] {
            let _ = server.handle(req(4, "invoke", json!({"handle": h, "name": "Quit"})), true, &mut authenticated);
        }
    }
}
