use std::collections::HashMap;
use std::sync::mpsc::{channel, Sender};
use std::sync::Mutex;
use serde_json::{json, Value};
use crate::protocol::{self, Request, Response};
use crate::session::{self, ReleaseOutcome, SessionCommand, SessionEvent, SessionRoute};

pub struct Server {
    pub token: Option<String>,
    routes: Mutex<HashMap<u64, SessionRoute>>,
    registry: std::sync::Arc<crate::registry::InstanceRegistry>,
}

impl Server {
    pub fn new(token: Option<String>, registry: std::sync::Arc<crate::registry::InstanceRegistry>) -> Self {
        Server { token, routes: Mutex::new(HashMap::new()), registry }
    }

    /// Handle one request, discarding the record of any handles it minted.
    /// Only useful where nothing owns the connection's lifetime, i.e. tests —
    /// a real connection must use [`Server::handle_tracked`] so it can release
    /// what it was given when it goes away.
    #[cfg(test)]
    pub fn handle(&self, req: Request, peer_is_loopback: bool, authenticated: &mut bool) -> Response {
        let mut ignored = Vec::new();
        // A session still has to be told where its events go; these tests are
        // not about events, so they go to a receiver nobody reads.
        let (events, _rx) = channel();
        self.handle_tracked(req, peer_is_loopback, authenticated, &mut ignored, &events)
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
        events: &Sender<SessionEvent>,
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
            "create" | "connect" | "connect_or_create" => {
                self.handle_create_or_connect(&req, new_handles, events)
            }
            "invoke" => self.handle_invoke(&req, new_handles),
            "const_load" => self.handle_const_load(&req),
            "release" => self.handle_release(&req),
            "leave_open" => self.handle_leave_open(&req),
            "subscribe" => self.handle_subscribe(&req),
            "unsubscribe" => self.handle_unsubscribe(&req),
            "release_event" => self.handle_release_event(&req),
            other => Response::err(req.id, "WineOLE::ProtocolError", &format!("unknown method {}", other)),
        }
    }

    /// Take one event off a session on its way to the wire.
    ///
    /// The frame is returned to be written; what happens here is the part
    /// that is not writing. An event's object arguments were minted inside
    /// the session and have never passed through [`Server::handle_tracked`],
    /// so nothing routes them yet -- and a client that is handed
    /// `{"$ole_ref": n}` in a callback will call methods on it. Registering
    /// them here, BEFORE the frame goes out, is what makes that call arrive
    /// at the session that owns the object rather than at
    /// `StaleReferenceError`.
    pub fn receive_event(&self, ev: SessionEvent) -> protocol::Event {
        if !ev.new_handles.is_empty() {
            // The event's own object is the obvious route, but a client may
            // already have released it while events were in flight; any
            // handle of the same session will do, and says the same thing.
            if let Some(route) = self.route_of_session(ev.frame.handle) {
                let mut routes = self.routes.lock().unwrap();
                for id in &ev.new_handles {
                    routes.insert(*id, route.clone());
                }
            }
        }
        ev.frame
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

        // Event-argument handles were minted on a session thread and reported
        // straight to the connection's event stream, so they never went
        // through `new_handles` and are not in `handles`. Every one of them
        // holds a `SessionRoute` in the routing table, and ONE surviving
        // route is enough to keep a session thread parked forever -- holding
        // the IDispatch proxies that keep EXCEL.EXE alive. So this takes out
        // every entry belonging to a session this connection was releasing,
        // which is what an id's top 32 bits say. Dropping those routes is
        // also what tells the session its connection is gone.
        //
        // Purging by SESSION rather than by id is safe because of an invariant
        // nothing here can check: a session belongs to exactly one connection.
        // It is created by that connection's `create`/`connect`, its route is
        // never handed to another, and only ids it minted carry its number --
        // so no entry this `retain` drops can belong to anybody else.
        let sessions: std::collections::HashSet<u64> = handles.iter().map(|h| h >> 32).collect();
        if !sessions.is_empty() {
            self.routes.lock().unwrap().retain(|id, _| !sessions.contains(&(id >> 32)));
        }
    }

    /// A route to the session that minted `id`.
    ///
    /// Every id this bridge hands out -- object handles and event seqs alike
    /// -- is `(session_seq << 32) | n`, so the session an id came from is
    /// written in its top 32 bits. `release_event` carries nothing but a seq,
    /// and one connection may hold several sessions, so something has to say
    /// which session a seq belongs to. A map from seq to route would answer
    /// the same question and be worse: holding a `SessionRoute` for every
    /// unreleased event keeps the session thread alive after its connection
    /// went away, which is the EXCEL.EXE-outlives-the-bridge failure this
    /// project has already paid for twice.
    fn route_of_session(&self, id: u64) -> Option<SessionRoute> {
        let session = id >> 32;
        self.routes
            .lock()
            .unwrap()
            .iter()
            .find(|(handle, _)| **handle >> 32 == session)
            .map(|(_, route)| route.clone())
    }

    fn release_one(&self, handle: u64) {
        let route = self.routes.lock().unwrap().remove(&handle);
        if let Some(route) = route {
            let (reply_tx, reply_rx) = channel();
            if route.send(SessionCommand::Release { handle, client_alive: false, reply: reply_tx }).is_ok() {
                // Wait for the acknowledgement: on the release that empties a
                // session's handle table this only arrives once the worker has
                // dropped its receiver, uninitialized its apartment, and
                // released every IDispatch — i.e. once the automated process
                // has actually been let go. The outcome itself (a
                // `ReleaseOutcome`) is not yet acted on -- Task 4 wires the
                // last-user cleanup decision.
                let _ = reply_rx.recv();
            }
        }
    }

    fn handle_create_or_connect(
        &self,
        req: &Request,
        new_handles: &mut Vec<u64>,
        events: &Sender<SessionEvent>,
    ) -> Response {
        let class_name = match req.params.get("class_name").and_then(|v| v.as_str()) {
            Some(s) => s.to_string(),
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing class_name"),
        };
        let cleanup = match crate::registry::parse_cleanup(&req.params) {
            Ok(c) => c,
            Err((class, msg)) => return Response::err(req.id, &class, &msg),
        };
        // Every session this connection creates pushes onto the one event
        // stream the connection owns; the frames carry the handle they came
        // from, which is what tells them apart at the other end.
        let events = events.clone();
        let registry = self.registry.clone();
        let spawned = match req.method.as_str() {
            "create" => session::spawn_create(class_name, events, registry, cleanup),
            "connect" => session::spawn_connect(class_name, events, registry, cleanup),
            _ => session::spawn_connect_or_create(class_name, events, registry, cleanup),
        };
        match spawned {
            Ok(handle) => {
                let root_handle_id = handle.root_handle_id;
                self.routes.lock().unwrap().insert(root_handle_id, handle.route.clone());
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
            // reaches the client as `WIN32OLERuntimeError: ""`. A `connect`
            // that landed on an instance being shut down surfaces via
            // `ready_tx` as a `ComError` whose message the session tagged with
            // the `WineOLE::InstanceClosingError:` prefix; map it back to its
            // own wire class here (see session.rs's participation block).
            Err(e) => {
                let msg = e.to_string();
                if let Some(stripped) = msg.strip_prefix("WineOLE::InstanceClosingError: ") {
                    Response::err(req.id, "WineOLE::InstanceClosingError", stripped)
                } else {
                    Response::err(req.id, "WIN32OLERuntimeError", &msg)
                }
            }
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

        let route = match self.route_for(handle) {
            Some(r) => r,
            None => {
                return Response::err(
                    req.id,
                    "WineOLE::StaleReferenceError",
                    &format!("unknown handle {}", handle),
                )
            }
        };

        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::Invoke { handle, name, args, named, reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        match reply_rx.recv() {
            Ok(Ok((json_result, new_ids))) => {
                let mut routes = self.routes.lock().unwrap();
                for id in new_ids {
                    routes.insert(id, route.clone());
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

        let route = match self.route_for(handle) {
            Some(r) => r,
            None => {
                return Response::err(
                    req.id,
                    "WineOLE::StaleReferenceError",
                    &format!("unknown handle {}", handle),
                )
            }
        };

        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::ConstLoad { handle, reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        match reply_rx.recv() {
            Ok(Ok(value)) => Response::ok(req.id, value),
            Ok(Err((class, message))) => Response::err(req.id, &class, &message),
            Err(_) => Response::err(req.id, "WineOLE::ProtocolError", "session did not reply"),
        }
    }

    /// `subscribe(handle, event, args = true) -> true`.
    ///
    /// There is no `advise` method next to this one, and there never will be:
    /// the session derives the COM `Advise` from the subscription itself, so
    /// a client cannot end up registered for an event it does not receive.
    fn handle_subscribe(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        let event = match req.params.get("event").and_then(|v| v.as_str()) {
            Some(s) => s.to_string(),
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing event"),
        };
        // Arguments by default: an event whose objects were never minted is
        // the surprising case, so it is the one the client has to ask for.
        let args = req.params.get("args").and_then(|v| v.as_bool()).unwrap_or(true);

        let route = match self.route_for(handle) {
            Some(r) => r,
            None => {
                return Response::err(
                    req.id,
                    "WineOLE::StaleReferenceError",
                    &format!("unknown handle {}", handle),
                )
            }
        };

        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::Subscribe { handle, event, args, reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        match reply_rx.recv() {
            Ok(Ok(())) => Response::ok(req.id, json!(true)),
            Ok(Err((class, message))) => Response::err(req.id, &class, &message),
            Err(_) => Response::err(req.id, "WineOLE::ProtocolError", "session did not reply"),
        }
    }

    /// `unsubscribe(handle, event) -> bool`, the bool being whether it had
    /// been subscribed. Removing the last name for an object is what takes
    /// the `Advise` back out.
    fn handle_unsubscribe(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        let event = match req.params.get("event").and_then(|v| v.as_str()) {
            Some(s) => s.to_string(),
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing event"),
        };

        // A handle whose session has already gone was not subscribed to
        // anything, and saying so is more useful to a client tidying up than
        // an error it has to rescue.
        let Some(route) = self.route_for(handle) else { return Response::ok(req.id, json!(false)) };

        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::Unsubscribe { handle, event, reply: reply_tx }).is_err() {
            return Response::ok(req.id, json!(false));
        }
        match reply_rx.recv() {
            Ok(was) => Response::ok(req.id, json!(was)),
            Err(_) => Response::err(req.id, "WineOLE::ProtocolError", "session did not reply"),
        }
    }

    /// `release_event(seq) -> true`.
    ///
    /// One event's arguments go together. An unknown seq is not an error --
    /// see the command's arm in the session loop for why.
    fn handle_release_event(&self, req: &Request) -> Response {
        let seq = match req.params.get("seq").and_then(|v| v.as_u64()) {
            Some(s) => s,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing seq"),
        };

        let Some(route) = self.route_of_session(seq) else { return Response::ok(req.id, json!(true)) };

        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::ReleaseEvent { seq, reply: reply_tx }).is_err() {
            return Response::ok(req.id, json!(true));
        }
        if let Ok(released) = reply_rx.recv() {
            // The session has dropped the objects; the routing entries that
            // pointed at them would otherwise outlive them, and each one
            // holds a route that keeps the session thread alive.
            if !released.is_empty() {
                let mut routes = self.routes.lock().unwrap();
                for id in released {
                    routes.remove(&id);
                }
            }
        }
        Response::ok(req.id, json!(true))
    }

    fn route_for(&self, handle: u64) -> Option<SessionRoute> {
        self.routes.lock().unwrap().get(&handle).cloned()
    }

    fn handle_release(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        let route = match self.route_for(handle) {
            Some(r) => r,
            None => return Response::ok(req.id, json!(null)),
        };
        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::Release { handle, client_alive: true, reply: reply_tx }).is_err() {
            self.routes.lock().unwrap().remove(&handle);
            return Response::ok(req.id, json!(null));
        }
        match reply_rx.recv() {
            Ok(ReleaseOutcome::Done) => {
                self.routes.lock().unwrap().remove(&handle);
                Response::ok(req.id, json!(null))
            }
            Ok(ReleaseOutcome::CleanupPending(seq)) => {
                // Keep the route: the session is still alive and must receive
                // the client's release_event for this seq.
                Response::ok(req.id, json!({ "cleanup": seq }))
            }
            Err(_) => {
                self.routes.lock().unwrap().remove(&handle);
                Response::ok(req.id, json!(null))
            }
        }
    }

    fn handle_leave_open(&self, req: &Request) -> Response {
        let handle = match req.params.get("handle").and_then(|v| v.as_u64()) {
            Some(h) => h,
            None => return Response::err(req.id, "WineOLE::ProtocolError", "missing handle"),
        };
        let route = match self.route_for(handle) {
            Some(r) => r,
            None => {
                return Response::err(req.id, "WineOLE::StaleReferenceError", &format!("unknown handle {}", handle))
            }
        };
        let (reply_tx, reply_rx) = channel();
        if route.send(SessionCommand::LeaveOpen { reply: reply_tx }).is_err() {
            return Response::err(req.id, "WineOLE::StaleReferenceError", "session is gone");
        }
        let _ = reply_rx.recv();
        Response::ok(req.id, json!(null))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn req(id: u64, method: &str, params: Value) -> Request {
        serde_json::from_value(json!({"id": id, "method": method, "params": params})).unwrap()
    }

    #[cfg(test)]
    fn test_server(token: Option<String>) -> Server {
        Server::new(token, std::sync::Arc::new(crate::registry::InstanceRegistry::new()))
    }

    #[test]
    fn test_ping_without_token_is_always_ok() {
        let server = test_server(None);
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({})), false, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_ping_with_wrong_token_from_non_loopback_is_rejected() {
        let server = test_server(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({"token": "wrong"})), false, &mut authenticated);
        assert!(resp.error.is_some());
        assert!(!authenticated);
    }

    #[test]
    fn test_ping_with_correct_token_from_non_loopback_is_accepted() {
        let server = test_server(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({"token": "secret"})), false, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_ping_with_no_token_from_loopback_is_ok_even_when_token_configured() {
        let server = test_server(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "ping", json!({})), true, &mut authenticated);
        assert!(resp.error.is_none());
        assert!(authenticated);
    }

    #[test]
    fn test_invoke_before_authenticated_is_rejected() {
        let server = test_server(Some("secret".to_string()));
        let mut authenticated = false;
        let resp = server.handle(req(1, "invoke", json!({"handle": 1, "name": "Foo"})), false, &mut authenticated);
        assert_eq!(resp.error.as_ref().unwrap().class, "WineOLE::AuthError");
    }

    #[test]
    fn test_invoke_on_unknown_handle_is_stale_reference_error() {
        let server = test_server(None);
        let mut authenticated = true;
        let resp = server.handle(req(1, "invoke", json!({"handle": 999, "name": "Foo"})), true, &mut authenticated);
        assert_eq!(resp.error.as_ref().unwrap().class, "WineOLE::StaleReferenceError");
    }

    #[test]
    fn test_create_invoke_release_against_real_excel() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
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
        let server = test_server(None);
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
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
        let mut authenticated = true;
        // Stands in for a connection's accumulated handle set in main.rs.
        let mut owned: Vec<u64> = Vec::new();
        // ... and for its event stream. Nothing here subscribes, so nothing
        // arrives on it; it only has to outlive the session.
        let (events, _event_rx) = channel::<SessionEvent>();

        let create_resp = server.handle_tracked(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        assert!(create_resp.error.is_none(), "{:?}", create_resp.error);
        let root = create_resp.result.unwrap()["$ole_ref"].as_u64().unwrap();
        assert_eq!(owned, vec![root], "create must report its root handle");

        let hide = server.handle_tracked(
            req(2, "invoke", json!({"handle": root, "name": "Visible=", "args": [false]})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        assert!(hide.error.is_none(), "{:?}", hide.error);

        // An invoke returning an object mints a handle the client never
        // explicitly releases — exactly the leak this test pins down.
        let workbooks = server.handle_tracked(
            req(3, "invoke", json!({"handle": root, "name": "Workbooks"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        assert!(workbooks.error.is_none(), "{:?}", workbooks.error);
        assert_eq!(owned.len(), 2, "the Workbooks object handle must be tracked too");

        // Keep a direct line to the session so we can observe its worker
        // thread exiting (the routing table is emptied by the teardown).
        let session_route = server.routes.lock().unwrap().get(&root).unwrap().clone();

        let quit = server.handle_tracked(
            req(4, "invoke", json!({"handle": root, "name": "Quit"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
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
            session_route
                .send(SessionCommand::Release { handle: root, client_alive: true, reply: tx })
                .is_err(),
            "the session worker thread should have exited on connection teardown"
        );
    }

    #[test]
    fn test_const_load_rpc_returns_excel_constants() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
        let mut authenticated = true;
        let mut owned = Vec::new();
        let (events, _event_rx) = channel::<SessionEvent>();

        let create_resp = server.handle_tracked(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        let handle = create_resp.result.unwrap()["$ole_ref"].as_u64().unwrap();

        let cl_resp = server.handle_tracked(
            req(2, "const_load", json!({"handle": handle})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        assert!(cl_resp.error.is_none(), "{:?}", cl_resp.error);
        let consts = cl_resp.result.unwrap();
        assert_eq!(consts["xlDown"].as_i64(), Some(-4121));

        let quit = server.handle_tracked(
            req(3, "invoke", json!({"handle": handle, "name": "Quit"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        assert!(quit.error.is_none());

        server.release_all(&owned);
    }

    #[test]
    fn test_connect_or_create_response_includes_created_flag_but_create_and_connect_do_not() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
        let mut authenticated = true;
        // Tracked, like every other Excel-driving test here. `Quit` alone is
        // not a teardown: it leaves this connection's handles in the routing
        // table, and ONE surviving route keeps a session thread parked
        // forever, holding the IDispatch proxies that keep EXCEL.EXE alive.
        // Measured on the untracked version of this test: an EXCEL.EXE
        // outlived the run, and every later test that needs "nothing is
        // running" tripped over it.
        let mut owned = Vec::new();
        let (events, _event_rx) = channel::<SessionEvent>();

        let create_resp = server.handle_tracked(
            req(1, "create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        let create_result = create_resp.result.expect("create should succeed");
        assert!(create_result.get("created").is_none(), "create's response must not include a created field");
        let handle = create_result["$ole_ref"].as_u64().unwrap();

        let coc_resp = server.handle_tracked(
            req(2, "connect_or_create", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        let coc_result = coc_resp.result.expect("connect_or_create should succeed");
        assert_eq!(coc_result["created"], json!(false), "connect_or_create must report created=false when something was already running");
        let coc_handle = coc_result["$ole_ref"].as_u64().unwrap();

        let connect_resp = server.handle_tracked(
            req(3, "connect", json!({"class_name": "Excel.Application"})),
            true,
            &mut authenticated,
            &mut owned,
            &events,
        );
        let connect_result = connect_resp.result.expect("connect should succeed");
        assert!(connect_result.get("created").is_none(), "connect's response must not include a created field");
        let connect_handle = connect_result["$ole_ref"].as_u64().unwrap();

        for h in [handle, coc_handle, connect_handle] {
            let _ = server.handle_tracked(
                req(4, "invoke", json!({"handle": h, "name": "Quit"})),
                true,
                &mut authenticated,
                &mut owned,
                &events,
            );
        }

        // What the connection going away would do, and the half `Quit` cannot
        // do for itself: drop the routes, so the session threads exit and let
        // go of their proxies.
        server.release_all(&owned);
    }

    /// One RPC on a server, made the way a connection makes it. The event
    /// tests below walk a chain of Excel objects, and spelling out five
    /// arguments per step would bury what each step is for.
    fn rpc(
        server: &Server,
        authenticated: &mut bool,
        owned: &mut Vec<u64>,
        events: &Sender<SessionEvent>,
        method: &str,
        params: Value,
    ) -> Response {
        // The id is echoed back and nothing here waits on one, so a counter
        // would add noise without adding meaning.
        server.handle_tracked(req(1, method, params), true, authenticated, owned, events)
    }

    fn ok_result(resp: Response) -> Value {
        assert!(resp.error.is_none(), "{:?}", resp.error);
        resp.result.expect("a successful response carries a result")
    }

    fn ole_ref(value: &Value) -> u64 {
        value["$ole_ref"]
            .as_u64()
            .unwrap_or_else(|| panic!("expected an object reference, got {}", value))
    }

    /// An event's object arguments are handles the client is expected to
    /// CALL, and nothing else registers them: they are minted inside the
    /// session and go straight out on the event stream, never through
    /// `handle_tracked`. So this drives the whole round trip -- subscribe,
    /// change a cell, take the event, call a method on the Range it carried
    /// -- because "the frame contained an $ole_ref" proves nothing about
    /// whether that ref resolves to anything.
    ///
    /// It then pins both ends of the handle's life: `release_event` takes the
    /// whole seq out, and a connection's teardown takes out what the client
    /// never released -- which matters far more than it sounds, because each
    /// of those routing entries holds a `SessionRoute`, and one surviving
    /// route keeps the session thread parked forever with the COM proxies
    /// that keep EXCEL.EXE alive.
    #[test]
    fn test_event_argument_handles_are_routed_then_released_as_a_whole_seq() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
        let mut auth = true;
        let mut owned: Vec<u64> = Vec::new();
        let (events, event_rx) = channel::<SessionEvent>();

        let create = ok_result(rpc(
            &server, &mut auth, &mut owned, &events,
            "create", json!({"class_name": "Excel.Application"}),
        ));
        let root = ole_ref(&create);
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Visible=", "args": [false]})));
        // Without this, `Quit` on a workbook with unsaved changes puts up a
        // save prompt no one can answer and leaves EXCEL.EXE behind.
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "DisplayAlerts=", "args": [false]})));

        assert_eq!(
            ok_result(rpc(&server, &mut auth, &mut owned, &events,
                "subscribe", json!({"handle": root, "event": "SheetChange"}))),
            json!(true),
        );

        let books = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Workbooks"}))));
        let book = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": books, "name": "Add"}))));
        let sheet = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": book, "name": "ActiveSheet"}))));
        let cell = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": sheet, "name": "Range", "args": ["A1"]}))));
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": cell, "name": "Value=", "args": [42]})));

        let frame = server.receive_event(
            event_rx
                .recv_timeout(std::time::Duration::from_secs(30))
                .expect("no event arrived within 30s"),
        );
        assert_eq!(frame.event, "SheetChange");
        let args = frame.args.clone().expect("args default to true");
        assert_eq!(args.len(), 2, "SheetChange(Sh, Target) has two arguments, got {:?}", args);
        let target = ole_ref(&args[1]);

        // The handle the client was handed in its callback must work.
        let address = ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": target, "name": "Address"})));
        assert!(
            address.as_str().unwrap_or_default().contains("$A$1"),
            "an event argument must be callable, not just present; got {:?}",
            address
        );

        assert_eq!(
            ok_result(rpc(&server, &mut auth, &mut owned, &events,
                "release_event", json!({"seq": frame.seq}))),
            json!(true),
        );
        let after = rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": target, "name": "Address"}));
        assert_eq!(
            after.error.expect("a released event argument must not resolve").class,
            "WineOLE::StaleReferenceError",
        );

        // A second event, deliberately NOT released: this is what a client
        // that drops its connection mid-callback leaves behind.
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": cell, "name": "Value=", "args": [43]})));
        let second = server.receive_event(
            event_rx
                .recv_timeout(std::time::Duration::from_secs(30))
                .expect("no second event arrived within 30s"),
        );
        let stranded: Vec<u64> =
            second.args.clone().expect("args default to true").iter().map(ole_ref).collect();
        assert!(!stranded.is_empty());
        {
            let routes = server.routes.lock().unwrap();
            for id in &stranded {
                assert!(
                    routes.contains_key(id),
                    "receive_event must route an event's argument handles: they are minted \
                     inside the session and go straight out on the event stream, so nothing \
                     else ever registers them. (That this happens before the frame goes out is \
                     main.rs's forwarder, and it is structural there rather than asserted \
                     here: receive_event consumes the SessionEvent and the frame is what it \
                     hands back, so there is nothing to send first.)"
                );
            }
        }

        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Quit"})));

        // The connection goes away. It never saw the stranded handles, so
        // nothing in `owned` names them.
        assert!(stranded.iter().all(|id| !owned.contains(id)));
        server.release_all(&owned);
        assert!(
            server.routes.lock().unwrap().is_empty(),
            "teardown must leave no routing entry behind, not even one the connection never \
             saw: each holds a route, and one route keeps the session -- and EXCEL.EXE -- alive"
        );

        // ...and the session must actually FINISH, not merely be told to. A
        // session ended by its last route going away has nobody to
        // acknowledge to, so there is no reply to wait on -- but it drops its
        // event sender as its thread exits, and that is observable from here
        // once this test lets go of its own. Waiting is not politeness: a
        // test process that exits first kills the thread mid-cleanup, and the
        // COM proxies it was releasing keep EXCEL.EXE alive for every later
        // test to trip over. (`handle_connection` waits for the same thing,
        // by joining the forwarder that reads this channel.)
        drop(events);
        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(30);
        let finished = loop {
            match event_rx.recv_timeout(std::time::Duration::from_millis(250)) {
                Err(std::sync::mpsc::RecvTimeoutError::Disconnected) => break true,
                _ if std::time::Instant::now() >= deadline => break false,
                // A late event: keep draining until the channel closes.
                _ => continue,
            }
        };
        assert!(
            finished,
            "the session worker must exit once the connection's teardown has dropped its last \
             route; it is still parked, holding the COM proxies that keep EXCEL.EXE alive"
        );
    }

    /// `args: false` is the branch that decides whether object handles are
    /// minted AT ALL, which is the whole reason a client would ask for it: an
    /// event it only wants to be told about should not leave it holding two
    /// COM references per callback to release.
    ///
    /// So "the frame carried no arguments" is not the claim. An implementation
    /// that minted the handles and then left them out of the JSON would pass
    /// that, and would leak exactly what this option exists to avoid. This
    /// asserts the absence three ways: nothing reported in `new_handles`,
    /// nothing added to the routing table, and -- the one that cannot be
    /// satisfied by a frame that merely omits them -- nothing taken out of the
    /// session's id space, measured by minting a handle either side of the
    /// event and finding the gap is exactly the event's own seq.
    #[test]
    fn test_subscribing_without_arguments_mints_no_handles_at_all() {
        // Held for the whole test -- see dispatch::EXCEL_TEST_LOCK's doc comment.
        let _guard = crate::dispatch::lock_excel_for_test();

        let server = test_server(None);
        let mut auth = true;
        let mut owned: Vec<u64> = Vec::new();
        let (events, event_rx) = channel::<SessionEvent>();

        let create = ok_result(rpc(
            &server, &mut auth, &mut owned, &events,
            "create", json!({"class_name": "Excel.Application"}),
        ));
        let root = ole_ref(&create);
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Visible=", "args": [false]})));
        // Without this, `Quit` on a workbook with unsaved changes puts up a
        // save prompt no one can answer and leaves EXCEL.EXE behind.
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "DisplayAlerts=", "args": [false]})));

        assert_eq!(
            ok_result(rpc(&server, &mut auth, &mut owned, &events,
                "subscribe", json!({"handle": root, "event": "SheetChange", "args": false}))),
            json!(true),
        );

        let books = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Workbooks"}))));
        let book = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": books, "name": "Add"}))));
        let sheet = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": book, "name": "ActiveSheet"}))));
        let cell = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": sheet, "name": "Range", "args": ["A1"]}))));

        // The last id minted before the event. Every id this session hands out
        // -- object handles and event seqs alike -- comes from one counter, so
        // the distance from here to the next one is a count of everything the
        // event consumed.
        let before = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Workbooks"}))));

        let routes_before = server.routes.lock().unwrap().len();

        // The change that raises SheetChange(Sh, Target) -- two object
        // arguments, which is what makes their absence worth asserting.
        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": cell, "name": "Value=", "args": [42]})));

        let ev = event_rx
            .recv_timeout(std::time::Duration::from_secs(30))
            .expect("no event arrived within 30s: args: false must not stop the event itself");
        assert!(
            ev.new_handles.is_empty(),
            "args: false must mint nothing, but the session reported {:?}",
            ev.new_handles
        );
        let frame = server.receive_event(ev);
        assert_eq!(frame.event, "SheetChange");
        assert!(
            frame.args.is_none(),
            "args: false sends null, which is what tells the client the objects were never \
             minted rather than that the event had none; got {:?}",
            frame.args
        );

        // Excel raises one SheetChange for one edit, but the assertion below
        // counts what actually arrived rather than assuming that: each
        // delivered event costs exactly one id (its seq) and nothing else.
        let mut delivered = 1u64;
        while let Ok(extra) =
            event_rx.recv_timeout(std::time::Duration::from_millis(750))
        {
            assert!(extra.new_handles.is_empty(), "args: false must mint nothing");
            assert!(server.receive_event(extra).args.is_none());
            delivered += 1;
        }

        assert_eq!(
            server.routes.lock().unwrap().len(),
            routes_before,
            "an event delivered under args: false must add no routing entries: there is \
             nothing to route, and an entry holds a SessionRoute that keeps the session alive"
        );

        let after = ole_ref(&ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Workbooks"}))));
        assert_eq!(
            after,
            before + delivered + 1,
            "the only id {delivered} event(s) may consume is one seq each: a gap wider than \
             that is arguments minted into the handle table and never handed to anyone -- \
             invisible in the frame, and released only when the session ends"
        );

        ok_result(rpc(&server, &mut auth, &mut owned, &events,
            "invoke", json!({"handle": root, "name": "Quit"})));
        server.release_all(&owned);
    }
}
