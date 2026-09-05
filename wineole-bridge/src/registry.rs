//! The bridge's instance registry: who is using each auto-created Office
//! instance, and who is allowed to shut it down.
//!
//! Keyed by the cross-apartment identity (see `identity.rs`). Two facts decide
//! shutdown, and they are orthogonal:
//!
//! - `auto_created` is PERMISSION -- may the bridge end this instance at all?
//!   A human's pre-existing Excel (reached via `connect`) is never
//!   auto-created, so the bridge never ends it.
//! - `users` is TIMING -- shutdown happens when the last user leaves.
//!
//! The instance is shut down iff `auto_created && users.is_empty()`. Naming
//! follows consequence, not history: the LAST user runs the cleanup, whichever
//! session created the instance.
//!
//! All locking lives here. Callers get a decision back and act on it; they
//! never hold the registry lock across a COM call.
use std::collections::{HashMap, HashSet};
use std::sync::Mutex;
use serde_json::Value;
use crate::identity::InstanceKey;

#[derive(Clone, Debug)]
pub struct Step {
    pub name: String,
    pub args: Vec<Value>,
}

#[derive(Clone, Debug)]
pub struct CleanupConfig {
    pub steps: Vec<Step>,
    pub callback: bool,
}

struct InstanceRecord {
    auto_created: bool,
    users: HashSet<u64>,
    closing: bool,
    steps: Vec<Step>,
    callback: bool,
}

pub enum JoinResult {
    Joined,
    Closing,
}

pub enum LeaveDecision {
    NotLast,
    RunCleanup {
        steps: Vec<Step>,
        // `steps` is run by the inline cleanup path; `callback` is what the
        // `$cleanup` path reads to decide whether to consult a live client's
        // closure first (emit `$cleanup` and wait) rather than quitting inline.
        callback: bool,
    },
}

pub struct InstanceRegistry {
    inner: Mutex<HashMap<InstanceKey, InstanceRecord>>,
}

impl InstanceRegistry {
    pub fn new() -> Self {
        InstanceRegistry { inner: Mutex::new(HashMap::new()) }
    }

    /// Record a session as a user of `key`.
    ///
    /// `auto_created_if_new` is the permission granted when this call is the
    /// first to register the instance: `true` from `create` and from
    /// `connect_or_create`'s create path, `false` from `connect` and from
    /// `connect_or_create`'s attach path. An EXISTING record keeps its flag --
    /// a later `connect` joining a `create`d instance does not revoke the
    /// permission, and a `create` cannot collide because a fresh process has a
    /// fresh OXID.
    ///
    /// A record that is already `closing` refuses new users: the caller
    /// (`connect`) turns that into `InstanceClosingError`, and
    /// `connect_or_create` never asks -- it detects the closing record before
    /// attaching and falls through to create.
    pub fn join(&self, key: InstanceKey, session: u64, auto_created_if_new: bool, cleanup: &CleanupConfig) -> JoinResult {
        let mut m = self.inner.lock().unwrap();
        match m.get_mut(&key) {
            Some(rec) if rec.closing => JoinResult::Closing,
            Some(rec) => {
                rec.users.insert(session);
                JoinResult::Joined
            }
            None => {
                let mut users = HashSet::new();
                users.insert(session);
                m.insert(
                    key,
                    InstanceRecord {
                        auto_created: auto_created_if_new,
                        users,
                        closing: false,
                        steps: cleanup.steps.clone(),
                        callback: cleanup.callback,
                    },
                );
                JoinResult::Joined
            }
        }
    }

    /// A session releasing its root handle leaves the instance.
    ///
    /// Returns `RunCleanup` (and marks the record `closing`) when this was the
    /// last user of an auto-created instance. Returns `NotLast` otherwise --
    /// and, when a non-auto-created record has just lost its last user, removes
    /// that record as bookkeeping (nobody is using it and nobody may quit it).
    pub fn on_root_release(&self, key: InstanceKey, session: u64) -> LeaveDecision {
        let mut m = self.inner.lock().unwrap();
        let Some(rec) = m.get_mut(&key) else { return LeaveDecision::NotLast };
        rec.users.remove(&session);
        if rec.auto_created && rec.users.is_empty() && !rec.closing {
            rec.closing = true;
            LeaveDecision::RunCleanup { steps: rec.steps.clone(), callback: rec.callback }
        } else {
            if rec.users.is_empty() && !rec.closing {
                m.remove(&key);
            }
            LeaveDecision::NotLast
        }
    }

    /// After a client closure ran (the `$cleanup` path): should the bridge now
    /// run the steps? Yes when the record is still `closing` -- a closure that
    /// called `leave_open` clears it, and then the steps must NOT run. Either
    /// way this session is done with the record, so it is removed here.
    pub fn confirm_cleanup(&self, key: InstanceKey) -> bool {
        let mut m = self.inner.lock().unwrap();
        match m.get(&key) {
            Some(rec) if rec.closing => {
                m.remove(&key);
                true
            }
            _ => {
                // leave_open cleared `closing`. Drop the record only if no new
                // user joined in the window (leave_open re-opened the record to
                // joiners). If someone joined, leave the record for them.
                if m.get(&key).map(|r| r.users.is_empty()).unwrap_or(false) {
                    m.remove(&key);
                }
                false
            }
        }
    }

    /// Remove the record after inline steps ran (the no-callback / dead-client
    /// path).
    pub fn finish_cleanup(&self, key: InstanceKey) {
        self.inner.lock().unwrap().remove(&key);
    }

    /// Revoke shutdown permission: this instance will not be cleaned up. Clears
    /// `closing` too, so a closure calling this from inside the `$cleanup` path
    /// stops the pending steps. NOP if the key is unknown (a non-participating
    /// session, or one whose key was never registered).
    pub fn leave_open(&self, key: InstanceKey) {
        if let Some(rec) = self.inner.lock().unwrap().get_mut(&key) {
            rec.auto_created = false;
            rec.closing = false;
        }
    }

    #[cfg(test)]
    pub fn snapshot(&self, key: InstanceKey) -> Option<(bool, usize, bool)> {
        self.inner.lock().unwrap().get(&key).map(|r| (r.auto_created, r.users.len(), r.closing))
    }

    /// `(auto_created, users, closing)` for every live record, order
    /// unspecified. Session-integration tests use this to prove co-use
    /// STRUCTURALLY -- two sessions on one `(OXID, OID)` collapse to a single
    /// record with two users -- without needing the key, which lives on the
    /// session threads.
    #[cfg(test)]
    pub fn record_summary(&self) -> Vec<(bool, usize, bool)> {
        self.inner.lock().unwrap().values().map(|r| (r.auto_created, r.users.len(), r.closing)).collect()
    }
}

/// Validate and read the optional `cleanup` object from a create/connect
/// request's params.
///
/// `Ok(None)` when there is no `cleanup` key -- the session does not
/// participate in the registry, and costs exactly what it costs today. A
/// malformed `cleanup` is a `ProtocolError` rather than a silently ignored
/// one: a client that asked for cleanup and typed it wrong must hear so, not
/// leak an Excel because the bridge quietly dropped the request.
pub fn parse_cleanup(params: &Value) -> Result<Option<CleanupConfig>, (String, String)> {
    let Some(cleanup) = params.get("cleanup") else { return Ok(None) };
    let err = |m: &str| ("WineOLE::ProtocolError".to_string(), m.to_string());
    let obj = cleanup.as_object().ok_or_else(|| err("cleanup must be an object"))?;
    let steps_val = obj.get("steps").ok_or_else(|| err("cleanup.steps is required"))?;
    let steps_arr = steps_val.as_array().ok_or_else(|| err("cleanup.steps must be an array"))?;
    let mut steps = Vec::with_capacity(steps_arr.len());
    for step in steps_arr {
        let s = step.as_object().ok_or_else(|| err("each cleanup step must be an object"))?;
        let name = s
            .get("name")
            .and_then(|v| v.as_str())
            .ok_or_else(|| err("each cleanup step needs a string name"))?
            .to_string();
        let args = match s.get("args") {
            None => Vec::new(),
            Some(a) => {
                let arr = a.as_array().ok_or_else(|| err("cleanup step args must be an array"))?;
                for arg in arr {
                    if arg.is_object() || arg.is_array() {
                        return Err(err(
                            "cleanup step args must be scalars (bool/number/string/null); \
                             object references cannot be used in cleanup steps",
                        ));
                    }
                }
                arr.clone()
            }
        };
        steps.push(Step { name, args });
    }
    let callback = obj.get("callback").and_then(|v| v.as_bool()).unwrap_or(false);
    Ok(Some(CleanupConfig { steps, callback }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::identity::synthetic_key;

    fn cfg(callback: bool) -> CleanupConfig {
        CleanupConfig {
            steps: vec![Step { name: "Quit".into(), args: vec![] }],
            callback,
        }
    }

    #[test]
    fn create_registers_auto_created_then_release_removes_it() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(10);
        assert!(matches!(reg.join(k, 1, true, &cfg(false)), JoinResult::Joined));
        assert_eq!(reg.snapshot(k), Some((true, 1, false)));
        match reg.on_root_release(k, 1) {
            LeaveDecision::RunCleanup { .. } => {}
            _ => panic!("last user of an auto-created instance must run cleanup"),
        }
        // Still present and closing until finish_cleanup.
        assert_eq!(reg.snapshot(k), Some((true, 0, true)));
        reg.finish_cleanup(k);
        assert_eq!(reg.snapshot(k), None);
    }

    #[test]
    fn co_users_only_the_last_runs_cleanup() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(11);
        assert!(matches!(reg.join(k, 1, true, &cfg(false)), JoinResult::Joined)); // creator
        assert!(matches!(reg.join(k, 2, false, &cfg(false)), JoinResult::Joined)); // co-user via connect
        assert_eq!(reg.snapshot(k), Some((true, 2, false)));
        // Creator leaves first: not last, record stays, flag preserved.
        assert!(matches!(reg.on_root_release(k, 1), LeaveDecision::NotLast));
        assert_eq!(reg.snapshot(k), Some((true, 1, false)));
        // Connector leaves last: runs cleanup.
        assert!(matches!(reg.on_root_release(k, 2), LeaveDecision::RunCleanup { .. }));
    }

    #[test]
    fn connect_does_not_overwrite_auto_created() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(12);
        reg.join(k, 1, true, &cfg(false)); // create -> auto_created true
        reg.join(k, 2, false, &cfg(false)); // connect must NOT flip it to false
        assert_eq!(reg.snapshot(k), Some((true, 2, false)));
    }

    #[test]
    fn leave_open_revokes_permission_so_release_runs_no_cleanup() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(13);
        reg.join(k, 1, true, &cfg(false));
        reg.leave_open(k);
        assert_eq!(reg.snapshot(k), Some((false, 1, false)));
        // Last user leaves a non-auto-created record: NotLast, and the record
        // is removed as bookkeeping.
        assert!(matches!(reg.on_root_release(k, 1), LeaveDecision::NotLast));
        assert_eq!(reg.snapshot(k), None);
    }

    #[test]
    fn leave_open_is_nop_for_unknown_or_already_false() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(14);
        reg.leave_open(k); // unknown: no panic, no record created
        assert_eq!(reg.snapshot(k), None);
        reg.join(k, 1, false, &cfg(false)); // connect -> already false
        reg.leave_open(k);
        assert_eq!(reg.snapshot(k), Some((false, 1, false)));
    }

    #[test]
    fn a_closing_record_refuses_new_users() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(15);
        reg.join(k, 1, true, &cfg(false));
        assert!(matches!(reg.on_root_release(k, 1), LeaveDecision::RunCleanup { .. })); // marks closing
        // A connect landing here must be refused.
        assert!(matches!(reg.join(k, 2, false, &cfg(false)), JoinResult::Closing));
    }

    #[test]
    fn confirm_cleanup_runs_steps_when_still_closing() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(16);
        reg.join(k, 1, true, &cfg(true));
        reg.on_root_release(k, 1); // closing = true
        assert!(reg.confirm_cleanup(k), "still closing -> steps run");
        assert_eq!(reg.snapshot(k), None);
    }

    #[test]
    fn confirm_cleanup_skips_steps_when_closure_called_leave_open() {
        let reg = InstanceRegistry::new();
        let k = synthetic_key(17);
        reg.join(k, 1, true, &cfg(true));
        reg.on_root_release(k, 1); // closing = true
        reg.leave_open(k); // closure revoked permission
        assert!(!reg.confirm_cleanup(k), "leave_open cleared closing -> no steps");
        assert_eq!(reg.snapshot(k), None); // users empty, record cleaned up
    }

    #[test]
    fn parse_cleanup_reads_steps_and_callback() {
        let v = serde_json::json!({
            "cleanup": {
                "steps": [{"name": "DisplayAlerts=", "args": [false]}, {"name": "Quit", "args": []}],
                "callback": true
            }
        });
        let cfg = super::parse_cleanup(&v).unwrap().unwrap();
        assert_eq!(cfg.steps.len(), 2);
        assert_eq!(cfg.steps[0].name, "DisplayAlerts=");
        assert_eq!(cfg.steps[0].args, vec![serde_json::json!(false)]);
        assert!(cfg.callback);
    }

    #[test]
    fn parse_cleanup_absent_is_none_and_defaults_callback_false() {
        assert!(super::parse_cleanup(&serde_json::json!({})).unwrap().is_none());
        let v = serde_json::json!({"cleanup": {"steps": []}});
        let cfg = super::parse_cleanup(&v).unwrap().unwrap();
        assert!(cfg.steps.is_empty());
        assert!(!cfg.callback);
    }

    #[test]
    fn parse_cleanup_rejects_object_or_array_args_and_bad_shapes() {
        // An $ole_ref would arrive as an object arg -- rejected.
        let ole_ref = serde_json::json!({"cleanup": {"steps": [{"name": "X", "args": [{"$ole_ref": 1}]}]}});
        assert!(super::parse_cleanup(&ole_ref).is_err());
        let bad_name = serde_json::json!({"cleanup": {"steps": [{"name": 5, "args": []}]}});
        assert!(super::parse_cleanup(&bad_name).is_err());
        let bad_args = serde_json::json!({"cleanup": {"steps": [{"name": "X", "args": 5}]}});
        assert!(super::parse_cleanup(&bad_args).is_err());
        let bad_steps = serde_json::json!({"cleanup": {"steps": "nope"}});
        assert!(super::parse_cleanup(&bad_steps).is_err());
    }
}
