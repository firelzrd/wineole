use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::io::{BufRead, Write};

#[derive(Debug, Deserialize)]
pub struct Request {
    pub id: u64,
    pub method: String,
    #[serde(default)]
    pub params: Value,
}

#[derive(Debug, Serialize)]
pub struct Response {
    pub id: u64,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub result: Option<Value>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub error: Option<RpcError>,
}

#[derive(Debug, Serialize)]
pub struct RpcError {
    pub class: String,
    pub message: String,
}

impl Response {
    pub fn ok(id: u64, result: Value) -> Self {
        Response { id, result: Some(result), error: None }
    }

    pub fn err(id: u64, class: &str, message: &str) -> Self {
        Response {
            id,
            result: None,
            error: Some(RpcError { class: class.to_string(), message: message.to_string() }),
        }
    }
}

/// A server-initiated frame. It deliberately has no `id`: the client tells
/// responses from events by whether `id` is present, so adding one here
/// would route an event to a waiter that does not exist.
#[derive(Debug, Serialize)]
pub struct Event {
    pub event: String,
    pub handle: u64,
    pub seq: u64,
    /// `None` when the subscription asked for `args: false`. Serialized as
    /// `null` rather than skipped, so the client can tell "the objects were
    /// never minted" from "this event had zero arguments".
    pub args: Option<Vec<Value>>,
}

pub fn write_event<W: Write>(writer: &mut W, ev: &Event) -> std::io::Result<()> {
    // One write, for the same reason write_response does it: see the comment
    // there. Splitting the newline off costs ~40 ms of delayed-ACK per frame.
    let mut line = serde_json::to_string(ev)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    line.push('\n');
    writer.write_all(line.as_bytes())?;
    writer.flush()
}

pub fn read_request<R: BufRead>(reader: &mut R) -> std::io::Result<Option<Request>> {
    let mut line = String::new();
    let n = reader.read_line(&mut line)?;
    if n == 0 {
        return Ok(None);
    }
    let req: Request = serde_json::from_str(line.trim_end())
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    Ok(Some(req))
}

pub fn write_response<W: Write>(writer: &mut W, resp: &Response) -> std::io::Result<()> {
    // One write, not two. The newline used to be written separately, which
    // put it in its own 1-byte TCP segment behind the body's -- and Nagle
    // holds a small segment while an earlier one is unacknowledged. The
    // client is blocked reading a whole line, so it has no data of its own
    // to piggyback the ACK on, and its ~40 ms delayed-ACK timer had to
    // expire before the newline could go out. Measured at the socket: body
    // in 1.3 ms, newline 42 ms later, on every single RPC.
    let mut line = serde_json::to_string(resp)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    line.push('\n');
    writer.write_all(line.as_bytes())?;
    writer.flush()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    /// Counts `write` calls so the "one write, not two" property can actually
    /// be asserted. A plain `Vec<u8>` records the bytes but not how many
    /// syscalls produced them, so it cannot catch a regression back to the
    /// split write that cost ~40 ms of delayed-ACK on every RPC.
    struct CountingWriter {
        buf: Vec<u8>,
        writes: usize,
    }

    impl CountingWriter {
        fn new() -> Self {
            CountingWriter { buf: Vec::new(), writes: 0 }
        }
    }

    impl Write for CountingWriter {
        fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
            self.writes += 1;
            self.buf.extend_from_slice(buf);
            Ok(buf.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    #[test]
    fn test_read_request_parses_invoke() {
        let input = b"{\"id\":1,\"method\":\"invoke\",\"params\":{\"handle\":3,\"name\":\"Add\"}}\n";
        let mut cursor = Cursor::new(&input[..]);
        let req = read_request(&mut cursor).unwrap().unwrap();
        assert_eq!(req.id, 1);
        assert_eq!(req.method, "invoke");
        assert_eq!(req.params["handle"], 3);
    }

    #[test]
    fn test_read_request_eof_returns_none() {
        let mut cursor = Cursor::new(&b""[..]);
        assert!(read_request(&mut cursor).unwrap().is_none());
    }

    #[test]
    fn test_write_response_ok() {
        let mut buf = Vec::new();
        write_response(&mut buf, &Response::ok(1, serde_json::json!({"$ole_ref": 7}))).unwrap();
        assert_eq!(String::from_utf8(buf).unwrap(), "{\"id\":1,\"result\":{\"$ole_ref\":7}}\n");
    }

    #[test]
    fn test_write_response_err() {
        let mut buf = Vec::new();
        write_response(&mut buf, &Response::err(2, "WIN32OLERuntimeError", "boom")).unwrap();
        assert_eq!(
            String::from_utf8(buf).unwrap(),
            "{\"id\":2,\"error\":{\"class\":\"WIN32OLERuntimeError\",\"message\":\"boom\"}}\n"
        );
    }

    #[test]
    fn test_write_response_emits_body_and_newline_in_one_write() {
        let mut w = CountingWriter::new();
        write_response(&mut w, &Response::ok(1, serde_json::json!({"$ole_ref": 7}))).unwrap();

        assert_eq!(
            String::from_utf8(w.buf.clone()).unwrap(),
            "{\"id\":1,\"result\":{\"$ole_ref\":7}}\n"
        );
        assert_eq!(
            w.writes, 1,
            "the body and its trailing newline must go out as a single write: \
             splitting them leaves a 1-byte segment behind Nagle, and the client \
             (blocked reading a whole line) has nothing to piggyback an ACK on, \
             so its ~40 ms delayed-ACK timer has to fire first"
        );
    }

    #[test]
    fn test_write_event_has_no_id_field() {
        let mut buf = Vec::new();
        let ev = Event {
            event: "SheetChange".to_string(),
            handle: 7,
            seq: 1183,
            args: Some(vec![serde_json::json!({"$ole_ref": 91})]),
        };
        write_event(&mut buf, &ev).unwrap();
        let line = String::from_utf8(buf).unwrap();

        // The client tells a response from an event by the presence of `id`,
        // so an event carrying one would be routed to a waiter that does not
        // exist and the callback would never run.
        assert!(!line.contains("\"id\""), "an event frame must not carry an id: {}", line);
        assert_eq!(
            line,
            "{\"event\":\"SheetChange\",\"handle\":7,\"seq\":1183,\"args\":[{\"$ole_ref\":91}]}\n"
        );
    }

    #[test]
    fn test_write_event_with_no_args_sends_null_not_an_empty_list() {
        let mut buf = Vec::new();
        let ev = Event { event: "Click".to_string(), handle: 3, seq: 9, args: None };
        write_event(&mut buf, &ev).unwrap();
        // `args: false` means "the objects were never minted", which is not
        // the same as "this event had zero arguments". The client shows nil,
        // not [].
        assert_eq!(
            String::from_utf8(buf).unwrap(),
            "{\"event\":\"Click\",\"handle\":3,\"seq\":9,\"args\":null}\n"
        );
    }

    #[test]
    fn test_write_event_emits_body_and_newline_in_one_write() {
        let mut w = CountingWriter::new();
        let ev = Event { event: "Click".to_string(), handle: 3, seq: 9, args: None };
        write_event(&mut w, &ev).unwrap();
        assert_eq!(
            w.writes, 1,
            "same Nagle/delayed-ACK trap as write_response: a lone trailing newline \
             sits behind an unacknowledged segment for ~40 ms"
        );
    }
}
