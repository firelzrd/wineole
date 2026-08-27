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
    let line = serde_json::to_string(resp)
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::InvalidData, e))?;
    writer.write_all(line.as_bytes())?;
    writer.write_all(b"\n")?;
    writer.flush()
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

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
}
