mod pump;
mod protocol;
mod value;
mod dispatch;
mod identity;
mod registry;
mod session;
mod server;
mod sink;

use std::env;
use std::io::{BufReader, Write};
use std::net::{IpAddr, TcpListener, TcpStream};
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::mpsc;
use std::sync::Arc;
use std::time::Duration;
use server::Server;

const DEFAULT_PORT: u16 = 47800;
const DEFAULT_IDLE_TIMEOUT_SECS: u64 = 1800;

fn main() {
    let port = match env::args().nth(1) {
        None => DEFAULT_PORT,
        Some(arg) => match arg.parse::<u16>() {
            Ok(p) => p,
            Err(e) => {
                eprintln!(
                    "WARNING: could not parse port argument {:?} ({}), falling back to {}",
                    arg, e, DEFAULT_PORT
                );
                DEFAULT_PORT
            }
        },
    };
    let bind_addr = env::var("WINEOLE_BIND").unwrap_or_else(|_| "127.0.0.1".to_string());
    let token = env::var("WINEOLE_TOKEN").ok();
    let idle_timeout_secs = match env::var("WINEOLE_IDLE_TIMEOUT_SECS") {
        Err(_) => DEFAULT_IDLE_TIMEOUT_SECS,
        Ok(raw) => match raw.parse::<u64>() {
            Ok(secs) => secs,
            Err(e) => {
                eprintln!(
                    "WARNING: could not parse WINEOLE_IDLE_TIMEOUT_SECS={:?} ({}), falling back to {}s",
                    raw, e, DEFAULT_IDLE_TIMEOUT_SECS
                );
                DEFAULT_IDLE_TIMEOUT_SECS
            }
        },
    };

    if !is_loopback_bind(&bind_addr) && token.is_none() {
        eprintln!(
            "WARNING: binding to {} with no WINEOLE_TOKEN configured — this exposes COM \
             automation to any host that can reach this port. Set WINEOLE_TOKEN to require \
             authentication.",
            bind_addr
        );
    }

    let listener = TcpListener::bind((bind_addr.as_str(), port))
        .unwrap_or_else(|e| panic!("failed to bind {}:{}: {}", bind_addr, port, e));
    println!("wineole-bridge listening on {}:{}", bind_addr, port);

    let registry = Arc::new(crate::registry::InstanceRegistry::new());
    let server = Arc::new(Server::new(token, registry));
    let active_connections = Arc::new(AtomicUsize::new(0));

    {
        let active_connections = active_connections.clone();
        std::thread::spawn(move || {
            let mut idle_elapsed_secs: u64 = 0;
            loop {
                std::thread::sleep(Duration::from_secs(30));
                if active_connections.load(Ordering::SeqCst) == 0 {
                    idle_elapsed_secs += 30;
                    if idle_elapsed_secs >= idle_timeout_secs {
                        println!("wineole-bridge: idle timeout reached, exiting");
                        std::process::exit(0);
                    }
                } else {
                    idle_elapsed_secs = 0;
                }
            }
        });
    }

    for stream in listener.incoming() {
        let stream = match stream {
            Ok(s) => s,
            Err(_) => continue,
        };
        let server = server.clone();
        let active_connections = active_connections.clone();
        active_connections.fetch_add(1, Ordering::SeqCst);
        std::thread::spawn(move || {
            handle_connection(stream, server);
            active_connections.fetch_sub(1, Ordering::SeqCst);
        });
    }
}

/// Is `addr` a loopback bind? A literal `localhost` counts; anything that is
/// not a parseable loopback IP (including `0.0.0.0`, `::`, and any concrete
/// external address) does not. Deliberately conservative: an address we cannot
/// classify is treated as non-loopback, so the warning errs towards being shown.
fn is_loopback_bind(addr: &str) -> bool {
    if addr.eq_ignore_ascii_case("localhost") {
        return true;
    }
    addr.trim_matches(|c| c == '[' || c == ']')
        .parse::<IpAddr>()
        .map(|ip| ip.is_loopback())
        .unwrap_or(false)
}

/// Everything that can go out on a connection's socket.
///
/// The socket has exactly one writer thread and this is what it writes, for a
/// reason the protocol makes unforgiving: the client reads whole lines. A
/// response written from the request loop while an event is being written
/// from a session's thread produces a torn line -- two half-frames the client
/// cannot parse and cannot resynchronise from.
enum OutFrame {
    Response(protocol::Response),
    Event(protocol::Event),
}

/// The whole of the one writer thread's job: take frames off `frames` and put
/// them on `writer`, one at a time, until every sender is gone.
///
/// Sequential BY CONSTRUCTION -- that is the point of the function existing.
/// Producers are concurrent (a request loop and an event forwarder, on
/// different threads), and this loop is where that concurrency stops: whatever
/// order the frames arrive in, each one is written whole before the next one
/// starts, so the client's `read_line` never sees two halves of different
/// frames spliced together.
///
/// Generic over `Write` rather than taking the `TcpStream` it is always given,
/// so a test can drive it through a socket it also holds the other end of.
///
/// Returning is part of the contract, not an implementation detail: the
/// connection joins this thread on the way out, and it can only be joined once
/// the last `Sender` has gone. See the shutdown at the foot of
/// [`handle_connection`].
fn write_frames<W: Write>(writer: &mut W, frames: mpsc::Receiver<OutFrame>) {
    for frame in frames {
        let ok = match frame {
            OutFrame::Response(r) => protocol::write_response(writer, &r).is_ok(),
            OutFrame::Event(e) => protocol::write_event(writer, &e).is_ok(),
        };
        // The socket is gone; draining the rest would only fail once per
        // frame.
        if !ok {
            break;
        }
    }
}

fn handle_connection(stream: TcpStream, server: Arc<Server>) {
    // Belt and braces for the Nagle/delayed-ACK stall fixed in
    // protocol.rs::write_response. That fix makes each response a single
    // write, which is enough today because every response is far below the
    // MSS. A response that did exceed it would still be split across
    // segments, and the trailing one would fall into the same trap -- this
    // keeps that from ever costing 40 ms again. Failure is not fatal: the
    // connection still works, just without the guarantee.
    let _ = stream.set_nodelay(true);

    let peer_is_loopback = stream.peer_addr().map(|a| a.ip().is_loopback()).unwrap_or(false);
    let mut authenticated = server.token.is_none() || peer_is_loopback;
    let mut reader = BufReader::new(match stream.try_clone() {
        Ok(cloned) => cloned,
        Err(e) => {
            eprintln!("Failed to clone stream for peer: {}", e);
            return;
        }
    });
    // One thread owns writing. See `OutFrame`.
    let (out_tx, out_rx) = mpsc::channel::<OutFrame>();
    let writer_handle = std::thread::spawn(move || {
        let mut writer = stream;
        write_frames(&mut writer, out_rx);
    });

    // Every session this connection creates pushes its events here. One
    // stream for the whole connection rather than one per session: the
    // frames already carry the handle they came from, and a thread per
    // session would only add a merge nobody asked for.
    let (event_tx, event_rx) = mpsc::channel::<session::SessionEvent>();
    let forwarder = {
        let server = server.clone();
        let out_tx = out_tx.clone();
        std::thread::spawn(move || {
            for ev in event_rx {
                // `receive_event` registers the handles this event minted for
                // its arguments before the frame goes out, so a client that
                // calls a method on one in its callback finds it routed.
                //
                // These two lines cannot be written the other way round, and
                // that is by design rather than by discipline: `receive_event`
                // CONSUMES the `SessionEvent`, and the frame exists only as
                // what it hands back, so there is nothing to send until the
                // registration has happened.
                let frame = server.receive_event(ev);
                if out_tx.send(OutFrame::Event(frame)).is_err() {
                    break;
                }
            }
        })
    };

    // Every handle this connection was ever handed — the root handle from
    // create/connect plus every object handle minted by an invoke. Nothing
    // else tracks them: a client that disconnects without releasing (the
    // normal case, since intermediate objects are never explicitly released)
    // would otherwise pin its session's worker thread forever, keeping the
    // automated application (e.g. EXCEL.EXE) alive with it.
    let mut owned_handles: Vec<u64> = Vec::new();

    loop {
        match protocol::read_request(&mut reader) {
            Ok(Some(req)) => {
                let resp = server.handle_tracked(
                    req,
                    peer_is_loopback,
                    &mut authenticated,
                    &mut owned_handles,
                    &event_tx,
                );
                if out_tx.send(OutFrame::Response(resp)).is_err() {
                    break;
                }
            }
            Ok(None) => break,
            Err(_) => break,
        }
    }

    // Runs for every exit path above — clean disconnect, malformed request,
    // and I/O error alike — because it sits after the loop rather than in any
    // one branch.
    server.release_all(&owned_handles);

    // Shut down in the order the frames flow, so nothing outlives what feeds
    // it. `release_all` above ended every session this connection created,
    // and a session drops its sender as its thread exits; dropping ours is
    // what lets `event_rx` finally disconnect.
    drop(event_tx);
    let _ = forwarder.join();
    drop(out_tx);
    let _ = writer_handle.join();
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{json, Value};
    use std::io::BufRead;
    use std::thread::JoinHandle;

    /// Two ends of one loopback connection: the same thing `handle_connection`
    /// is given, with the client's end kept here to read what came out.
    fn socket_pair() -> (TcpStream, TcpStream) {
        let listener = TcpListener::bind("127.0.0.1:0").expect("bind a loopback listener");
        let addr = listener.local_addr().expect("local_addr");
        let client = TcpStream::connect(addr).expect("connect to the listener");
        let (server, _) = listener.accept().expect("accept the connection");
        (client, server)
    }

    /// `join`, bounded. Every failure these tests are about is a thread that
    /// never finishes, and an unbounded `join` on one hangs the suite instead
    /// of failing it.
    fn join_within<T: Send + 'static>(handle: JoinHandle<T>, what: &str, secs: u64) -> T {
        let (tx, rx) = mpsc::channel();
        std::thread::spawn(move || {
            let _ = tx.send(handle.join());
        });
        match rx.recv_timeout(Duration::from_secs(secs)) {
            Ok(Ok(value)) => value,
            Ok(Err(_)) => panic!("{what} panicked"),
            Err(_) => panic!("{what} did not finish within {secs}s"),
        }
    }

    /// Read whole lines off `stream` until it closes.
    fn read_lines(stream: TcpStream) -> JoinHandle<Vec<String>> {
        std::thread::spawn(move || {
            BufReader::new(stream).lines().map_while(Result::ok).collect()
        })
    }

    /// The property the single writer thread exists for: a client reads whole
    /// LINES, and two threads writing to one socket produce torn ones -- two
    /// half-frames it can neither parse nor resynchronise from.
    ///
    /// So this pushes responses and events from two threads at once, which is
    /// exactly what a connection does (its request loop answers while its
    /// event forwarder forwards), and asserts what the client would see: every
    /// line is one whole JSON frame, and every frame sent arrives exactly
    /// once. The frames are padded well past a TCP segment so that a writer
    /// that did not own the socket would interleave visibly rather than by
    /// luck.
    ///
    /// It also pins that `write_frames` RETURNS when its senders are gone,
    /// which is what makes `writer_handle.join()` at the foot of
    /// `handle_connection` terminate rather than deadlock.
    #[test]
    fn test_one_writer_turns_two_concurrent_producers_into_whole_lines() {
        const EACH: usize = 200;
        // Comfortably more than one segment, so an unsynchronised writer's
        // frames would be spliced in the middle rather than at a boundary.
        let padding = "x".repeat(2000);

        let (client, server) = socket_pair();
        let reader = read_lines(client);

        let (out_tx, out_rx) = mpsc::channel::<OutFrame>();
        let writer = std::thread::spawn(move || {
            let mut server = server;
            write_frames(&mut server, out_rx);
        });

        let responses = {
            let tx = out_tx.clone();
            let padding = padding.clone();
            std::thread::spawn(move || {
                for i in 0..EACH {
                    let body = json!({"$ole_ref": i, "padding": padding});
                    tx.send(OutFrame::Response(protocol::Response::ok(i as u64, body)))
                        .expect("the writer must still be running");
                }
            })
        };
        let events = {
            let tx = out_tx.clone();
            std::thread::spawn(move || {
                for i in 0..EACH {
                    let frame = protocol::Event {
                        event: "SheetChange".to_string(),
                        handle: 7,
                        seq: i as u64,
                        args: Some(vec![json!({"$ole_ref": i, "padding": padding})]),
                    };
                    tx.send(OutFrame::Event(frame)).expect("the writer must still be running");
                }
            })
        };

        join_within(responses, "the response producer", 30);
        join_within(events, "the event producer", 30);
        drop(out_tx);
        join_within(writer, "write_frames", 30);
        let lines = join_within(reader, "the reader", 30);

        assert_eq!(
            lines.len(),
            EACH * 2,
            "every frame sent must arrive as exactly one line"
        );

        let mut response_ids: Vec<u64> = Vec::new();
        let mut event_seqs: Vec<u64> = Vec::new();
        for line in &lines {
            let frame: Value = serde_json::from_str(line).unwrap_or_else(|e| {
                panic!(
                    "every line the client reads must be one whole JSON frame, or it can \
                     neither parse it nor resynchronise: {e} in {:?}",
                    &line[..line.len().min(120)]
                )
            });
            match frame.get("id").and_then(Value::as_u64) {
                Some(id) => {
                    assert!(
                        frame.get("event").is_none(),
                        "a response must not be half of an event: {frame}"
                    );
                    response_ids.push(id);
                }
                None => {
                    assert_eq!(
                        frame["event"], json!("SheetChange"),
                        "a frame with no id is an event, whole: {frame}"
                    );
                    event_seqs.push(frame["seq"].as_u64().expect("an event frame carries a seq"));
                }
            }
        }

        response_ids.sort_unstable();
        event_seqs.sort_unstable();
        let expected: Vec<u64> = (0..EACH as u64).collect();
        assert_eq!(response_ids, expected, "every response must arrive exactly once");
        assert_eq!(event_seqs, expected, "every event must arrive exactly once");
    }

    /// The connection's four-step shutdown, which nothing used to reach at
    /// all: `drop(event_tx)` -> join the forwarder -> `drop(out_tx)` -> join
    /// the writer. Each step ends what the next one waits for, so joining a
    /// thread before dropping the sender that feeds it does not fail, it
    /// HANGS -- which is why the assertion here is that `handle_connection`
    /// returns.
    ///
    /// `no_such_method` keeps it away from COM: the point is the plumbing --
    /// requests in, frames out through the one writer, and a clean teardown
    /// when the client goes away -- not what any particular method does.
    #[test]
    fn test_a_connection_answers_through_its_writer_and_shuts_down_when_the_client_goes() {
        let (client, server_end) = socket_pair();
        let server = Arc::new(Server::new(None, Arc::new(crate::registry::InstanceRegistry::new())));
        let connection = std::thread::spawn(move || handle_connection(server_end, server));

        let mut writer = client.try_clone().expect("clone the client end");
        let mut reader = BufReader::new(client);
        for id in 1..=5u64 {
            let line = json!({"id": id, "method": "no_such_method", "params": {}}).to_string();
            writeln!(writer, "{line}").expect("the connection must accept a request");
        }
        writer.flush().expect("flush");

        for id in 1..=5u64 {
            let mut line = String::new();
            reader.read_line(&mut line).expect("a response must come back");
            let frame: Value = serde_json::from_str(line.trim_end())
                .unwrap_or_else(|e| panic!("a response must be one whole line: {e} in {line:?}"));
            assert_eq!(frame["id"], json!(id), "responses come back in order, one per request");
            assert_eq!(frame["error"]["class"], json!("WineOLE::ProtocolError"));
        }

        // The client goes away, which is the ordinary end of a connection.
        drop(reader);
        drop(writer);
        join_within(
            connection,
            "handle_connection: it must shut its threads down in the order the frames flow \
             (drop the event sender, join the forwarder, drop the frame sender, join the \
             writer); joining before dropping waits for a sender that is still held, forever",
            60,
        );
    }
}
