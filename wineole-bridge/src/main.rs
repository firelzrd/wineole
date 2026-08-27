mod protocol;
mod value;
mod dispatch;
mod session;
mod server;

use std::env;
use std::io::BufReader;
use std::net::{IpAddr, TcpListener, TcpStream};
use std::sync::atomic::{AtomicUsize, Ordering};
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

    let server = Arc::new(Server::new(token));
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

fn handle_connection(stream: TcpStream, server: Arc<Server>) {
    let peer_is_loopback = stream.peer_addr().map(|a| a.ip().is_loopback()).unwrap_or(false);
    let mut authenticated = server.token.is_none() || peer_is_loopback;
    let mut reader = BufReader::new(match stream.try_clone() {
        Ok(cloned) => cloned,
        Err(e) => {
            eprintln!("Failed to clone stream for peer: {}", e);
            return;
        }
    });
    let mut writer = stream;

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
                let resp =
                    server.handle_tracked(req, peer_is_loopback, &mut authenticated, &mut owned_handles);
                if protocol::write_response(&mut writer, &resp).is_err() {
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
}
