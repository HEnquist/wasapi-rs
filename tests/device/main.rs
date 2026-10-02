// Integration tests that need a real audio device.
//
// Gated behind the `device-tests` feature so a plain `cargo test` never builds or
// runs them, see the [[test]] entry in Cargo.toml. Run them with:
//
//     cargo test --features device-tests --test device -- --test-threads=1
//
// One test binary rather than several, because the cable has a single render and a
// single capture endpoint and an exclusive mode client owns the endpoint. Keeping
// everything in one process lets a mutex in `support` serialise access, which it
// could not do across separate test binaries. See tests/device/support.rs.
//
// If anyone ever wants to run these under `cargo nextest`, which runs separate
// binaries in parallel, that needs a test group with `max-threads = 1`.

// Not every module uses every helper, and `clippy --all-targets -- -D warnings`
// turns an unused one into an error.
#[allow(dead_code)]
mod support;

mod capability;
mod enumeration;
mod errors;
mod streams;
