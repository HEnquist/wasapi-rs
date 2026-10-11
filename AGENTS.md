# AGENTS

## Repo shape

`wasapi-rs` is mostly a thin safe wrapper around the Windows WASAPI interfaces.
That matters when changing or testing it: a lot of behaviour belongs to Windows,
the audio driver, or the device, not to this crate.

The main job of a change here is usually one of these:

- expose a WASAPI capability clearly on the Rust side
- preserve the crate's invariants around cached state and error mapping
- document where the wrapper cannot make a stronger guarantee than WASAPI does

## Testing strategy

Keep tests simple and local to the crate behaviour they are proving.
This crate is very easy to over-test in the wrong direction.
A complicated end to end test can stop telling us whether a failure came from the
crate, from WASAPI, from the test program, or from the audio device.

Do not build a music player, recorder, analyser, or signal processing pipeline
inside a test unless the point of the test is that code itself.
If a test starts to need elaborate buffering, generated audio content, file IO,
or analysis of captured samples, stop and ask whether it is still testing the
wrapper rather than the surrounding scaffolding.

Prefer assertions on the crate's own behaviour:

- return values and error mapping
- direction, share mode, timing mode, and other cached state
- public API behaviour visible to downstream crates
- format and period arithmetic done by this crate
- simple invariants such as frame counts, buffer sizes, and flag round trips

Prefer the smallest test surface that can prove the point:

- unit tests in the owning module for pure helper logic
- `tests/public_api.rs` for public API behaviour that needs no audio device
- `tests/device/` only when the behaviour cannot be exercised without a real endpoint

## Device-backed tests

The device-backed suite is intentionally narrow.
It exists to prove that the wrapper can open streams and move buffers through the
API, not to prove audio quality or implement a realistic application.

When working in `tests/device/`, keep these rules:

- Keep helpers small.
- Render simple data, usually zeros.
- Capture data only far enough to prove the stream ran.
- Assert on what the library reports, especially frame counts and expected errors.
- Do not add signal generation, mixing, decoding, or sample analysis unless there
  is no smaller way to prove crate behaviour.
- Do not make one test responsible for multiple behaviours when separate tests can
  isolate them.

The existing fixture in `tests/device/support.rs` is the model.
It serialises access to the virtual cable, initialises COM once per test thread,
and deliberately avoids turning the suite into an audio application.

## Running tests

Use the normal suite first:

```sh
cargo test
```

That covers crate behaviour which does not need a live audio device.

The hardware-backed suite is behind a feature and a dedicated test target:

```sh
cargo test --features device-tests --test device -- --test-threads=1
```

Notes:

- `--test-threads=1` is required because exclusive mode owns the endpoint.
- The suite looks for VB-Cable endpoints.
- Without the cable, tests print a skip message and return.
- Set `WASAPI_REQUIRE_DEVICE=1` to turn missing-device skips into failures.
- Set `WASAPI_TEST_LOG` to `error`, `warn`, `info`, `debug`, or `trace` when
  extra library logging is useful.

## Change guidance

When adding or editing tests:

- start from the narrowest existing test file that already owns the behaviour
- prefer proving one wrapper contract per test
- keep hardware assumptions explicit
- avoid broad refactors of the test harness unless the harness itself is the bug

When a behaviour is mostly Windows-owned, document the limitation and test only
the part this crate is responsible for.