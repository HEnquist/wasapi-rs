# AGENTS

## Repo shape

`wasapi-rs` is mostly a thin safe wrapper around the Windows WASAPI interfaces.
That matters when changing or testing it: a lot of behaviour belongs to Windows,
the audio driver, or the device, not to this crate.

The main job of a change here is usually one of these:

- expose a WASAPI capability clearly on the Rust side
- preserve the crate's invariants around cached state and error mapping
- document where the wrapper cannot make a stronger guarantee than WASAPI does

Do not try to "fix" strange Windows or WASAPI behaviour in the wrapper unless the
crate already has a clear, deliberate policy for doing so.
If an API returns odd, incomplete, or plainly bogus values, the default approach is
to expose that behaviour honestly, document it, and test the crate's handling of it
rather than inventing a cleaner story.
The application loopback capture path is the model here: several methods are known
not to work or to report nonsense, and the crate documents those limits instead of
pretending they can be normalised away.

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

## Crate docstrings

The crate docs should explain the Rust wrapper, not rewrite the WASAPI reference.
When a type or method mostly mirrors a Windows interface or constant, link to the
relevant Microsoft documentation and keep the crate doc focused on what is specific
here.

Prefer documenting things like these:

- what the wrapper adds or guarantees
- cached state and invariants the rest of the crate relies on
- apartment or COM lifetime restrictions that matter to Rust callers
- places where the wrapper deliberately differs from, narrows, or cannot fully
  model WASAPI
- crate-specific failure modes, caveats, and examples

If the underlying API is weird, say so plainly.
Prefer "this Windows call returns bogus values in this mode" over adding prose that
implies the wrapper made the behaviour reliable when it did not.

Avoid long restatements of baseline WASAPI behaviour when the Microsoft docs already
cover it well.
Use links to `learn.microsoft.com` for the API reference, and spend the doc comment
budget on the parts a Rust caller would not get from the Windows signature alone.

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