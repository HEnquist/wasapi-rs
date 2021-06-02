# Wasapi bindings for Rust

The aim of this crate is to provide easy and safe access to the Wasapi API for audio playback and capture. Both shared and exclusive modes are supported.

The Windows types and methods are thinly wrapped into Rust structs, so the presented API is quite similar to the original Windows API.

For details on how to use Wasapi, please see [the Windows documentation](https://docs.microsoft.com/en-us/windows/win32/coreaudio/core-audio-interfaces).

Bindings are generated automatically using the [windows](https://crates.io/crates/windows) crate.

## Examples

- The `playsine` example plays a sine wave in shared mode on the default output device.

- The `loopback` example shows how to simultaneously capture and render sound, with separate threads for capture and render.