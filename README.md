# Wasapi bindings for Rust

The aim of this crate is to provide easy and safe access to the Wasapi API for audio playback and capture.

Most things map closely to something in the Windows API.

For details on how to use Wasapi, please see [the Windows documentation](https://docs.microsoft.com/en-us/windows/win32/coreaudio/core-audio-interfaces).

Both shared and exclusive modes are supported.

Bindings are generated automatically using the [windows](https://crates.io/crates/windows) crate.

The `loopback` example shows how to simultaneously capture and render sound, with separate threads for capture and render.