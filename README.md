# Wasapi bindings for Rust

This create is meant as a thin wrapper for the Wasapi API.

Most things map directly to something in the Windows API. 

Bindings are generated automatically using the `windows` crate.

The `loopback` example shows how to simultaneously capture and render sound. 