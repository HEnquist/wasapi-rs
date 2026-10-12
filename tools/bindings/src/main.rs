use std::path::{Path, PathBuf};

// Read at run time rather than with env!, which bakes in the path of whichever copy
// of the repo the binary was built in, and a stale build then writes into that copy.
fn repo_root() -> PathBuf {
    let manifest_dir =
        std::env::var_os("CARGO_MANIFEST_DIR").expect("run the generator with cargo run");
    Path::new(&manifest_dir)
        .ancestors()
        .nth(2)
        .expect("generator lives under tools/bindings")
        .to_path_buf()
}

fn main() {
    let tools = repo_root().join("tools").join("bindings");

    windows_bindgen::builder()
        .input_default()
        .output(repo_root().join("src").join("bindings").join("generated.rs"))
        .filter_file(tools.join("filters.txt"))
        .flat()
        // Only the callback interfaces that the crate implements get `_Impl` traits.
        // The names must be fully qualified, a bare name silently matches nothing.
        // The default metadata puts all of Win32 in the single `Windows.Win32` namespace.
        .implements([
            "Windows.Win32.IAudioSessionEvents",
            "Windows.Win32.IMMNotificationClient",
            "Windows.Win32.IActivateAudioInterfaceCompletionHandler",
        ])
        .write();

    // The device tests use their own bindings, the way a caller of the raw accessors would.
    windows_bindgen::builder()
        .input_default()
        .output(repo_root().join("tests").join("device").join("bindings.rs"))
        .filter_file(tools.join("filters_device_tests.txt"))
        .flat()
        .implements(["Windows.Win32.IActivateAudioInterfaceCompletionHandler"])
        .write();
}
