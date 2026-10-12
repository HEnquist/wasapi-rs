use std::path::{Path, PathBuf};

fn repo_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .ancestors()
        .nth(2)
        .expect("generator lives under tools/bindings")
        .to_path_buf()
}

fn main() {
    let filters = repo_root().join("tools").join("bindings").join("filters.txt");
    let out = repo_root().join("src").join("bindings").join("generated.rs");

    windows_bindgen::builder()
        .input_default()
        .output(&out)
        .filter_file(&filters)
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
}