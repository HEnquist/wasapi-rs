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
    let out = repo_root().join("src").join("bindings.rs");

    windows_bindgen::builder()
        .input_default()
        .output(&out)
        .filter_file(&filters)
        .flat()
        .implements([
            "IAudioSessionEvents",
            "IMMNotificationClient",
            "IActivateAudioInterfaceCompletionHandler",
        ])
        .write();
}