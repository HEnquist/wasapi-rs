// Record audio from a single application, and save the raw samples
// to the file 'recorded.raw'.
//
// This example captures from Firefox by default.
// Give the executable name of another application as an argument to capture from that one:
//
//     cargo run --example record_application -- vlc.exe

use std::collections::VecDeque;
use std::error::{self};
use std::fs::File;
use std::io::prelude::*;
use std::process::Command;
use std::sync::mpsc;

use std::thread;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;
type Res<T> = Result<T, Box<dyn error::Error>>;

// Capture loop, capture samples and send in chunks of "chunksize" frames to channel
fn capture_loop(
    tx_capt: std::sync::mpsc::SyncSender<Vec<u8>>,
    chunksize: usize,
    process_id: u32,
) -> Res<()> {
    initialize_mta().ok().unwrap();

    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 48000, 2, None);
    let blockalign = desired_format.get_blockalign();
    debug!("Desired capture format: {:?}", desired_format);
    let autoconvert = true;
    let include_tree = true;

    let mut audio_client = AudioClient::new_application_loopback_client(process_id, include_tree)?;
    let mode = StreamMode::EventsShared {
        autoconvert,
        buffer_duration_hns: 0,
    };
    audio_client.initialize_client(&desired_format, &Direction::Capture, &mode)?;

    debug!("initialized capture");
    let h_event = audio_client.set_get_eventhandle().unwrap();

    let capture_client = audio_client.get_audiocaptureclient().unwrap();

    // just eat the reallocation because querying the buffer size gives massive values.
    let mut sample_queue: VecDeque<u8> = VecDeque::new();

    audio_client.start_stream().unwrap();

    loop {
        while sample_queue.len() > (blockalign as usize * chunksize) {
            debug!("pushing samples");
            let mut chunk = vec![0u8; blockalign as usize * chunksize];
            // One byte at a time, to keep the example simple. This costs a call and an
            // index update per byte, which starts to matter at a high sample rate and
            // channel count. A deque holds its contents as one or two contiguous runs,
            // so an application that needs the throughput should get them from
            // `as_slices`, copy each one with `copy_from_slice` and then drain the front.
            for element in chunk.iter_mut() {
                *element = sample_queue.pop_front().unwrap();
            }
            tx_capt.send(chunk).unwrap();
        }
        trace!("capturing");

        let new_frames = capture_client.get_next_packet_size()?.unwrap_or(0);
        let additional = (new_frames as usize * blockalign as usize)
            .saturating_sub(sample_queue.capacity() - sample_queue.len());
        sample_queue.reserve(additional);
        if new_frames > 0 {
            capture_client
                .read_from_device_to_deque(&mut sample_queue)
                .unwrap();
        }
        if h_event.wait_for_event(3000).is_err() {
            error!("timeout error, stopping capture");
            audio_client.stop_stream().unwrap();
            break;
        }
    }
    Ok(())
}

// Find the main process of an application, the one its other processes were started from.
// Capturing with `include_tree` needs this one, since a child process only covers its own subtree.
// The lookup asks PowerShell, which ships with Windows, to keep the example free of dependencies.
// It picks the process whose parent is not another process with the same name.
// If the application runs as several independent instances, it takes the first one.
// The name goes in through an environment variable, so that it is never parsed as part of the script.
fn find_main_process(name: &str) -> Res<u32> {
    // Win32_Process names include the extension.
    let name = if name.to_lowercase().ends_with(".exe") {
        name.to_string()
    } else {
        format!("{name}.exe")
    };
    let script = "$p = Get-CimInstance Win32_Process | Where-Object Name -eq $env:PROCESS_NAME; \
                  ($p | Where-Object { $p.ProcessId -notcontains $_.ParentProcessId }).ProcessId";
    let output = Command::new("powershell")
        .args(["-NoProfile", "-Command", script])
        .env("PROCESS_NAME", &name)
        .output()?;
    let stderr = String::from_utf8_lossy(&output.stderr);
    if !output.status.success() || !stderr.trim().is_empty() {
        return Err(format!("the process lookup failed: {}", stderr.trim()).into());
    }
    let stdout = String::from_utf8(output.stdout)?;
    let pids = stdout
        .split_whitespace()
        .map(str::parse)
        .collect::<Result<Vec<u32>, _>>()?;
    let pid = *pids
        .first()
        .ok_or(format!("no process named {name} is running"))?;
    if pids.len() > 1 {
        warn!("{name} runs as several separate instances {pids:?}, capturing {pid}");
    }
    Ok(pid)
}

// Main loop
fn main() -> Res<()> {
    let _ = SimpleLogger::init(
        LevelFilter::Trace,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );

    let name = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "firefox.exe".to_string());
    let process_id = find_main_process(&name)?;

    info!("Capture Process ID: {}", process_id);

    let (tx_capt, rx_capt): (
        std::sync::mpsc::SyncSender<Vec<u8>>,
        std::sync::mpsc::Receiver<Vec<u8>>,
    ) = mpsc::sync_channel(2);
    let chunksize = 4096;

    // Capture
    let _handle = thread::Builder::new()
        .name("Capture".to_string())
        .spawn(move || {
            let result = capture_loop(tx_capt, chunksize, process_id);
            if let Err(err) = result {
                error!("Capture failed with error {}", err);
            }
        });

    let mut outfile = File::create("recorded.raw")?;
    info!("Saving captured raw data to 'recorded.raw'");

    loop {
        match rx_capt.recv() {
            Ok(chunk) => {
                debug!("writing to file");
                outfile.write_all(&chunk)?;
            }
            Err(err) => {
                error!("Some error {}", err);
                return Ok(());
            }
        }
    }
}
