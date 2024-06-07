use std::collections::VecDeque;
use std::error::{self};
use std::fs::File;
use std::io::prelude::*;
use std::sync::mpsc;

use std::thread;
use sysinfo::{ProcessRefreshKind, RefreshKind, System};
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
    audio_client.initialize_client(
        &desired_format,
        0,
        &Direction::Capture,
        &ShareMode::Shared,
        autoconvert,
    )?;

    debug!("initialized capture");
    let h_event = audio_client.set_get_eventhandle().unwrap();

    let capture_client = audio_client.get_audiocaptureclient().unwrap();

    let mut sample_queue: VecDeque<u8> = VecDeque::new(); // just eat the reallocation because querying the buffer size gives massive values.

    audio_client.start_stream().unwrap();

    loop {
        while sample_queue.len() > (blockalign as usize * chunksize) {
            debug!("pushing samples");
            let mut chunk = vec![0u8; blockalign as usize * chunksize];
            for element in chunk.iter_mut() {
                *element = sample_queue.pop_front().unwrap();
            }
            tx_capt.send(chunk).unwrap();
        }
        trace!("capturing");

        let new_frames = capture_client.get_next_nbr_frames()?.unwrap_or(0);
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

    let refreshes = RefreshKind::new().with_processes(ProcessRefreshKind::everything());
    let system = System::new_with_specifics(refreshes);
    let process_ids = system.processes_by_name("firefox.exe");
    let mut process_id = 0;
    for process in process_ids {
        // Note: When capturing audio windows allows you to capture an app's entire process tree, however you must ensure you use the parent as the target process ID
        process_id = process.parent().unwrap_or(process.pid()).as_u32();
    }

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
