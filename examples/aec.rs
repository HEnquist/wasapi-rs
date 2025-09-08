use std::collections::VecDeque;
use std::error;
use std::fs::File;
use std::io::prelude::*;
use std::sync::mpsc;
use std::thread;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;

type Res<T> = Result<T, Box<dyn error::Error>>;

// Capture loop, capture samples and send in chunks of "chunksize" frames to channel
fn capture_loop(tx_capt: std::sync::mpsc::SyncSender<Vec<u8>>, chunksize: usize) -> Res<()> {
    let input_device = get_default_device_for_role(&Direction::Capture, &Role::Communications)?;

    let mut audio_client = input_device.get_iaudioclient()?;

    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);

    let blockalign = desired_format.get_blockalign();
    debug!("Desired capture format: {:?}", desired_format);

    let (def_time, min_time) = audio_client.get_device_period()?;
    debug!("default period {}, min period {}", def_time, min_time);

    // Set the category as communications, so that audio effects like AEC can be applied.
    // this category is not valid for speaker loopback stream, so only set it for input capture.
    let properties = AudioClientProperties::new().set_category(StreamCategory::Communications);
    audio_client.set_properties(properties)?;

    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: min_time,
    };
    audio_client.initialize_client(&desired_format, &Direction::Capture, &mode)?;
    debug!("initialized capture");

    // Enable Acoustic Echo Cancellation if it is supported.
    if audio_client.is_aec_supported()? {
        let aec_ctrl = audio_client.get_aec_control()?;

        let output_device = get_default_device(&Direction::Render)?;
        let render_endpoint_id = output_device.get_id()?;
        // Pass the endpoint id of the audio render endpoint that should be used as the reference stream for AEC.
        aec_ctrl.set_echo_cancellation_render_endpoint(Some(render_endpoint_id))?;
        debug!(
            "AEC enabled with render endpoint: {}",
            output_device.get_interface_friendlyname()?
        );
    } else {
        warn!(
            "AEC not supported on input device: {}",
            input_device.get_interface_friendlyname()?
        );
    }

    let h_event = audio_client.set_get_eventhandle()?;

    let buffer_frame_count = audio_client.get_buffer_size()?;

    let render_client = audio_client.get_audiocaptureclient()?;
    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(
        100 * blockalign as usize * (1024 + 2 * buffer_frame_count as usize),
    );
    let session_control = audio_client.get_audiosessioncontrol()?;

    debug!("state before start: {:?}", session_control.get_state());
    audio_client.start_stream()?;
    debug!("state after start: {:?}", session_control.get_state());

    loop {
        while sample_queue.len() > (blockalign as usize * chunksize) {
            debug!("pushing samples");
            let mut chunk = vec![0u8; blockalign as usize * chunksize];
            for element in chunk.iter_mut() {
                *element = sample_queue.pop_front().unwrap();
            }
            tx_capt.send(chunk)?;
        }
        trace!("capturing");
        render_client.read_from_device_to_deque(&mut sample_queue)?;
        if h_event.wait_for_event(3000).is_err() {
            error!("timeout error, stopping capture");
            audio_client.stop_stream()?;
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

    initialize_mta().ok()?;

    let (tx_capt, rx_capt): (
        std::sync::mpsc::SyncSender<Vec<u8>>,
        std::sync::mpsc::Receiver<Vec<u8>>,
    ) = mpsc::sync_channel(2);
    let chunksize = 4096;

    // Capture
    let _handle = thread::Builder::new()
        .name("Capture".to_string())
        .spawn(move || {
            let result = capture_loop(tx_capt, chunksize);
            if let Err(err) = result {
                error!("Capture failed with error {}", err);
            }
        });

    let mut outfile = File::create("aec-recorded.raw")?;
    info!("Saving captured raw data to 'aec-recorded.raw'");

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
