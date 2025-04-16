use std::collections::VecDeque;
use std::error;
use std::sync::mpsc;
use std::thread;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;

type Res<T> = Result<T, Box<dyn error::Error>>;

// Playback loop, play samples received from channel
fn playback_loop(rx_play: std::sync::mpsc::Receiver<Vec<u8>>) -> Res<()> {
    let device = get_default_device(&Direction::Render)?;
    let mut audio_client = device.get_iaudioclient()?;
    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);

    let blockalign = desired_format.get_blockalign();
    debug!("Desired playback format: {:?}", desired_format);

    let (def_time, min_time) = audio_client.get_device_period()?;
    debug!("default period {}, min period {}", def_time, min_time);

    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: min_time,
    };
    audio_client.initialize_client(&desired_format, &Direction::Render, &mode)?;
    debug!("initialized playback");

    let h_event = audio_client.set_get_eventhandle()?;

    let mut buffer_frame_count = audio_client.get_buffer_size()?;

    let render_client = audio_client.get_audiorenderclient()?;
    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(
        100 * blockalign as usize * (1024 + 2 * buffer_frame_count as usize),
    );
    audio_client.start_stream()?;
    loop {
        buffer_frame_count = audio_client.get_available_space_in_frames()?;
        trace!("New buffer frame count {}", buffer_frame_count);
        while sample_queue.len() < (blockalign as usize * buffer_frame_count as usize) {
            debug!("need more samples");
            match rx_play.try_recv() {
                Ok(chunk) => {
                    trace!("got chunk");
                    for element in chunk.iter() {
                        sample_queue.push_back(*element);
                    }
                }
                Err(mpsc::TryRecvError::Empty) => {
                    warn!("no data, filling with zeros");
                    for _ in 0..((blockalign as usize * buffer_frame_count as usize)
                        - sample_queue.len())
                    {
                        sample_queue.push_back(0);
                    }
                }
                Err(_) => {
                    error!("Channel is closed");
                    break;
                }
            }
        }

        trace!("write");
        render_client.write_to_device_from_deque(
            buffer_frame_count as usize,
            &mut sample_queue,
            None,
        )?;
        trace!("write ok");
        if h_event.wait_for_event(100000).is_err() {
            error!("error, stopping playback");
            audio_client.stop_stream()?;
            break;
        }
    }
    Ok(())
}

// Capture loop, capture samples and send in chunks of "chunksize" frames to channel
fn capture_loop(tx_capt: std::sync::mpsc::SyncSender<Vec<u8>>, chunksize: usize) -> Res<()> {
    let device = get_default_device(&Direction::Capture)?;
    let mut audio_client = device.get_iaudioclient()?;

    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2, None);

    let blockalign = desired_format.get_blockalign();
    debug!("Desired capture format: {:?}", desired_format);

    let (def_time, min_time) = audio_client.get_device_period()?;
    debug!("default period {}, min period {}", def_time, min_time);

    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: min_time,
    };
    audio_client.initialize_client(&desired_format, &Direction::Capture, &mode)?;
    debug!("initialized capture");

    let h_event = audio_client.set_get_eventhandle()?;

    let buffer_frame_count = audio_client.get_buffer_size()?;

    let render_client = audio_client.get_audiocaptureclient()?;
    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(
        100 * blockalign as usize * (1024 + 2 * buffer_frame_count as usize),
    );
    audio_client.start_stream()?;
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
        if h_event.wait_for_event(1000000).is_err() {
            error!("error, stopping capture");
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

    let (tx_play, rx_play): (
        std::sync::mpsc::SyncSender<Vec<u8>>,
        std::sync::mpsc::Receiver<Vec<u8>>,
    ) = mpsc::sync_channel(2);
    let (tx_capt, rx_capt): (
        std::sync::mpsc::SyncSender<Vec<u8>>,
        std::sync::mpsc::Receiver<Vec<u8>>,
    ) = mpsc::sync_channel(2);
    let chunksize = 4096;

    // Playback
    let _handle = thread::Builder::new()
        .name("Player".to_string())
        .spawn(move || {
            let result = playback_loop(rx_play);
            if let Err(err) = result {
                error!("Playback failed with error {}", err);
            }
        });

    // Capture
    let _handle = thread::Builder::new()
        .name("Capture".to_string())
        .spawn(move || {
            let result = capture_loop(tx_capt, chunksize);
            if let Err(err) = result {
                error!("Capture failed with error {}", err);
            }
        });

    loop {
        match rx_capt.recv() {
            Ok(chunk) => {
                debug!("sending");
                tx_play.send(chunk).unwrap();
            }
            Err(err) => error!("Some error {}", err),
        }
    }
}
