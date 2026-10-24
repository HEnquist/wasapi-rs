// Play two sine waves on the default output device, each in its own named audio session.
//
// Each stream is initialized with its own session GUID, so the volume mixer shows two
// entries for this one process, one per stream, with the names set here instead of
// the name of the executable. Plays for 20 seconds, which leaves time to open the
// volume mixer and change the volume of each stream separately.

use std::f64::consts::PI;
use std::thread;
use std::time::{Duration, Instant};
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;

const PLAY_TIME: Duration = Duration::from_secs(20);

// Fixed GUIDs, so that the sessions are the same each time the example runs.
const LOW_SESSION: GUID = GUID::from_u128(0x5d0c_2b1e_8a43_4f6e_9c27_1e4a_b3f0_7d61);
const HIGH_SESSION: GUID = GUID::from_u128(0x5d0c_2b1e_8a43_4f6e_9c27_1e4a_b3f0_7d62);

fn play_sine(freq: f64, session: GUID, name: &str) -> Result<(), WasapiError> {
    initialize_mta().ok()?;
    let device = DeviceEnumerator::new()?.get_default_device(&Direction::Render)?;
    let mut audio_client = device.get_iaudioclient()?;
    let channels = 2;
    let format = WaveFormat::new(32, 32, &SampleType::Float, 44100, channels, None);
    let blockalign = format.get_blockalign() as usize;
    let (def_time, _min_time) = audio_client.get_device_period()?;
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: def_time,
    };
    audio_client.initialize_client(&format, &Direction::Render, &mode, Some(session))?;
    audio_client
        .get_audiosessioncontrol()?
        .set_display_name(name)?;
    debug!("{name}: initialized in session {session:?}");

    let h_event = audio_client.set_get_eventhandle()?;
    let render_client = audio_client.get_audiorenderclient()?;
    let mut phase = 0.0f64;
    let mut write_frames = |nbr_frames: usize| {
        let mut data = vec![0u8; nbr_frames * blockalign];
        for frame in data.chunks_exact_mut(blockalign) {
            phase = (phase + freq / 44100.0).fract();
            let sample = ((phase * 2.0 * PI).sin() * 0.1) as f32;
            for value in frame.chunks_exact_mut(blockalign / channels) {
                value.copy_from_slice(&sample.to_le_bytes());
            }
        }
        render_client.write_to_device(nbr_frames, &data, None)
    };

    // Fill the buffer before starting the stream, see the playsine example.
    write_frames(audio_client.get_available_space_in_frames()? as usize)?;
    audio_client.start_stream()?;
    let start = Instant::now();
    while start.elapsed() < PLAY_TIME {
        h_event.wait_for_event(1000)?;
        write_frames(audio_client.get_available_space_in_frames()? as usize)?;
    }
    audio_client.stop_stream()
}

fn main() {
    let _ = SimpleLogger::init(LevelFilter::Debug, Config::default());

    let low = thread::spawn(|| play_sine(440.0, LOW_SESSION, "wasapi-rs low tone"));
    let high = thread::spawn(|| play_sine(880.0, HIGH_SESSION, "wasapi-rs high tone"));
    for (name, handle) in [("low", low), ("high", high)] {
        match handle.join().unwrap() {
            Ok(()) => info!("{name} tone done"),
            Err(err) => error!("{name} tone failed: {err}"),
        }
    }
}
