// Play a sine wave in shared mode on the default output device, with automatic stream routing.
//
// The stream follows the default device. Change the default output device in the Windows
// sound settings while this plays, and the sound moves over to the new device.

use std::f64::consts::PI;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;

fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Debug,
        ConfigBuilder::new()
            .set_time_format_rfc3339()
            .set_time_offset_to_local()
            .unwrap()
            .build(),
    );

    initialize_mta().unwrap();

    let mut audio_client = AudioClient::new_default_device_client(Direction::Render).unwrap();

    // Always use the same format, and let the engine convert to whatever the current
    // default device wants.
    let channels = 2;
    let samplerate = 44100;
    let format = WaveFormat::new(32, 32, SampleType::Float, samplerate, channels, None);
    let blockalign = format.get_blockalign() as usize;
    let (def_time, _min_time) = audio_client.get_device_period().unwrap();
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: def_time,
    };
    audio_client
        .initialize_client(&format, Direction::Render, mode, None)
        .unwrap();
    let h_event = audio_client.set_get_eventhandle().unwrap();
    let render_client = audio_client.get_audiorenderclient().unwrap();

    let delta_phase = 2.0 * PI * 1000.0 / samplerate as f64;
    let mut phase = 0.0f64;
    let mut write_frames = |nbr_frames: usize| -> Result<(), WasapiError> {
        let mut data = vec![0u8; nbr_frames * blockalign];
        for frame in data.chunks_exact_mut(blockalign) {
            let sample = (0.1 * phase.sin()) as f32;
            phase = (phase + delta_phase) % (2.0 * PI);
            for value in frame.chunks_exact_mut(4) {
                value.copy_from_slice(&sample.to_le_bytes());
            }
        }
        render_client.write_to_device(&data, None)
    };

    write_frames(audio_client.get_available_space_in_frames().unwrap() as usize).unwrap();
    audio_client.start_stream().unwrap();
    info!("Playing, change the default output device to move the stream");
    loop {
        if let Err(err) = h_event.wait_for_event(1000) {
            error!("Waiting for the event failed, {err}");
            break;
        }
        let result = audio_client
            .get_available_space_in_frames()
            .and_then(|frames| write_frames(frames as usize));
        if let Err(err) = result {
            error!("Writing to the device failed, {err}");
            break;
        }
    }
    audio_client.stop_stream().ok();
}
