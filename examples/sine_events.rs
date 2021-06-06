use std::f64::consts::PI;
use wasapi::*;
use windows::initialize_mta;

#[macro_use]
extern crate log;
use simplelog::*;

struct SineGenerator {
    time: f64,
    freq: f64,
    delta_t: f64,
    amplitude: f64,
}

impl SineGenerator {
    fn new(freq: f64, fs: f64, amplitude: f64) -> Self {
        SineGenerator {
            time: 0.0,
            freq,
            delta_t: 1.0/fs,
            amplitude,
        }
    }
}

impl Iterator for SineGenerator {
    type Item = f32;
    fn next(&mut self) -> Option<f32> {
        self.time += self.delta_t;
        let output = ((self.freq * self.time * PI * 2.).sin() * self.amplitude) as f32;
        Some(output)
    }
}


// Main loop
fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Debug,
        ConfigBuilder::new()
            .set_time_format_str("%H:%M:%S%.3f")
            .build(),
    );

    initialize_mta().unwrap();

    let mut gen = SineGenerator::new(1000.0, 44100.0, 0.1);

    let channels = 2;
    let device = get_default_device(&Direction::Render).unwrap();
    let mut audio_client = device.get_iaudioclient().unwrap();
    let desired_format = WaveFormat::new(32, 32, &SampleType::Float, 44100, channels);

    let blockalign = desired_format.get_blockalign();
    debug!("Desired playback format: {:?}", desired_format);

    let (def_time, min_time) = audio_client.get_periods().unwrap();
    debug!("default period {}, min period {}", def_time, min_time);

    audio_client.initialize_client(
        &desired_format,
        def_time as i64,
        &Direction::Render,
        &ShareMode::Shared,
        true,
    ).unwrap();
    debug!("initialized playback");

    let h_event = audio_client.set_get_eventhandle().unwrap();

    let render_client = audio_client.get_audiorenderclient().unwrap();

    let mut callbacks = EventCallbacks::new();

    callbacks.set_simple_volume_callback(|vol, mute| println!("New simple volume {}, mute {}", vol, mute));
    callbacks.set_state_callback(|state| println!("New state {:?}", state));
    callbacks.set_channel_volume_callback(|index, vol| println!("New channel volume {}, channel {}", vol, index));
    callbacks.set_disconnected_callback(|reason| println!("Disconnected, reason {:?}", reason));

    let sessioncontrol = audio_client.get_audiosessioncontrol().unwrap();
    sessioncontrol.register_session_notification(callbacks).unwrap();

    audio_client.start_stream().unwrap();
    loop {
        let buffer_frame_count = audio_client.get_available_space_in_frames().unwrap();

        let mut data = vec![0u8; buffer_frame_count as usize * blockalign as usize];
        for frame in data.chunks_exact_mut(blockalign as usize) {
            let sample = gen.next().unwrap();
            let sample_bytes = sample.to_le_bytes();
            for value in frame.chunks_exact_mut(blockalign as usize/channels as usize) {
                for (bufbyte, sinebyte) in value.iter_mut().zip(sample_bytes.iter()) {
                    *bufbyte = *sinebyte;
                }
            }
        }

        trace!("write");
        render_client.write_to_device(
            buffer_frame_count as usize,
            blockalign as usize,
            &data,
        ).unwrap();
        trace!("write ok");
        if h_event.wait_for_event(1000).is_err() {
            error!("error, stopping playback");
            audio_client.stop_stream().unwrap();
            break;
        }
    }
}
