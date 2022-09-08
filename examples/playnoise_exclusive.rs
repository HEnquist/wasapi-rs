use rand::prelude::*;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;

// Main loop
fn main() {
    let _ = SimpleLogger::init(
        LevelFilter::Debug,
        ConfigBuilder::new()
            .set_time_format_str("%H:%M:%S%.3f")
            .build(),
    );

    initialize_mta().unwrap();

    let channels = 2;
    let device = get_default_device(&Direction::Render).unwrap();
    let mut audio_client = device.get_iaudioclient().unwrap();
    let desired_format = WaveFormat::new(24, 24, &SampleType::Int, 44100, channels, None);
    //let desired_format = desired_format.to_waveformatex().unwrap();

    //desired_format.wave_fmt.Format.cbSize = 0;
    //desired_format.wave_fmt.Format.wFormatTag = WAVE_FORMAT_PCM as u16;
    //desired_format.wave_fmt.dwChannelMask = 0;

    // Make sure the format is supported, panic if not.
    let desired_format = audio_client
        .is_supported_exclusive_with_quirks(&desired_format)
        .unwrap();

    // Blockalign is the number of bytes per frame
    let blockalign = desired_format.get_blockalign();
    debug!("Desired playback format: {:?}", desired_format);

    let (def_period, min_period) = audio_client.get_periods().unwrap();

    // Set some period as an example, using 128 byte alignment to satisfy for example Intel HDA
    let desired_period = audio_client.calculate_aligned_period_near(3*min_period/2, Some(128), &desired_format).unwrap();

    debug!("periods in 100ns units {}, minimum {}, wanted {}", def_period, min_period, desired_period);


    audio_client
        .initialize_client(
            &desired_format,
            desired_period as i64,
            &Direction::Render,
            &ShareMode::Exclusive,
            false,
        )
        .unwrap();
    debug!("initialized playback");

    let mut rng = rand::thread_rng();

    let h_event = audio_client.set_get_eventhandle().unwrap();

    let render_client = audio_client.get_audiorenderclient().unwrap();

    audio_client.start_stream().unwrap();
    loop {
        let buffer_frame_count = audio_client.get_available_space_in_frames().unwrap();

        let mut data = vec![0u8; buffer_frame_count as usize * blockalign as usize];
        for frame in data.chunks_exact_mut(blockalign as usize) {
            let sample: u32 = rng.gen();
            let sample_bytes = sample.to_le_bytes();
            for value in frame.chunks_exact_mut(blockalign as usize / channels as usize) {
                for (bufbyte, samplebyte) in value.iter_mut().zip(sample_bytes.iter()) {
                    *bufbyte = *samplebyte;
                }
            }
        }

        trace!("write");
        render_client
            .write_to_device(
                buffer_frame_count as usize,
                blockalign as usize,
                &data,
                None,
            )
            .unwrap();
        trace!("write ok");
        if h_event.wait_for_event(1000).is_err() {
            error!("error, stopping playback");
            audio_client.stop_stream().unwrap();
            break;
        }
    }
}
