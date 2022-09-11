use rand::prelude::*;
use wasapi::*;

#[macro_use]
extern crate log;
use simplelog::*;
use windows::core::Error;

// A selection of the possible errors
use windows::Win32::Foundation::E_INVALIDARG;
use windows::Win32::Media::Audio::{
    AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED, AUDCLNT_E_DEVICE_IN_USE, AUDCLNT_E_ENDPOINT_CREATE_FAILED,
    AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED, AUDCLNT_E_UNSUPPORTED_FORMAT,
};

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

    // Make sure the format is supported, panic if not.
    let desired_format = audio_client
        .is_supported_exclusive_with_quirks(&desired_format)
        .unwrap();

    // Blockalign is the number of bytes per frame
    let blockalign = desired_format.get_blockalign();
    debug!("Desired playback format: {:?}", desired_format);

    let (def_period, min_period) = audio_client.get_periods().unwrap();

    // Set some period as an example, using 128 byte alignment to satisfy for example Intel HDA
    let desired_period = audio_client
        .calculate_aligned_period_near(3 * min_period / 2, Some(128), &desired_format)
        .unwrap();

    debug!(
        "periods in 100ns units {}, minimum {}, wanted {}",
        def_period, min_period, desired_period
    );

    let init_result = audio_client.initialize_client(
        &desired_format,
        desired_period as i64,
        &Direction::Render,
        &ShareMode::Exclusive,
        false,
    );
    match init_result {
        Ok(()) => debug!("IAudioClient::Initialize ok"),
        Err(e) => {
            if let Some(werr) = e.downcast_ref::<Error>() {
                // Some of the possible errors. See the documentation for the full list and descriptions.
                // https://docs.microsoft.com/en-us/windows/win32/api/audioclient/nf-audioclient-iaudioclient-initialize
                match werr.code() {
                    E_INVALIDARG => error!("IAudioClient::Initialize: Invalid argument"),
                    AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED => {
                        error!("IAudioClient::Initialize: Unaligned buffer, adjust the period")
                    }
                    AUDCLNT_E_DEVICE_IN_USE => {
                        error!("IAudioClient::Initialize: The device is already in use")
                    }
                    AUDCLNT_E_UNSUPPORTED_FORMAT => error!(
                        "IAudioClient::Initialize The device does not support the audio format"
                    ),
                    AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED => {
                        error!("IAudioClient::Initialize: Exclusive mode is not allowed")
                    }
                    AUDCLNT_E_ENDPOINT_CREATE_FAILED => {
                        error!("IAudioClient::Initialize: Failed to create endpoint")
                    }
                    _ => error!(
                        "IAudioClient::Initialize: Other error, HRESULT: {:#010x}, info: {:?}",
                        werr.code().0,
                        werr.info()
                    ),
                };
                panic!("IAudioClient::Initialize failed")
            } else {
                panic!("IAudioClient::Initialize: Other error {:?}", e);
            }
        }
    };

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
