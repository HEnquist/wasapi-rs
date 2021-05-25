use std::collections::VecDeque;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::sync::mpsc;
use std::sync::{Arc, Barrier, RwLock};
use std::thread;
use windows::initialize_mta;
use std::error;
use wasapi::wasapi::*;

type Res<T> = Result<T, Box<dyn error::Error>>;

// Playback loop, play samples received from channel
fn playback_loop(rx_play: std::sync::mpsc::Receiver<Vec<u8>>) -> Res<()> {
    let collection = DeviceCollection::new(&Direction::Render)?;
    let device = collection.get_device_with_name("SPDIF Interface (FX-AUDIO-DAC-X6)")?;
    let mut audio_client = device.get_iaudioclient()?;
    // int16
    //let desired_format_ex = WaveFormat::new(16, 16, &SampleType::Int, 44100, 2);
    //let sharemode = ShareMode::Exclusive;
    // float32
    let desired_format_ex = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2);
    let sharemode = ShareMode::Shared;
    

    let blockalign = desired_format_ex.get_blockalign();
    desired_format_ex.print_waveformat();

    

    let supported_format = match audio_client.is_supported(&desired_format_ex, &sharemode)? {
        FormatSupported::Yes => desired_format_ex,
        FormatSupported::ClosestMatch(modified_format) => modified_format,
    };
    supported_format.print_waveformat();

    let (def_time, min_time) = audio_client.get_periods()?;
    println!("default period {}, min period {}", def_time, min_time);


    audio_client.initialize_client(&supported_format, min_time as i64, &Direction::Render, &sharemode)?;
    println!("initialized playback");

    let h_event = audio_client.set_get_eventhandle()?;

    let mut buffer_frame_count = audio_client.get_bufferframecount()?;

    let render_client = audio_client.get_audiorenderclient()?;
    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(100*blockalign as usize * (1024 + 2*buffer_frame_count as usize));
    audio_client.start_stream()?;
    loop {
        buffer_frame_count = audio_client.get_available_frames()?;
        println!("New buffer frame count {}", buffer_frame_count);
        while sample_queue.len() < (blockalign as usize * buffer_frame_count as usize) {
            println!("need more samples");
            match rx_play.try_recv() {
                Ok(chunk) => {
                    println!("got chunk");
                    for element in chunk.iter() {
                        sample_queue.push_back(*element);
                    }
                }
                Err(mpsc::TryRecvError::Empty) => {
                    println!("no data, filling with zeros");
                    for _ in 0..((blockalign as usize * buffer_frame_count as usize) - sample_queue.len())  {
                        sample_queue.push_back(0);
                    }
                }
                Err(_) => {
                    println!("oops");
                    break;
                }
            }
            println!("deque len2 {}", sample_queue.len());
        }
        //println!("wait for buf");

        println!("write");
        render_client.write_to_device_from_deque(buffer_frame_count as usize, blockalign as usize, &mut sample_queue )?;
        println!("write ok");
        if h_event.wait_for_event(100000).is_err() {
            println!("error, stopping playback");
            audio_client.stop_stream()?;
            break;
        }
    }
    Ok(())
}


// Capture loop, capture samples and send in chunks of "chunksize" frames to channel
fn capture_loop(tx_capt: std::sync::mpsc::SyncSender<Vec<u8>>, chunksize: usize) -> Res<()> {
    let collection = DeviceCollection::new(&Direction::Capture)?;
    let device = collection.get_device_with_name("CABLE Output (VB-Audio Virtual Cable)")?;
    let mut audio_client = device.get_iaudioclient()?;

    // int16
    //let desired_format_ex = WaveFormat::new(16, 16, &SampleType::Int, 44100, 2);
    //let sharemode = ShareMode::Exclusive;
    // float32
    let desired_format_ex = WaveFormat::new(32, 32, &SampleType::Float, 44100, 2);
    let sharemode = ShareMode::Shared;

    let blockalign = desired_format_ex.get_blockalign();
    println!("\nCapture requested");
    desired_format_ex.print_waveformat();

    let supported_format = match audio_client.is_supported(&desired_format_ex, &sharemode)? {
        FormatSupported::Yes => desired_format_ex,
        FormatSupported::ClosestMatch(modified_format) => modified_format,
    };
    println!("\nCapture got");
    supported_format.print_waveformat();
    let (def_time, min_time) = audio_client.get_periods()?;
    println!("default period {}, min period {}", def_time, min_time);


    audio_client.initialize_client(&supported_format, min_time as i64, &Direction::Capture, &sharemode)?;
    println!("initialized capture");

    let h_event = audio_client.set_get_eventhandle()?;

    let buffer_frame_count = audio_client.get_bufferframecount()?;

    let render_client = audio_client.get_audiocaptureclient()?;
    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(100*blockalign as usize * (1024 + 2*buffer_frame_count as usize));
    audio_client.start_stream()?;
    loop {
        //println!("deque len {}", sample_queue.len());
        while sample_queue.len() > (blockalign as usize * chunksize as usize) {
            println!("pushing samples");
            let mut chunk = vec![0u8; blockalign as usize * chunksize as usize];
            for element in chunk.iter_mut() {
                *element = sample_queue.pop_front().unwrap();
            }
            tx_capt.send(chunk)?;
        }
        println!("capturing");
        render_client.read_from_device_to_deque(blockalign as usize, &mut sample_queue)?;
        println!("captured");
        if h_event.wait_for_event(1000000).is_err() {
            println!("error, stopping capture");
            audio_client.stop_stream()?;
            break;
        }
    }
    Ok(())
}

// Main loop
fn main() -> Res<()> {
    initialize_mta()?;
    let (tx_play, rx_play): (std::sync::mpsc::SyncSender<Vec<u8>>, std::sync::mpsc::Receiver<Vec<u8>>) = mpsc::sync_channel(2);
    let (tx_capt, rx_capt): (std::sync::mpsc::SyncSender<Vec<u8>>, std::sync::mpsc::Receiver<Vec<u8>>) = mpsc::sync_channel(2);
    let buffer_fill = Arc::new(AtomicUsize::new(0));
    let buffer_fill_clone = buffer_fill.clone();
    let chunksize = 4096;
    
    // Playback
    let _handle = thread::Builder::new()
        .name("Player".to_string())
        .spawn(move || {
            let result = playback_loop(rx_play);
            if let Err(err) = result {
                println!("Playback failed with error {}", err);
            }
        });

    // Capture
    let _handle = thread::Builder::new()
        .name("Capture".to_string())
        .spawn(move || {
            let result = capture_loop(tx_capt, chunksize);
            if let Err(err) = result {
                println!("Capture failed with error {}", err);
            }
        });

    loop {
        match rx_capt.recv() {
            Ok(chunk) => {
                println!("sending");
                tx_play.send(chunk).unwrap();
            },
            Err(err) => println!("Some error {}", err),
        }
        //return Ok(());
    }
}
