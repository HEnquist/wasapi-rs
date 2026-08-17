// List all audio devices and the processes that are using them.
//
// Prints the peak level of each device, and of every active session on it.

use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();

    for direction in [Direction::Capture, Direction::Render] {
        println!("The following {direction} devices are being used by:");
        for device in &enumerator.get_device_collection(&direction).unwrap() {
            let dev = device.unwrap();
            let manager = dev.get_iaudiosessionmanager().unwrap();
            let sessions = manager.get_audiosessionenumerator().unwrap();

            let dev_meter = dev.get_audiometerinformation().unwrap();
            let dev_peak = dev_meter.get_peak_value().unwrap();

            println!(
                "Device: {:?}, peak: {dev_peak:.3}",
                dev.get_friendlyname().unwrap()
            );

            for i in 0..sessions.get_count().unwrap() {
                let control = sessions.get_session(i).unwrap();
                let state = control.get_state().unwrap();
                if state != SessionState::Active {
                    continue;
                }
                let process_id = control.get_process_id().unwrap();
                let identifier = control.get_session_identifier().unwrap();
                let meter = control.get_audiometerinformation().unwrap();
                let peak = meter.get_peak_value().unwrap();

                println!(" - In use by process: {process_id}, peak: {peak:.3}");
                println!("   session: {identifier}");
            }
        }
    }
}
