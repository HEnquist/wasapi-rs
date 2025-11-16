use wasapi::*;

fn main() {
    initialize_mta().unwrap();

    let enumerator = DeviceEnumerator::new().unwrap();

    println!("The following input devices are being used by:");
    for device in &enumerator
        .get_device_collection(&Direction::Capture)
        .unwrap()
    {
        let dev = device.unwrap();
        let manager = dev.get_iaudiosessionmanager().unwrap();
        let enumerator = manager.get_audiosessionenumerator().unwrap();

        println!("Device: {:?}", &dev.get_friendlyname().unwrap());

        for i in 0..enumerator.get_count().unwrap() {
            let control = enumerator.get_session(i).unwrap();
            let process_id = control.get_process_id().unwrap();

            println!(" - In use by process: {:?}", process_id);
        }
    }
}
