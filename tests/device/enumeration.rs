// Device enumeration.
//
// No audio is played or captured here. What is worth testing is the hand-written part:
// the `DeviceCollection` iterator, which tracks its index itself, and the two lookups
// that have to find their way back to the same device. The property accessors are
// one-line calls into Windows, so they are printed rather than asserted.

use crate::support;
use wasapi::*;

/// The cable has to be there when CI said it would be.
///
/// Every other test skips when the cable is missing, so without this one a device job
/// where the install silently failed would be green and prove nothing.
#[test]
fn cable_is_present() {
    support::init_com();
    let enumerator = DeviceEnumerator::new().unwrap();
    let render = support::cable_endpoints(&enumerator, &Direction::Render);
    let capture = support::cable_endpoints(&enumerator, &Direction::Capture);
    println!(
        "found {} render and {} capture endpoints of \"{}\"",
        render.len(),
        capture.len(),
        support::CABLE_INTERFACE_NAME
    );
    if render.is_empty() || capture.is_empty() {
        support::skip("VB-Cable is not installed");
    }
}

/// Print every endpoint, as a debugging aid rather than a test.
///
/// Run with --nocapture, which the CI job does. When a device test fails, this is the
/// record of what the runner actually had: the endpoint names after a driver only
/// install, whether the 16 channel endpoint exists, and the formats and data ranges
/// each endpoint reports.
#[test]
fn endpoint_inventory() {
    support::init_com();
    let enumerator = DeviceEnumerator::new().unwrap();
    let describe_format = |format: Result<WaveFormat, WasapiError>| match format {
        Ok(f) => format!(
            "{} ch, {} Hz, {}/{} bits, {:?}",
            f.get_nchannels(),
            f.get_samplespersec(),
            f.get_validbitspersample(),
            f.get_bitspersample(),
            f.get_subformat()
        ),
        Err(e) => format!("<{e}>"),
    };
    for direction in [Direction::Render, Direction::Capture] {
        let collection = enumerator.get_device_collection(&direction).unwrap();
        println!(
            "\n{direction}: {} devices",
            collection.get_nbr_devices().unwrap()
        );
        for device in (&collection).into_iter().flatten() {
            println!("  {}", support::describe(&device));
            println!(
                "    device {} | mix {}",
                describe_format(device.get_device_format()),
                describe_format(device.get_iaudioclient().and_then(|c| c.get_mixformat()))
            );
            match device.get_data_ranges() {
                Ok(ranges) if ranges.is_empty() => println!("    ranges: none declared"),
                Ok(ranges) => println!("    ranges: {ranges:?}"),
                Err(e) => println!("    ranges: <{e}>"),
            }
        }
    }
}

/// Indexed access and the iterator have to see the same devices in the same order.
///
/// `DeviceCollectionIter` counts its own index, so an off by one here is a real
/// possibility. The collection also filters on `DEVICE_STATE_ACTIVE`, so every member
/// has to report that state.
#[test]
fn collection_agrees_with_iterator() {
    let Some(fx) = support::fixture() else { return };
    for direction in [Direction::Render, Direction::Capture] {
        let collection = fx.enumerator.get_device_collection(&direction).unwrap();
        let count = collection.get_nbr_devices().unwrap();
        assert_eq!(collection.get_direction(), direction);

        let by_index: Vec<String> = (0..count)
            .map(|n| collection.get_device_at_index(n).unwrap().get_id().unwrap())
            .collect();
        let by_iterator: Vec<String> = (&collection)
            .into_iter()
            .map(|device| device.unwrap().get_id().unwrap())
            .collect();
        assert_eq!(by_index, by_iterator, "{direction} devices");
        assert_eq!(by_index.len(), count as usize);

        for device in (&collection).into_iter().flatten() {
            assert_eq!(device.get_direction(), direction);
            assert_eq!(device.get_state().unwrap(), DeviceState::Active);
        }
        // One past the end is an error, not a panic.
        assert!(collection.get_device_at_index(count).is_err());
    }
}

/// An id and a name both have to find the same device again.
///
/// `get_device` works out the direction through `IMMEndpoint::GetDataFlow` rather than
/// inheriting it from the collection, which is the part worth checking.
#[test]
fn device_lookup_round_trips() {
    let Some(fx) = support::fixture() else { return };
    for device in [&fx.cable.render, &fx.cable.capture] {
        let id = device.get_id().unwrap();
        let name = device.get_friendlyname().unwrap();

        let by_id = fx.enumerator.get_device(&id).unwrap();
        assert_eq!(by_id.get_id().unwrap(), id);
        assert_eq!(by_id.get_friendlyname().unwrap(), name);
        assert_eq!(by_id.get_direction(), device.get_direction());

        let collection = fx
            .enumerator
            .get_device_collection(&device.get_direction())
            .unwrap();
        let by_name = collection.get_device_with_name(&name).unwrap();
        assert_eq!(by_name.get_id().unwrap(), id);
    }

    let collection = fx
        .enumerator
        .get_device_collection(&Direction::Render)
        .unwrap();
    assert!(matches!(
        collection.get_device_with_name("no such device"),
        Err(WasapiError::DeviceNotFound(_))
    ));
}

/// Registering for device notifications and dropping the registration is clean.
///
/// `DeviceEventRegistration` unregisters in its `Drop`, which is hand-written. The
/// dispatch of each callback is already covered by the unit tests in src/events.rs.
#[test]
fn notification_registration_round_trip() {
    let Some(fx) = support::fixture() else { return };
    let register = || {
        fx.enumerator
            .register_notification_callback(DeviceEventCallbacks::new())
            .unwrap()
    };
    let (first, second) = (register(), register());
    drop(first);
    drop(second);
    drop(register());
}
