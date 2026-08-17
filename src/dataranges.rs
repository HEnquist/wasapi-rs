//! Reading the capabilities that a driver declares for a device.
//!
//! A Wasapi endpoint is backed by a pin on a kernel streaming filter,
//! and a WDM audio driver declares what that pin accepts as a set of data ranges.
//! Getting at them means walking the device topology from the endpoint,
//! through the topology filter of the device, to the wave filter that does the streaming,
//! and then querying that filter directly.

use std::collections::HashSet;
use std::mem::size_of;
use std::ptr::from_ref;

use windows::core::{Interface, GUID, HRESULT, PCWSTR};
use windows::Win32::Foundation::{
    CloseHandle, ERROR_INSUFFICIENT_BUFFER, ERROR_MORE_DATA, GENERIC_READ, GENERIC_WRITE, HANDLE,
};
use windows::Win32::Media::Audio::{Connector, IConnector, IDeviceTopology, IMMDevice, IPart};
use windows::Win32::Media::KernelStreaming::{
    KSPROPSETID_Pin, IOCTL_KS_PROPERTY, KSDATAFORMAT_0, KSDATAFORMAT_SUBTYPE_PCM,
    KSDATAFORMAT_TYPE_AUDIO, KSIDENTIFIER_0_0, KSMULTIPLE_ITEM, KSPIN_DATAFLOW_IN,
    KSPIN_DATAFLOW_OUT, KSPROPERTY_PIN, KSPROPERTY_PIN_CTYPES, KSPROPERTY_PIN_DATAFLOW,
    KSPROPERTY_PIN_DATARANGES, KSPROPERTY_TYPE_GET, KSP_PIN,
};
use windows::Win32::Media::Multimedia::KSDATAFORMAT_SUBTYPE_IEEE_FLOAT;
use windows::Win32::Storage::FileSystem::{
    CreateFileW, FILE_FLAGS_AND_ATTRIBUTES, FILE_SHARE_READ, FILE_SHARE_WRITE, OPEN_EXISTING,
};
use windows::Win32::System::Com::CLSCTX_ALL;
use windows::Win32::System::IO::DeviceIoControl;

use crate::{Direction, SampleType, WasapiRes, WaveFormat};

/// A KSDATAFORMAT is 64 bytes, the audio fields of a KSDATARANGE_AUDIO follow after it.
const AUDIO_RANGE_SIZE: usize = size_of::<KSDATAFORMAT_0>() + 5 * size_of::<u32>();

/// The errors that mean the reply did not fit in the buffer.
const MORE_DATA: HRESULT = HRESULT::from_win32(ERROR_MORE_DATA.0);
const INSUFFICIENT_BUFFER: HRESULT = HRESULT::from_win32(ERROR_INSUFFICIENT_BUFFER.0);

/// One capability range, as declared by a device driver.
///
/// A range is a cross product and over-reports.
/// A device declaring two to eight channels at 44.1 to 192 kHz
/// is not promising that every combination in that box works,
/// so a range is an upper bound that still has to be confirmed with
/// [is_supported](crate::AudioClient::is_supported).
/// A driver may also declare several ranges, one per format or rate.
///
/// This mirrors a
/// [KSDATARANGE_AUDIO](https://learn.microsoft.com/en-us/windows-hardware/drivers/ddi/ksmedia/ns-ksmedia-ksdatarange_audio).
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DataRange {
    /// The largest number of channels.
    pub max_channels: u32,
    /// The smallest container size in bits.
    pub min_bits_per_sample: u32,
    /// The largest container size in bits.
    pub max_bits_per_sample: u32,
    /// The lowest sample rate.
    pub min_samplerate: u32,
    /// The highest sample rate.
    pub max_samplerate: u32,
    /// The subformat, normally PCM or IEEE float.
    /// An all zero GUID is a wildcard, see [DataRange::sample_type].
    pub subformat: GUID,
}

impl DataRange {
    /// Get the sample type of the range, if it has one.
    ///
    /// This returns `None` in two cases.
    /// The first is a wildcard, an all zero GUID,
    /// `KSDATAFORMAT_SUBTYPE_WILDCARD` in ksmedia.h.
    /// A driver declares a wildcard for a property it does not want to restrict,
    /// so a wildcard subformat means the pin takes any of them.
    /// The second is a subformat that is neither PCM nor float,
    /// a compressed one for instance, which has no [SampleType] to map to.
    pub fn sample_type(&self) -> Option<SampleType> {
        match self.subformat {
            KSDATAFORMAT_SUBTYPE_PCM => Some(SampleType::Int),
            KSDATAFORMAT_SUBTYPE_IEEE_FLOAT => Some(SampleType::Float),
            _ => None,
        }
    }

    /// Check if a format falls inside this range.
    ///
    /// A range without a sample type of its own, see [DataRange::sample_type],
    /// is matched on the other properties alone.
    /// A range that cannot be interpreted then never excludes a format,
    /// which is the safe direction to err in,
    /// since the cost is a query that comes back negative.
    pub fn covers(&self, wave_fmt: &WaveFormat) -> bool {
        let samplerate = wave_fmt.get_samplespersec();
        let storebits = wave_fmt.get_bitspersample() as u32;
        let matching_type = match (self.sample_type(), wave_fmt.get_subformat()) {
            (Some(declared), Ok(wanted)) => declared == wanted,
            _ => true,
        };
        matching_type
            && wave_fmt.get_nchannels() as u32 <= self.max_channels
            && (self.min_bits_per_sample..=self.max_bits_per_sample).contains(&storebits)
            && (self.min_samplerate..=self.max_samplerate).contains(&samplerate)
    }
}

/// Check if a format falls inside any of the ranges.
/// An empty set of ranges means nothing is known, and everything is then accepted.
pub(crate) fn covered_by_any(ranges: &[DataRange], wave_fmt: &WaveFormat) -> bool {
    ranges.is_empty() || ranges.iter().any(|range| range.covers(wave_fmt))
}

/// Read the data ranges that the driver declares for a device.
///
/// This needs a device with a kernel streaming filter behind it.
/// A device without one has nothing to ask, and gets an empty list.
pub(crate) fn read_data_ranges(
    device: &IMMDevice,
    direction: Direction,
) -> WasapiRes<Vec<DataRange>> {
    let topology: IDeviceTopology = unsafe { device.Activate(CLSCTX_ALL, None)? };
    let endpoint_side = unsafe { topology.GetConnector(0)? };
    let device_side: IPart = unsafe { endpoint_side.GetConnectedTo()? }.cast()?;
    let topology_filter = filter_id(&device_side);

    // The endpoint connects to the topology filter of the device, which holds
    // the volume and mute controls. The wave filter that does the streaming is
    // on the other side of it, upstream for a render device and downstream for a capture device.
    let upstream = matches!(direction, Direction::Render);
    let mut wave_filters: Vec<(String, HANDLE)> = Vec::new();
    // Some drivers put the endpoint on a filter of their own, and then the
    // wave filter is in the other direction. Try both before giving up.
    for upstream in [upstream, !upstream] {
        let connectors = reachable_connectors(&device_side, upstream);
        debug!(
            "Found {} connectors {} of the endpoint",
            connectors.len(),
            if upstream { "upstream" } else { "downstream" }
        );
        for connector in connectors {
            let Ok(remote) = (unsafe { connector.GetConnectedTo() }) else {
                continue;
            };
            let Ok(remote) = remote.cast::<IPart>() else {
                continue;
            };
            let Some(filter) = filter_id(&remote) else {
                continue;
            };
            if Some(&filter) == topology_filter.as_ref()
                || wave_filters.iter().any(|(id, _)| *id == filter)
            {
                continue;
            }
            match open_filter(&filter) {
                Ok(handle) => wave_filters.push((filter, handle)),
                Err(err) => debug!("Could not open the filter {filter}, {err}"),
            }
        }
        if !wave_filters.is_empty() {
            break;
        }
    }
    // Some drivers have no separate wave filter, and then the streaming pins
    // are on the same filter as the endpoint connects to.
    if wave_filters.is_empty() {
        if let Some(filter) = topology_filter {
            debug!("Found no wave filter, trying the filter of the endpoint itself");
            match open_filter(&filter) {
                Ok(handle) => wave_filters.push((filter, handle)),
                Err(err) => debug!("Could not open the filter {filter}, {err}"),
            }
        }
    }

    // Take the pins that stream in the direction of the device,
    // data goes into a render filter and out of a capture filter.
    let wanted_flow = if upstream {
        KSPIN_DATAFLOW_IN
    } else {
        KSPIN_DATAFLOW_OUT
    };
    let mut ranges = Vec::new();
    for (id, filter) in &wave_filters {
        let pins = query_u32(*filter, &pin_property(KSPROPERTY_PIN_CTYPES, 0)).unwrap_or(0);
        debug!("The filter {id} has {pins} pins");
        for pin in 0..pins {
            let flow = query_u32(*filter, &pin_property(KSPROPERTY_PIN_DATAFLOW, pin));
            if flow.unwrap_or(0) != wanted_flow.0 as u32 {
                continue;
            }
            match query_bytes(*filter, &pin_property(KSPROPERTY_PIN_DATARANGES, pin)) {
                Ok(reply) => {
                    for range in parse_data_ranges(&reply) {
                        if !ranges.contains(&range) {
                            ranges.push(range);
                        }
                    }
                }
                Err(err) => debug!("Could not read the data ranges of pin {pin}, {err}"),
            }
        }
    }
    for (_, filter) in wave_filters {
        let _ = unsafe { CloseHandle(filter) };
    }
    debug!("The driver declares {} data ranges", ranges.len());
    Ok(ranges)
}

/// Get the device id of the filter that a part belongs to.
fn filter_id(part: &IPart) -> Option<String> {
    let topology: IDeviceTopology = unsafe { part.GetTopologyObject() }.ok()?;
    unsafe { topology.GetDeviceId().ok()?.to_string() }.ok()
}

/// Collect the connectors that can be reached from a part,
/// by walking through the subunits of the same filter.
fn reachable_connectors(start: &IPart, upstream: bool) -> Vec<IConnector> {
    let mut connectors = Vec::new();
    let mut seen = HashSet::new();
    let mut queue = vec![start.clone()];
    while let Some(part) = queue.pop() {
        let Ok(global_id) = (unsafe { part.GetGlobalId() }) else {
            continue;
        };
        if !seen.insert(unsafe { global_id.to_string() }.unwrap_or_default()) {
            continue;
        }
        let next = if upstream {
            unsafe { part.EnumPartsIncoming() }
        } else {
            unsafe { part.EnumPartsOutgoing() }
        };
        let Ok(next) = next else { continue };
        for index in 0..unsafe { next.GetCount() }.unwrap_or(0) {
            let Ok(part) = (unsafe { next.GetPart(index) }) else {
                continue;
            };
            if unsafe { part.GetPartType() } == Ok(Connector) {
                if let Ok(connector) = part.cast::<IConnector>() {
                    connectors.push(connector);
                }
            } else {
                queue.push(part);
            }
        }
    }
    connectors
}

/// Open a kernel streaming filter by its device interface path.
/// The device ids from the topology have a `{2}.` prefix that has to go.
fn open_filter(device_id: &str) -> WasapiRes<HANDLE> {
    let path = match device_id.find("}.") {
        Some(pos) if device_id.starts_with('{') => &device_id[pos + 2..],
        _ => device_id,
    };
    let wide: Vec<u16> = path.encode_utf16().chain(std::iter::once(0)).collect();
    let handle = unsafe {
        CreateFileW(
            PCWSTR(wide.as_ptr()),
            GENERIC_READ.0 | GENERIC_WRITE.0,
            FILE_SHARE_READ | FILE_SHARE_WRITE,
            None,
            OPEN_EXISTING,
            FILE_FLAGS_AND_ATTRIBUTES(0),
            None,
        )?
    };
    Ok(handle)
}

/// Build a pin property request.
fn pin_property(id: KSPROPERTY_PIN, pin_id: u32) -> KSP_PIN {
    let mut property = KSP_PIN::default();
    property.Property.Anonymous.Anonymous = KSIDENTIFIER_0_0 {
        Set: KSPROPSETID_Pin,
        Id: id.0 as u32,
        Flags: KSPROPERTY_TYPE_GET,
    };
    property.PinId = pin_id;
    property
}

/// Send a property request to an open filter.
/// Returns the number of bytes the driver has, or would have, written.
fn ks_property(filter: HANDLE, property: &KSP_PIN, buffer: Option<&mut [u8]>) -> WasapiRes<u32> {
    let (data, size) = match buffer {
        Some(buffer) => (Some(buffer.as_mut_ptr().cast()), buffer.len() as u32),
        None => (None, 0),
    };
    let mut returned = 0u32;
    let result = unsafe {
        DeviceIoControl(
            filter,
            IOCTL_KS_PROPERTY,
            Some(from_ref(property).cast()),
            size_of::<KSP_PIN>() as u32,
            data,
            size,
            Some(&mut returned),
            None,
        )
    };
    match result {
        Ok(()) => Ok(returned),
        // A buffer that is too small is not a failure here,
        // the driver then reports the size it needs.
        Err(err) if matches!(err.code(), MORE_DATA | INSUFFICIENT_BUFFER) => Ok(returned),
        Err(err) => Err(err.into()),
    }
}

/// Query a property that returns a single u32.
fn query_u32(filter: HANDLE, property: &KSP_PIN) -> WasapiRes<u32> {
    let mut buffer = [0u8; size_of::<u32>()];
    ks_property(filter, property, Some(&mut buffer))?;
    Ok(u32::from_le_bytes(buffer))
}

/// Query a property that returns a variable size reply.
/// The first call learns the size, the second one gets the data.
fn query_bytes(filter: HANDLE, property: &KSP_PIN) -> WasapiRes<Vec<u8>> {
    let needed = ks_property(filter, property, None)?;
    if needed as usize <= size_of::<KSMULTIPLE_ITEM>() {
        return Ok(Vec::new());
    }
    let mut buffer = vec![0u8; needed as usize];
    let returned = ks_property(filter, property, Some(&mut buffer))?;
    buffer.truncate(returned as usize);
    Ok(buffer)
}

/// Pick the audio data ranges out of a KSMULTIPLE_ITEM reply.
/// The reply is a header followed by a list of KSDATARANGE structures
/// of varying size, each padded to a multiple of eight bytes.
fn parse_data_ranges(buffer: &[u8]) -> Vec<DataRange> {
    let mut ranges = Vec::new();
    if buffer.len() < size_of::<KSMULTIPLE_ITEM>() {
        return ranges;
    }
    let header: KSMULTIPLE_ITEM = unsafe { std::ptr::read_unaligned(buffer.as_ptr().cast()) };
    let mut offset = size_of::<KSMULTIPLE_ITEM>();
    for _ in 0..header.Count {
        if offset + size_of::<KSDATAFORMAT_0>() > buffer.len() {
            debug!("The list of data ranges is truncated at offset {offset}");
            break;
        }
        let format: KSDATAFORMAT_0 =
            unsafe { std::ptr::read_unaligned(buffer[offset..].as_ptr().cast()) };
        let size = format.FormatSize as usize;
        if format.MajorFormat == KSDATAFORMAT_TYPE_AUDIO
            && size >= AUDIO_RANGE_SIZE
            && offset + AUDIO_RANGE_SIZE <= buffer.len()
        {
            let field = |nbr: usize| {
                let start = offset + size_of::<KSDATAFORMAT_0>() + 4 * nbr;
                u32::from_le_bytes(buffer[start..start + 4].try_into().unwrap())
            };
            ranges.push(DataRange {
                max_channels: field(0),
                min_bits_per_sample: field(1),
                max_bits_per_sample: field(2),
                min_samplerate: field(3),
                max_samplerate: field(4),
                subformat: format.SubFormat,
            });
        }
        if size == 0 {
            debug!("Got a data range of zero size, skipping the rest");
            break;
        }
        offset += size.div_ceil(8) * 8;
    }
    ranges
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Lay out a GUID the way it appears in a reply from the driver.
    fn guid_bytes(guid: &GUID) -> [u8; 16] {
        let mut bytes = [0u8; 16];
        bytes[0..4].copy_from_slice(&guid.data1.to_le_bytes());
        bytes[4..6].copy_from_slice(&guid.data2.to_le_bytes());
        bytes[6..8].copy_from_slice(&guid.data3.to_le_bytes());
        bytes[8..16].copy_from_slice(&guid.data4);
        bytes
    }

    fn range(channels: u32, bits: (u32, u32), rates: (u32, u32), subformat: GUID) -> DataRange {
        DataRange {
            max_channels: channels,
            min_bits_per_sample: bits.0,
            max_bits_per_sample: bits.1,
            min_samplerate: rates.0,
            max_samplerate: rates.1,
            subformat,
        }
    }

    #[test]
    fn a_range_covers_the_formats_inside_it() {
        let declared = range(8, (16, 24), (44100, 192000), KSDATAFORMAT_SUBTYPE_PCM);
        let inside = WaveFormat::new(24, 24, &SampleType::Int, 96000, 8, None);
        assert!(declared.covers(&inside));

        for outside in [
            WaveFormat::new(24, 24, &SampleType::Int, 96000, 9, None),
            WaveFormat::new(32, 32, &SampleType::Int, 96000, 8, None),
            WaveFormat::new(16, 16, &SampleType::Int, 22050, 8, None),
            WaveFormat::new(16, 16, &SampleType::Int, 384000, 8, None),
            WaveFormat::new(24, 24, &SampleType::Float, 96000, 8, None),
        ] {
            assert!(!declared.covers(&outside), "{outside:?}");
        }
    }

    #[test]
    fn a_wildcard_range_ignores_the_sample_type() {
        let declared = range(2, (16, 16), (48000, 48000), GUID::zeroed());
        assert_eq!(declared.sample_type(), None);
        assert!(declared.covers(&WaveFormat::new(16, 16, &SampleType::Int, 48000, 2, None)));
        assert!(declared.covers(&WaveFormat::new(16, 16, &SampleType::Float, 48000, 2, None)));
    }

    #[test]
    fn a_format_is_covered_if_any_range_covers_it() {
        let declared = [
            range(2, (24, 24), (48000, 48000), KSDATAFORMAT_SUBTYPE_PCM),
            range(8, (16, 16), (44100, 48000), KSDATAFORMAT_SUBTYPE_PCM),
        ];
        // Each format is outside one of the ranges but inside the other.
        let packed_24 = WaveFormat::new(24, 24, &SampleType::Int, 48000, 2, None);
        let eight_channels = WaveFormat::new(16, 16, &SampleType::Int, 44100, 8, None);
        assert!(covered_by_any(&declared, &packed_24));
        assert!(covered_by_any(&declared, &eight_channels));
        // A combination that no single range covers.
        let both = WaveFormat::new(24, 24, &SampleType::Int, 44100, 8, None);
        assert!(!covered_by_any(&declared, &both));
        // Without any ranges nothing is known, so everything passes.
        assert!(covered_by_any(&[], &both));
    }

    #[test]
    fn a_reply_with_ranges_is_parsed() {
        // Two ranges, an audio one and something else that must be skipped.
        let mut reply = Vec::new();
        reply.extend_from_slice(&(8u32 + 88 + 72).to_le_bytes()); // Size
        reply.extend_from_slice(&2u32.to_le_bytes()); // Count

        let mut audio = Vec::new();
        audio.extend_from_slice(&(AUDIO_RANGE_SIZE as u32).to_le_bytes()); // FormatSize
        audio.extend_from_slice(&[0u8; 12]); // Flags, SampleSize, Reserved
        audio.extend_from_slice(&guid_bytes(&KSDATAFORMAT_TYPE_AUDIO));
        audio.extend_from_slice(&guid_bytes(&KSDATAFORMAT_SUBTYPE_PCM));
        audio.extend_from_slice(&[0u8; 16]); // Specifier
        for value in [6u32, 16, 32, 44100, 192000] {
            audio.extend_from_slice(&value.to_le_bytes());
        }
        audio.resize(audio.len().div_ceil(8) * 8, 0);
        reply.extend_from_slice(&audio);

        let mut other = vec![0u8; 72];
        other[0..4].copy_from_slice(&72u32.to_le_bytes());
        reply.extend_from_slice(&other);

        let parsed = parse_data_ranges(&reply);
        assert_eq!(parsed.len(), 1);
        assert_eq!(
            parsed[0],
            range(6, (16, 32), (44100, 192000), KSDATAFORMAT_SUBTYPE_PCM)
        );
    }

    #[test]
    fn a_short_or_broken_reply_is_handled() {
        assert!(parse_data_ranges(&[]).is_empty());
        assert!(parse_data_ranges(&[0u8; 4]).is_empty());
        // A count of one but no range following it.
        let mut truncated = 8u32.to_le_bytes().to_vec();
        truncated.extend_from_slice(&1u32.to_le_bytes());
        assert!(parse_data_ranges(&truncated).is_empty());
        // A range of zero size must not loop forever.
        let mut zero_size = 200u32.to_le_bytes().to_vec();
        zero_size.extend_from_slice(&50u32.to_le_bytes());
        zero_size.extend_from_slice(&[0u8; 200]);
        assert!(parse_data_ranges(&zero_size).is_empty());
    }
}
