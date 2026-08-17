//! Probing of the formats a device supports in exclusive mode.

use std::collections::{BTreeSet, HashMap};

use crate::{AudioClient, SampleType, WasapiRes, WaveFormat};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;

/// The channel count ceiling used when nothing better is known.
pub const DEFAULT_MAX_CHANNELS: usize = 32;

// Standard rates in each family, from the base rate upward through the multiples.
const FAMILY_48_RATES: &[usize] = &[48000, 96000, 192000, 384000, 768000];
const FAMILY_44_RATES: &[usize] = &[44100, 88200, 176400, 352800, 705600];

// Sub-multiples and the 32 kHz family, probed after the upward scan.
const REMAINING_RATES: &[usize] = &[
    24000, 12000, 6000, 22050, 11025, 5512, 16000, 8000, 32000, 64000,
];

/// A sample format to probe for, as stored bits, valid bits and sample type.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
struct Candidate {
    storebits: usize,
    validbits: usize,
    sample_type: SampleType,
}

impl Candidate {
    const fn new(storebits: usize, validbits: usize, sample_type: SampleType) -> Self {
        Candidate {
            storebits,
            validbits,
            sample_type,
        }
    }
}

/// The sample formats that get probed, in the order they are tried.
/// Both 24 bit layouts are probed, packed in three bytes and padded in four.
const CANDIDATE_FORMATS: &[Candidate] = &[
    Candidate::new(16, 16, SampleType::Int),
    Candidate::new(24, 24, SampleType::Int),
    Candidate::new(32, 24, SampleType::Int),
    Candidate::new(32, 32, SampleType::Int),
    Candidate::new(32, 32, SampleType::Float),
];

/// Accepted channel mask per channel count.
type ChannelMaskMap = HashMap<usize, u32>;

/// The device query the probing logic is built on.
/// Implemented for [AudioClient], and for fake devices in the unit tests.
trait FormatChecker {
    fn check_exclusive(&self, wave_fmt: &WaveFormat) -> WasapiRes<WaveFormat>;
}

impl FormatChecker for AudioClient {
    fn check_exclusive(&self, wave_fmt: &WaveFormat) -> WasapiRes<WaveFormat> {
        self.is_supported_exclusive_with_quirks(wave_fmt)
    }
}

/// Probes a device for the formats it supports in exclusive mode.
///
/// The probes are available at three levels of detail,
/// for a single rate and channel count, for a single rate,
/// and for all rates. The more limited ones are a lot faster,
/// so an application that already knows what it wants should not run a full scan.
///
/// The accepted channel masks are cached in the struct and shared between the calls,
/// so reusing the same instance for several probes is much faster than making a new one for each.
///
/// # How the probing works
///
/// Wasapi has no structured way of asking a device what it supports.
/// The only option is to call `IsFormatSupported` for every combination
/// of sample rate, channel count, sample format and channel mask.
/// Brute forcing the full matrix is thousands of calls and takes several
/// seconds per device, so the full scan prunes the search space:
///
/// - The 48 kHz and 44.1 kHz families are probed interleaved from the base rate upward.
///   The first hit establishes an upper channel count limit,
///   and a reduced sample format set that all later probes reuse.
/// - Within a single rate, the sample format candidates are narrowed
///   as soon as the first channel count succeeds with fewer than the full set.
/// - The accepted channel mask of each channel count is cached and reused,
///   which avoids repeating the mask renegotiation of
///   [is_supported_exclusive_with_quirks](AudioClient::is_supported_exclusive_with_quirks).
/// - Each family gets an early cutoff. Once a family has a hit,
///   a miss at the next rate deactivates it, and the upward scan stops
///   when both families are inactive.
/// - The remaining low rates and the 32 kHz family are probed
///   using only the channel counts found during the upward scan.
///
/// These heuristics cut the probing time down to something reasonable on normal hardware,
/// but they are still heuristics.
/// An unusual device may support combinations that fall outside the probed ones.
///
/// The probed sample formats are 16 bit integer, 24 bit integer packed in three bytes,
/// 24 bit integer padded in four bytes, 32 bit integer and 32 bit float.
///
/// ```no_run
/// use wasapi::{CapabilityProbe, Direction, DeviceEnumerator, DEFAULT_MAX_CHANNELS};
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let device = DeviceEnumerator::new()?.get_default_device(&Direction::Render)?;
/// let mut probe = CapabilityProbe::new(device.get_iaudioclient()?);
///
/// // Everything the device accepts at 48 kHz, for up to eight channels.
/// let formats = probe.supported_formats_at_rate(48000, 8);
///
/// // Everything the device accepts, at any rate.
/// let all = probe.supported_formats_all_rates(DEFAULT_MAX_CHANNELS);
/// # Ok(())
/// # }
/// ```
pub struct CapabilityProbe {
    client: AudioClient,
    channel_masks: ChannelMaskMap,
}

impl CapabilityProbe {
    /// Create a new probe for the device of the given [AudioClient].
    ///
    /// The client must not have been initialized,
    /// and it can not be used for streaming while the probing runs.
    pub fn new(client: AudioClient) -> Self {
        CapabilityProbe {
            client,
            channel_masks: ChannelMaskMap::new(),
        }
    }

    /// Get a reference to the [AudioClient] the probe was created with.
    pub fn client(&self) -> &AudioClient {
        &self.client
    }

    /// Get the formats the device accepts at the given sample rate and channel count.
    ///
    /// This is the cheapest probe, at most one query per sample format.
    pub fn supported_formats(&mut self, samplerate: usize, channels: usize) -> Vec<WaveFormat> {
        probe_formats(
            &self.client,
            &mut self.channel_masks,
            samplerate,
            channels,
            CANDIDATE_FORMATS,
        )
        .into_iter()
        .map(|(_, wave_fmt)| wave_fmt)
        .collect()
    }

    /// Get the formats the device accepts at the given sample rate,
    /// for every channel count from one up to and including `max_channels`.
    pub fn supported_formats_at_rate(
        &mut self,
        samplerate: usize,
        max_channels: usize,
    ) -> Vec<WaveFormat> {
        probe_rate(
            &self.client,
            &mut self.channel_masks,
            samplerate,
            1..=max_channels,
            CANDIDATE_FORMATS,
        )
        .formats
    }

    /// Get the formats the device accepts at any of the standard sample rates,
    /// for channel counts up to and including `max_channels`.
    ///
    /// This is the full scan. It is the most expensive probe by far,
    /// and the one that leans hardest on the pruning heuristics,
    /// see the [struct documentation](CapabilityProbe).
    /// Pass [DEFAULT_MAX_CHANNELS] unless the channel count is known to be lower.
    pub fn supported_formats_all_rates(&mut self, max_channels: usize) -> Vec<WaveFormat> {
        scan_all_rates(&self.client, &mut self.channel_masks, max_channels)
    }
}

/// Probe every candidate format at a single rate and channel count.
/// Returns the accepted formats, paired with the candidate that produced them.
fn probe_formats<C: FormatChecker>(
    checker: &C,
    channel_masks: &mut ChannelMaskMap,
    samplerate: usize,
    channels: usize,
    candidates: &[Candidate],
) -> Vec<(Candidate, WaveFormat)> {
    let mut supported = Vec::new();
    if channels == 0 {
        return supported;
    }
    let mut preferred_mask = channel_masks.get(&channels).copied();
    if let Some(mask) = preferred_mask {
        trace!("Probing {samplerate} Hz, {channels} ch using cached channel mask {mask:#010x}");
    }
    for candidate in candidates {
        let requested = WaveFormat::new(
            candidate.storebits,
            candidate.validbits,
            &candidate.sample_type,
            samplerate,
            channels,
            preferred_mask,
        );
        let Ok(accepted) = checker.check_exclusive(&requested) else {
            trace!("Unsupported {samplerate} Hz, {channels} ch, format {candidate:?}");
            continue;
        };
        trace!("Supported {samplerate} Hz, {channels} ch, format {candidate:?}");
        if accepted.wave_fmt.Format.wFormatTag == WAVE_FORMAT_EXTENSIBLE as u16 {
            let mask = accepted.get_dwchannelmask();
            if channel_masks.insert(channels, mask) != Some(mask) {
                debug!("Channel count {channels} will use channel mask {mask:#010x}");
            }
            preferred_mask = Some(mask);
            supported.push((*candidate, accepted));
        } else {
            // The device only accepted the format in the simpler WAVEFORMATEX representation.
            // That is a known driver quirk for one and two channel formats,
            // and the format to use for streaming is still the WAVEFORMATEXTENSIBLE one.
            trace!("Accepted as WAVEFORMATEX, reporting the WAVEFORMATEXTENSIBLE form");
            supported.push((*candidate, requested));
        }
    }
    supported
}

/// The outcome of probing a single sample rate.
struct RateProbe {
    /// All accepted formats.
    formats: Vec<WaveFormat>,
    /// The channel counts that had at least one accepted format.
    channel_counts: BTreeSet<usize>,
    /// The union of the candidates that were accepted at any channel count.
    supported_candidates: Vec<Candidate>,
}

/// Probe a single rate for the given channel counts.
/// The candidate formats are narrowed as soon as a channel count
/// succeeds with fewer than all of them.
fn probe_rate<C, I>(
    checker: &C,
    channel_masks: &mut ChannelMaskMap,
    samplerate: usize,
    channel_counts: I,
    candidates: &[Candidate],
) -> RateProbe
where
    C: FormatChecker,
    I: IntoIterator<Item = usize>,
{
    trace!("Probing {samplerate} Hz using sample formats {candidates:?}");
    let mut result = RateProbe {
        formats: Vec::new(),
        channel_counts: BTreeSet::new(),
        supported_candidates: Vec::new(),
    };
    let mut narrowed: Option<Vec<Candidate>> = None;
    for channels in channel_counts {
        let active = narrowed.as_deref().unwrap_or(candidates);
        let supported = probe_formats(checker, channel_masks, samplerate, channels, active);
        if supported.is_empty() {
            trace!("No supported formats at {samplerate} Hz, {channels} ch");
            continue;
        }
        let found: Vec<Candidate> = supported.iter().map(|(candidate, _)| *candidate).collect();
        debug!("Found support at {samplerate} Hz, {channels} ch with formats {found:?}");
        if narrowed.is_none() && found.len() < candidates.len() {
            debug!("Narrowing the formats for the rest of the {samplerate} Hz sweep to {found:?}");
            narrowed = Some(found.clone());
        }
        for candidate in found {
            if !result.supported_candidates.contains(&candidate) {
                result.supported_candidates.push(candidate);
            }
        }
        result.channel_counts.insert(channels);
        result
            .formats
            .extend(supported.into_iter().map(|(_, wave_fmt)| wave_fmt));
    }
    result
}

/// Probe all the standard rates, up to the given channel count.
fn scan_all_rates<C: FormatChecker>(
    checker: &C,
    channel_masks: &mut ChannelMaskMap,
    max_channels: usize,
) -> Vec<WaveFormat> {
    debug!("Starting exclusive mode scan with channel ceiling {max_channels}");
    let mut formats = Vec::new();
    let mut channel_counts = BTreeSet::new();
    let mut learned: Option<Vec<Candidate>> = None;

    // Probe the two main families interleaved from the base rate upward.
    // The first hit at any rate gives the channel limit for all the later probes.
    // A family that has had a hit is deactivated by the first miss after it.
    let families = [FAMILY_48_RATES, FAMILY_44_RATES];
    let mut channel_limit = 0;
    let mut hit = [false; 2];
    let mut active = [true; 2];
    for step in 0..FAMILY_48_RATES.len().max(FAMILY_44_RATES.len()) {
        if !active.iter().any(|is_active| *is_active) {
            debug!("Stopping the upward scan, both families are inactive");
            break;
        }
        for (family_nbr, family) in families.iter().enumerate() {
            if !active[family_nbr] {
                continue;
            }
            let Some(&rate) = family.get(step) else {
                continue;
            };
            let limit = if channel_limit > 0 {
                channel_limit
            } else {
                max_channels
            };
            let candidates = learned.as_deref().unwrap_or(CANDIDATE_FORMATS);
            let result = probe_rate(checker, channel_masks, rate, 1..=limit, candidates);
            if let Some(&highest) = result.channel_counts.iter().next_back() {
                hit[family_nbr] = true;
                channel_limit = channel_limit.max(highest);
                debug!(
                    "Rate {rate} Hz gave at most {highest} channels, limit is now {channel_limit}"
                );
                if learned.is_none() {
                    debug!(
                        "Learned the sample formats {:?} from {rate} Hz, reusing them",
                        result.supported_candidates
                    );
                    learned = Some(result.supported_candidates);
                }
            } else if hit[family_nbr] {
                active[family_nbr] = false;
                debug!("Stopping at {rate} Hz, this family had a miss after its earlier hits");
            }
            channel_counts.extend(&result.channel_counts);
            formats.extend(result.formats);
        }
    }

    // Probe the sub-multiples and the 32 kHz family.
    // Reuse the channel counts found above, or take the full range if nothing was found.
    let remaining_counts: Vec<usize> = if channel_counts.is_empty() {
        debug!("Probing the remaining rates with the full channel range, nothing was found so far");
        (1..=max_channels).collect()
    } else {
        debug!("Probing the remaining rates with the channel counts {channel_counts:?}");
        channel_counts.iter().copied().collect()
    };
    for &rate in REMAINING_RATES {
        let candidates = learned.as_deref().unwrap_or(CANDIDATE_FORMATS);
        let result = probe_rate(
            checker,
            channel_masks,
            rate,
            remaining_counts.iter().copied(),
            candidates,
        );
        if learned.is_none() && !result.supported_candidates.is_empty() {
            debug!(
                "Learned the sample formats {:?} from {rate} Hz, reusing them",
                result.supported_candidates
            );
            learned = Some(result.supported_candidates);
        }
        formats.extend(result.formats);
    }
    debug!("The exclusive mode scan found {} formats", formats.len());
    formats
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{make_channelmasks, WasapiError};
    use std::cell::RefCell;

    /// A single query made to the fake device.
    #[derive(Clone, Copy, Debug, Eq, PartialEq)]
    struct Query {
        samplerate: usize,
        channels: usize,
        mask: u32,
        candidate: Candidate,
    }

    /// A fake device that accepts a fixed set of rates, channel counts and formats.
    /// It only accepts a single channel mask per channel count,
    /// and renegotiates the mask like
    /// [is_supported_exclusive_with_quirks](AudioClient::is_supported_exclusive_with_quirks) does.
    struct FakeDevice {
        rates: Vec<usize>,
        max_channels: usize,
        skipped_channels: Vec<usize>,
        formats: Vec<Candidate>,
        queries: RefCell<Vec<Query>>,
    }

    impl FakeDevice {
        fn new(rates: &[usize], max_channels: usize, formats: &[Candidate]) -> Self {
            FakeDevice {
                rates: rates.to_vec(),
                max_channels,
                skipped_channels: Vec::new(),
                formats: formats.to_vec(),
                queries: RefCell::new(Vec::new()),
            }
        }

        /// Punch a hole in the supported channel counts.
        fn without_channels(mut self, channels: &[usize]) -> Self {
            self.skipped_channels = channels.to_vec();
            self
        }

        /// The only mask this fake accepts for a channel count.
        /// The last of the suggested masks, so that the mask always needs renegotiation.
        fn accepted_mask(channels: usize) -> u32 {
            *make_channelmasks(channels).last().unwrap()
        }

        fn queries(&self) -> Vec<Query> {
            self.queries.borrow().clone()
        }

        fn queries_for(&self, samplerate: usize, channels: usize) -> Vec<Query> {
            self.queries()
                .into_iter()
                .filter(|q| q.samplerate == samplerate && q.channels == channels)
                .collect()
        }
    }

    impl FormatChecker for FakeDevice {
        fn check_exclusive(&self, wave_fmt: &WaveFormat) -> WasapiRes<WaveFormat> {
            let channels = wave_fmt.get_nchannels() as usize;
            let candidate = Candidate::new(
                wave_fmt.get_bitspersample() as usize,
                wave_fmt.get_validbitspersample() as usize,
                wave_fmt.get_subformat()?,
            );
            let samplerate = wave_fmt.get_samplespersec() as usize;
            self.queries.borrow_mut().push(Query {
                samplerate,
                channels,
                mask: wave_fmt.get_dwchannelmask(),
                candidate,
            });
            if !self.rates.contains(&samplerate)
                || channels > self.max_channels
                || self.skipped_channels.contains(&channels)
                || !self.formats.contains(&candidate)
            {
                return Err(WasapiError::UnsupportedFormat);
            }
            let mut accepted = wave_fmt.clone();
            accepted.wave_fmt.dwChannelMask = Self::accepted_mask(channels);
            Ok(accepted)
        }
    }

    const S16: Candidate = Candidate::new(16, 16, SampleType::Int);
    const S24_3: Candidate = Candidate::new(24, 24, SampleType::Int);
    const S32: Candidate = Candidate::new(32, 32, SampleType::Int);

    /// Describe a format as rate, channels, stored bits and valid bits.
    fn describe(wave_fmt: &WaveFormat) -> (u32, u16, u16, u16) {
        (
            wave_fmt.get_samplespersec(),
            wave_fmt.get_nchannels(),
            wave_fmt.get_bitspersample(),
            wave_fmt.get_validbitspersample(),
        )
    }

    #[test]
    fn probe_returns_the_supported_formats() {
        let device = FakeDevice::new(&[48000], 2, &[S16, S32]);
        let mut masks = ChannelMaskMap::new();
        let supported = probe_formats(&device, &mut masks, 48000, 2, CANDIDATE_FORMATS);

        let found: Vec<Candidate> = supported.iter().map(|(c, _)| *c).collect();
        assert_eq!(found, vec![S16, S32]);
        assert_eq!(describe(&supported[0].1), (48000, 2, 16, 16));
        assert_eq!(describe(&supported[1].1), (48000, 2, 32, 32));
        // Every candidate is tried, also the ones that fail.
        assert_eq!(device.queries().len(), CANDIDATE_FORMATS.len());
    }

    #[test]
    fn probe_returns_nothing_for_unsupported_rates_and_channel_counts() {
        let device = FakeDevice::new(&[48000], 2, &[S16]);
        let mut masks = ChannelMaskMap::new();
        assert!(probe_formats(&device, &mut masks, 44100, 2, CANDIDATE_FORMATS).is_empty());
        assert!(probe_formats(&device, &mut masks, 48000, 4, CANDIDATE_FORMATS).is_empty());
        assert!(probe_formats(&device, &mut masks, 48000, 0, CANDIDATE_FORMATS).is_empty());
    }

    #[test]
    fn the_accepted_channel_mask_is_cached_and_reused() {
        let device = FakeDevice::new(&[48000, 96000], 2, &[S16, S32]);
        let mut masks = ChannelMaskMap::new();
        let accepted = FakeDevice::accepted_mask(2);

        probe_formats(&device, &mut masks, 48000, 2, CANDIDATE_FORMATS);
        assert_eq!(masks.get(&2), Some(&accepted));
        // The first query of the first probe still uses the default mask.
        assert_ne!(device.queries_for(48000, 2)[0].mask, accepted);

        probe_formats(&device, &mut masks, 96000, 2, CANDIDATE_FORMATS);
        // The cached mask is used from the very first query of the second probe.
        assert!(device
            .queries_for(96000, 2)
            .iter()
            .all(|q| q.mask == accepted));
    }

    #[test]
    fn the_formats_are_narrowed_after_the_first_channel_count() {
        let device = FakeDevice::new(&[48000], 4, &[S32]);
        let mut masks = ChannelMaskMap::new();
        let result = probe_rate(&device, &mut masks, 48000, 1..=4, CANDIDATE_FORMATS);

        assert_eq!(result.supported_candidates, vec![S32]);
        assert_eq!(result.channel_counts, BTreeSet::from([1, 2, 3, 4]));
        // The first channel count pays for all the candidates, the rest only probe S32.
        assert_eq!(device.queries_for(48000, 1).len(), CANDIDATE_FORMATS.len());
        for channels in 2..=4 {
            let queries = device.queries_for(48000, channels);
            assert_eq!(queries.len(), 1);
            assert_eq!(queries[0].candidate, S32);
        }
    }

    #[test]
    fn a_rate_probe_covers_all_channel_counts() {
        // A device with a gap, it takes two and four channels but not three.
        let device = FakeDevice::new(&[48000], 4, &[S16]).without_channels(&[3]);
        let mut masks = ChannelMaskMap::new();
        let mut result = probe_rate(&device, &mut masks, 48000, [2, 3, 4], CANDIDATE_FORMATS);
        result.formats.retain(|fmt| fmt.get_nchannels() == 4);
        assert_eq!(result.channel_counts, BTreeSet::from([2, 4]));
        assert_eq!(result.formats.len(), 1);
    }

    #[test]
    fn the_full_scan_finds_all_the_supported_combinations() {
        let device = FakeDevice::new(&[44100, 48000, 96000, 32000], 2, &[S16, S24_3]);
        let mut masks = ChannelMaskMap::new();
        let mut found: Vec<(u32, u16, u16, u16)> = scan_all_rates(&device, &mut masks, 8)
            .iter()
            .map(describe)
            .collect();
        found.sort_unstable();

        let mut expected = Vec::new();
        for rate in [32000, 44100, 48000, 96000] {
            for channels in [1, 2] {
                expected.push((rate, channels, 16, 16));
                expected.push((rate, channels, 24, 24));
            }
        }
        expected.sort_unstable();
        assert_eq!(found, expected);
    }

    #[test]
    fn the_full_scan_stops_a_family_after_a_miss() {
        // 192 kHz is missing, so the 48 kHz family is dropped before 384 kHz.
        let device = FakeDevice::new(&[48000, 96000, 384000], 2, &[S32]);
        let mut masks = ChannelMaskMap::new();
        let found: Vec<u32> = scan_all_rates(&device, &mut masks, 8)
            .iter()
            .map(|fmt| fmt.get_samplespersec())
            .collect();

        assert!(found.contains(&96000));
        assert!(!found.contains(&384000));
        assert!(!device.queries().iter().any(|q| q.samplerate == 384000));
        assert!(!device.queries().iter().any(|q| q.samplerate == 768000));
        // The 44.1 kHz family never had a hit, so it is probed to the end.
        assert!(device.queries().iter().any(|q| q.samplerate == 705600));
    }

    #[test]
    fn the_full_scan_limits_the_channel_counts_of_the_later_rates() {
        let device = FakeDevice::new(&[48000, 44100, 32000], 2, &[S16]);
        let mut masks = ChannelMaskMap::new();
        scan_all_rates(&device, &mut masks, 8);

        // The first rate probes the full range, the ceiling drops to two after that.
        assert!(device.queries().iter().any(|q| q.channels == 8));
        assert!(!device
            .queries()
            .iter()
            .any(|q| q.samplerate == 44100 && q.channels > 2));
        // The low rates only use the channel counts that were found.
        assert!(!device
            .queries()
            .iter()
            .any(|q| q.samplerate == 32000 && q.channels > 2));
    }

    #[test]
    fn the_full_scan_of_a_device_without_support_finds_nothing() {
        let device = FakeDevice::new(&[], 0, &[]);
        let mut masks = ChannelMaskMap::new();
        assert!(scan_all_rates(&device, &mut masks, 2).is_empty());
        assert!(masks.is_empty());
        // Nothing was found, so the low rates are probed with the full channel range.
        assert!(device
            .queries()
            .iter()
            .any(|q| q.samplerate == 32000 && q.channels == 2));
    }
}
