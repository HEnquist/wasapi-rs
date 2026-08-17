//! Probing of the formats a device supports in exclusive mode.

use std::collections::{BTreeSet, HashMap};

use crate::{AudioClient, DataRange, Device, SampleType, WasapiRes, WaveFormat, covered_by_any};
use windows::Win32::Media::KernelStreaming::WAVE_FORMAT_EXTENSIBLE;

/// The channel count ceiling that a scan uses for a device
/// that declares no [DataRange]s of its own.
pub const DEFAULT_MAX_CHANNELS: usize = 32;

// Standard rates in each family, from the base rate upward through the multiples.
const FAMILY_48_RATES: &[usize] = &[48000, 96000, 192000, 384000, 768000];
const FAMILY_44_RATES: &[usize] = &[44100, 88200, 176400, 352800, 705600];

// Sub-multiples and the 32 kHz family, probed after the upward scan.
const REMAINING_RATES: &[usize] = &[
    24000, 12000, 6000, 22050, 11025, 5512, 16000, 8000, 32000, 64000,
];

/// Every rate of the three lists above, in ascending order.
/// Used when the declared capabilities of the driver make the staged scan unnecessary.
const ALL_RATES: &[usize] = &[
    5512, 6000, 8000, 11025, 12000, 16000, 22050, 24000, 32000, 44100, 48000, 64000, 88200, 96000,
    176400, 192000, 352800, 384000, 705600, 768000,
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
/// seconds per device, so the search space has to be cut down.
///
/// The probed sample formats are 16 bit integer, 24 bit integer packed in three bytes,
/// 24 bit integer padded in four bytes, 32 bit integer and 32 bit float.
/// The accepted channel mask of each channel count is cached and reused,
/// which avoids repeating the mask renegotiation of
/// [is_supported_exclusive_with_quirks](AudioClient::is_supported_exclusive_with_quirks).
///
/// The channel counts run from one up to a ceiling.
/// With [DataRange]s that ceiling is the largest channel count the driver declares,
/// which is exact. Without them it is [DEFAULT_MAX_CHANNELS], a guess that is
/// deliberately generous, since a channel count above it would go unnoticed.
/// The staged scan below then lowers it to the highest count that worked,
/// as soon as any rate succeeds.
///
/// ## With the ranges the driver declares
///
/// When the probe has [DataRange]s, from [CapabilityProbe::new] or
/// [CapabilityProbe::set_data_ranges], they give real bounds on the rates,
/// channel counts and sample formats.
/// The scan then only asks about the combinations that fall inside them,
/// and needs no guessing at all.
///
/// The ranges are declared per pin and over-report, so every combination
/// inside them is still confirmed with a query.
/// A driver that declares too little would make the scan miss something,
/// but on the devices this has been tried on, the declared ranges and the
/// staged scan below agree exactly, and the bounded scan is several times faster.
///
/// ## Without them
///
/// A probe without ranges, because the device declares none or because they were
/// cleared with [CapabilityProbe::set_data_ranges], falls back to a staged scan
/// that guesses instead.
/// Reading the ranges means walking the topology of the device,
/// and drivers build those in ways that are hard to cover in full,
/// so this is what keeps a device that cannot be walked from
/// looking like a device without any capabilities:
///
/// - The 48 kHz and 44.1 kHz families are probed interleaved from the base rate upward.
///   The first hit establishes an upper channel count limit,
///   and a reduced sample format set that all later probes reuse.
/// - Within each rate of the scan, the sample format candidates are narrowed
///   as soon as the first channel count succeeds with fewer than the full set.
/// - Each family gets an early cutoff. Once a family has a hit,
///   a miss at the next rate deactivates it, and the upward scan stops
///   when both families are inactive.
/// - The remaining low rates and the 32 kHz family are probed
///   using only the channel counts found during the upward scan.
///
/// These heuristics cut the probing time down to something reasonable on normal hardware,
/// but they are still heuristics.
/// A device that supports 48, 96 and 384 kHz but not 192 kHz loses the top rate to the
/// cutoff, and a format that only works at some other channel count can be narrowed away.
///
/// # Probing a device that is in use
///
/// The probing only queries, it never initializes a client or starts a stream,
/// so it does not disturb anything that is playing or recording.
/// A device that another process holds in exclusive mode can still be probed,
/// and gives the same answers as an idle one.
///
/// # Channel masks
///
/// Each returned format carries the first channel mask the device accepted for that channel count,
/// and that mask is then reused for the rest of the probing.
/// A device may well accept several masks for the same channel count,
/// for example both of the 5.1 layouts for six channels,
/// but the probing stops at the first one and the others are never tried.
///
/// To find every layout a device accepts, build the formats with
/// [WaveFormat::new] and a mask from [make_channelmasks](crate::make_channelmasks),
/// and query them one by one with [AudioClient::is_supported].
/// That function returns the masks that are worth trying for a channel count,
/// with the most likely one first, and a mask of your own is built from the
/// [SPEAKER_FRONT_LEFT](crate::SPEAKER_FRONT_LEFT) and friends constants.
///
/// ```no_run
/// use wasapi::{CapabilityProbe, Direction, DeviceEnumerator};
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let device = DeviceEnumerator::new()?.get_default_device(&Direction::Render)?;
/// let mut probe = CapabilityProbe::new(&device)?;
///
/// // Everything the device accepts at 48 kHz.
/// let formats = probe.supported_formats_at_rate(48000);
///
/// // Everything the device accepts, at any rate.
/// let all = probe.supported_formats_all_rates();
/// # Ok(())
/// # }
/// ```
pub struct CapabilityProbe {
    client: AudioClient,
    channel_masks: ChannelMaskMap,
    data_ranges: Vec<DataRange>,
}

impl CapabilityProbe {
    /// Create a new probe for a [Device].
    ///
    /// This gets an [AudioClient] of its own for the device, and reads the
    /// [DataRange]s that the driver declares, which are used to narrow down the search.
    /// A device that declares none, see [Device::get_data_ranges],
    /// gets the staged scan instead.
    ///
    /// The probing only queries the client it holds, and never initializes it,
    /// so it does not interfere with a client used for streaming.
    pub fn new(device: &Device) -> WasapiRes<Self> {
        let client = device.get_iaudioclient()?;
        let data_ranges = device.get_data_ranges().unwrap_or_else(|err| {
            debug!("Could not read the data ranges of the device, {err}");
            Vec::new()
        });
        Ok(CapabilityProbe {
            client,
            channel_masks: ChannelMaskMap::new(),
            data_ranges,
        })
    }

    /// Get the [DataRange]s the probe uses to narrow down the search.
    /// The list is empty when the probe has none, and then nothing is skipped.
    pub fn data_ranges(&self) -> &[DataRange] {
        &self.data_ranges
    }

    /// Set the [DataRange]s the probe uses to narrow down the search.
    /// An empty list turns the narrowing off,
    /// which is the way to ignore what the driver declares.
    pub fn set_data_ranges(&mut self, data_ranges: Vec<DataRange>) {
        self.data_ranges = data_ranges;
    }

    /// Get the formats the device accepts at the given sample rate and channel count.
    ///
    /// This is the cheapest probe, at most one query per sample format.
    pub fn supported_formats(&mut self, samplerate: usize, channels: usize) -> Vec<WaveFormat> {
        self.probing()
            .formats(samplerate, channels, CANDIDATE_FORMATS)
            .into_iter()
            .map(|(_, wave_fmt)| wave_fmt)
            .collect()
    }

    /// Get the formats the device accepts at the given sample rate,
    /// for every channel count the device can have.
    ///
    /// The channel counts go up to the ceiling described in the
    /// [struct documentation](CapabilityProbe).
    /// A single rate gives nothing to learn from, unlike the full scan,
    /// so a device that declares no [DataRange]s is probed all the way up to
    /// [DEFAULT_MAX_CHANNELS] here, and every sample format is tried for every
    /// channel count. None of the pruning of the full scan is used,
    /// so a format that only works at a single channel count is still found.
    /// Use [CapabilityProbe::supported_formats] instead
    /// when only one channel count is of interest.
    pub fn supported_formats_at_rate(&mut self, samplerate: usize) -> Vec<WaveFormat> {
        let mut probing = self.probing();
        let ceiling = probing.channel_ceiling();
        probing
            .rate(samplerate, 1..=ceiling, CANDIDATE_FORMATS, false)
            .formats
    }

    /// Get the formats the device accepts at any of the standard sample rates,
    /// for every channel count the device can have,
    /// see the [struct documentation](CapabilityProbe) for the ceiling that is used.
    ///
    /// This is the full scan, and the most expensive probe by far.
    /// Without any [DataRange]s it is also the one that leans hardest
    /// on the pruning heuristics, see the [struct documentation](CapabilityProbe).
    pub fn supported_formats_all_rates(&mut self) -> Vec<WaveFormat> {
        let mut probing = self.probing();
        let ceiling = probing.channel_ceiling();
        probing.all_rates(ceiling)
    }

    /// Borrow the client, the channel mask cache and the data ranges as a [Probing].
    fn probing(&mut self) -> Probing<'_, AudioClient> {
        Probing {
            checker: &self.client,
            channel_masks: &mut self.channel_masks,
            data_ranges: &self.data_ranges,
        }
    }
}

/// The state that the probing logic works on, borrowed from a [CapabilityProbe].
///
/// The logic lives here instead of directly on [CapabilityProbe] so that it can be
/// generic over [FormatChecker]. That is what lets the unit tests run the real
/// pruning logic against a fake device, since the real one needs hardware.
struct Probing<'a, C: FormatChecker> {
    checker: &'a C,
    channel_masks: &'a mut ChannelMaskMap,
    data_ranges: &'a [DataRange],
}

impl<C: FormatChecker> Probing<'_, C> {
    /// Probe every candidate format at a single rate and channel count.
    /// Returns the accepted formats, paired with the candidate that produced them.
    fn formats(
        &mut self,
        samplerate: usize,
        channels: usize,
        candidates: &[Candidate],
    ) -> Vec<(Candidate, WaveFormat)> {
        let mut supported = Vec::new();
        if channels == 0 {
            return supported;
        }
        let mut preferred_mask = self.channel_masks.get(&channels).copied();
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
            if !covered_by_any(self.data_ranges, &requested) {
                trace!(
                    "Skipping {samplerate} Hz, {channels} ch, format {candidate:?}, the driver declares no range for it"
                );
                continue;
            }
            let Ok(accepted) = self.checker.check_exclusive(&requested) else {
                trace!("Unsupported {samplerate} Hz, {channels} ch, format {candidate:?}");
                continue;
            };
            trace!("Supported {samplerate} Hz, {channels} ch, format {candidate:?}");
            if accepted.wave_fmt.Format.wFormatTag == WAVE_FORMAT_EXTENSIBLE as u16 {
                let mask = accepted.get_dwchannelmask();
                if self.channel_masks.insert(channels, mask) != Some(mask) {
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

    /// Probe a single rate for the given channel counts.
    /// With `narrow` the candidate formats are cut down as soon as a channel count
    /// succeeds with fewer than all of them.
    fn rate<I>(
        &mut self,
        samplerate: usize,
        channel_counts: I,
        candidates: &[Candidate],
        narrow: bool,
    ) -> RateProbe
    where
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
            let supported = self.formats(samplerate, channels, active);
            if supported.is_empty() {
                trace!("No supported formats at {samplerate} Hz, {channels} ch");
                continue;
            }
            let found: Vec<Candidate> = supported.iter().map(|(candidate, _)| *candidate).collect();
            debug!("Found support at {samplerate} Hz, {channels} ch with formats {found:?}");
            if narrow && narrowed.is_none() && found.len() < candidates.len() {
                debug!(
                    "Narrowing the formats for the rest of the {samplerate} Hz sweep to {found:?}"
                );
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

    /// The highest channel count to probe.
    /// The declared ranges give a real bound, and without them
    /// there is nothing better than a generous guess.
    fn channel_ceiling(&self) -> usize {
        self.data_ranges
            .iter()
            .map(|range| range.max_channels as usize)
            .max()
            .unwrap_or(DEFAULT_MAX_CHANNELS)
    }

    /// Probe all the standard rates, up to the given channel count.
    fn all_rates(&mut self, ceiling: usize) -> Vec<WaveFormat> {
        if !self.data_ranges.is_empty() {
            return self.all_rates_within_ranges(ceiling);
        }
        self.all_rates_staged(ceiling)
    }

    /// Probe the rates and channel counts that the driver declares support for.
    /// The declared ranges are real bounds, so none of the guessing of the staged scan is needed.
    fn all_rates_within_ranges(&mut self, max_channels: usize) -> Vec<WaveFormat> {
        let declared_channels = self
            .data_ranges
            .iter()
            .map(|range| range.max_channels as usize)
            .max()
            .unwrap_or(0);
        let ceiling = declared_channels.min(max_channels);
        debug!(
            "Starting exclusive mode scan with channel ceiling {ceiling}, \
             the driver declares at most {declared_channels} channels"
        );
        let mut formats = Vec::new();
        for &rate in ALL_RATES {
            let declared = self.data_ranges.iter().any(|range| {
                (range.min_samplerate..=range.max_samplerate).contains(&(rate as u32))
            });
            if !declared {
                trace!("Skipping {rate} Hz, the driver declares no range for it");
                continue;
            }
            let result = self.rate(rate, 1..=ceiling, CANDIDATE_FORMATS, false);
            formats.extend(result.formats);
        }
        debug!("The exclusive mode scan found {} formats", formats.len());
        formats
    }

    /// Probe all the standard rates in stages, pruning the search as it goes.
    /// This is what is left when the driver declares nothing.
    fn all_rates_staged(&mut self, max_channels: usize) -> Vec<WaveFormat> {
        debug!("Starting staged exclusive mode scan with channel ceiling {max_channels}");
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
                let result = self.rate(rate, 1..=limit, candidates, true);
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
            debug!(
                "Probing the remaining rates with the full channel range, nothing was found so far"
            );
            (1..=max_channels).collect()
        } else {
            debug!("Probing the remaining rates with the channel counts {channel_counts:?}");
            channel_counts.iter().copied().collect()
        };
        for &rate in REMAINING_RATES {
            let candidates = learned.as_deref().unwrap_or(CANDIDATE_FORMATS);
            let result = self.rate(rate, remaining_counts.iter().copied(), candidates, true);
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{WasapiError, make_channelmasks};
    use std::cell::RefCell;

    /// Build a probing state for a fake device, without any declared ranges.
    fn probing<'a>(
        device: &'a FakeDevice,
        channel_masks: &'a mut ChannelMaskMap,
    ) -> Probing<'a, FakeDevice> {
        Probing {
            checker: device,
            channel_masks,
            data_ranges: &[],
        }
    }

    /// Build a probing state for a fake device with declared ranges.
    fn probing_with<'a>(
        device: &'a FakeDevice,
        channel_masks: &'a mut ChannelMaskMap,
        data_ranges: &'a [DataRange],
    ) -> Probing<'a, FakeDevice> {
        Probing {
            checker: device,
            channel_masks,
            data_ranges,
        }
    }

    /// A range of PCM formats, as a driver would declare it.
    fn declared(max_channels: u32, bits: (u32, u32), rates: (u32, u32)) -> DataRange {
        DataRange {
            max_channels,
            min_bits_per_sample: bits.0,
            max_bits_per_sample: bits.1,
            min_samplerate: rates.0,
            max_samplerate: rates.1,
            subformat: windows::Win32::Media::KernelStreaming::KSDATAFORMAT_SUBTYPE_PCM,
        }
    }

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
        mono_only_formats: Vec<Candidate>,
        queries: RefCell<Vec<Query>>,
    }

    impl FakeDevice {
        fn new(rates: &[usize], max_channels: usize, formats: &[Candidate]) -> Self {
            FakeDevice {
                rates: rates.to_vec(),
                max_channels,
                skipped_channels: Vec::new(),
                formats: formats.to_vec(),
                mono_only_formats: Vec::new(),
                queries: RefCell::new(Vec::new()),
            }
        }

        /// Make some formats work with a single channel only.
        fn only_with_one_channel(mut self, formats: &[Candidate]) -> Self {
            self.mono_only_formats = formats.to_vec();
            self
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
                || (channels > 1 && self.mono_only_formats.contains(&candidate))
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
        let supported = probing(&device, &mut masks).formats(48000, 2, CANDIDATE_FORMATS);

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
        assert!(
            probing(&device, &mut masks)
                .formats(44100, 2, CANDIDATE_FORMATS)
                .is_empty()
        );
        assert!(
            probing(&device, &mut masks)
                .formats(48000, 4, CANDIDATE_FORMATS)
                .is_empty()
        );
        assert!(
            probing(&device, &mut masks)
                .formats(48000, 0, CANDIDATE_FORMATS)
                .is_empty()
        );
    }

    #[test]
    fn the_accepted_channel_mask_is_cached_and_reused() {
        let device = FakeDevice::new(&[48000, 96000], 2, &[S16, S32]);
        let mut masks = ChannelMaskMap::new();
        let accepted = FakeDevice::accepted_mask(2);

        probing(&device, &mut masks).formats(48000, 2, CANDIDATE_FORMATS);
        assert_eq!(masks.get(&2), Some(&accepted));
        // The first query of the first probe still uses the default mask.
        assert_ne!(device.queries_for(48000, 2)[0].mask, accepted);

        probing(&device, &mut masks).formats(96000, 2, CANDIDATE_FORMATS);
        // The cached mask is used from the very first query of the second probe.
        assert!(
            device
                .queries_for(96000, 2)
                .iter()
                .all(|q| q.mask == accepted)
        );
    }

    #[test]
    fn the_formats_are_narrowed_after_the_first_channel_count() {
        let device = FakeDevice::new(&[48000], 4, &[S32]);
        let mut masks = ChannelMaskMap::new();
        let result = probing(&device, &mut masks).rate(48000, 1..=4, CANDIDATE_FORMATS, true);

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
    fn without_narrowing_all_formats_are_probed_at_every_channel_count() {
        // S32 only works with one channel, which narrowing would drop after the first count.
        let device = FakeDevice::new(&[48000], 4, &[S16, S32]).only_with_one_channel(&[S32]);
        let mut masks = ChannelMaskMap::new();
        let result = probing(&device, &mut masks).rate(48000, 1..=4, CANDIDATE_FORMATS, false);

        assert_eq!(result.supported_candidates, vec![S16, S32]);
        for channels in 1..=4 {
            assert_eq!(
                device.queries_for(48000, channels).len(),
                CANDIDATE_FORMATS.len()
            );
        }
    }

    #[test]
    fn a_rate_probe_covers_all_channel_counts() {
        // A device with a gap, it takes two and four channels but not three.
        let device = FakeDevice::new(&[48000], 4, &[S16]).without_channels(&[3]);
        let mut masks = ChannelMaskMap::new();
        let mut result =
            probing(&device, &mut masks).rate(48000, [2, 3, 4], CANDIDATE_FORMATS, true);
        result.formats.retain(|fmt| fmt.get_nchannels() == 4);
        assert_eq!(result.channel_counts, BTreeSet::from([2, 4]));
        assert_eq!(result.formats.len(), 1);
    }

    #[test]
    fn the_full_scan_finds_all_the_supported_combinations() {
        let device = FakeDevice::new(&[44100, 48000, 96000, 32000], 2, &[S16, S24_3]);
        let mut masks = ChannelMaskMap::new();
        let mut found: Vec<(u32, u16, u16, u16)> = probing(&device, &mut masks)
            .all_rates(8)
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
        let found: Vec<u32> = probing(&device, &mut masks)
            .all_rates(8)
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
        probing(&device, &mut masks).all_rates(8);

        // The first rate probes the full range, the ceiling drops to two after that.
        assert!(device.queries().iter().any(|q| q.channels == 8));
        assert!(
            !device
                .queries()
                .iter()
                .any(|q| q.samplerate == 44100 && q.channels > 2)
        );
        // The low rates only use the channel counts that were found.
        assert!(
            !device
                .queries()
                .iter()
                .any(|q| q.samplerate == 32000 && q.channels > 2)
        );
    }

    #[test]
    fn the_declared_ranges_keep_the_scan_inside_them() {
        let device = FakeDevice::new(&[44100, 48000], 2, &[S16, S32]);
        let mut masks = ChannelMaskMap::new();
        // The driver only declares two channels, 16 bit, and the two rates.
        let ranges = [declared(2, (16, 16), (44100, 48000))];
        let found: Vec<(u32, u16, u16, u16)> = probing_with(&device, &mut masks, &ranges)
            .all_rates(DEFAULT_MAX_CHANNELS)
            .iter()
            .map(describe)
            .collect();

        assert_eq!(
            found,
            vec![
                (44100, 1, 16, 16),
                (44100, 2, 16, 16),
                (48000, 1, 16, 16),
                (48000, 2, 16, 16)
            ]
        );
        // Nothing outside the declared ranges is even asked about.
        assert!(
            device
                .queries()
                .iter()
                .all(|q| q.channels <= 2 && q.candidate == S16)
        );
        assert!(
            !device
                .queries()
                .iter()
                .any(|q| q.samplerate < 44100 || q.samplerate > 48000)
        );
    }

    #[test]
    fn the_declared_ranges_find_a_rate_that_the_staged_scan_misses() {
        // A device with a hole at 192 kHz, which cuts the staged scan short.
        let device = FakeDevice::new(&[48000, 96000, 384000], 2, &[S32]);
        let mut masks = ChannelMaskMap::new();
        let staged: Vec<u32> = probing(&device, &mut masks)
            .all_rates(8)
            .iter()
            .map(|fmt| fmt.get_samplespersec())
            .collect();
        assert!(!staged.contains(&384000));

        let device = FakeDevice::new(&[48000, 96000, 384000], 2, &[S32]);
        let mut masks = ChannelMaskMap::new();
        let ranges = [declared(2, (16, 32), (48000, 384000))];
        let bounded: Vec<u32> = probing_with(&device, &mut masks, &ranges)
            .all_rates(8)
            .iter()
            .map(|fmt| fmt.get_samplespersec())
            .collect();
        assert!(bounded.contains(&384000));
    }

    #[test]
    fn the_declared_ranges_keep_all_the_formats_of_every_channel_count() {
        // S32 only works with one channel, which the staged scan would narrow away.
        let device = FakeDevice::new(&[48000], 4, &[S16, S32]).only_with_one_channel(&[S32]);
        let mut masks = ChannelMaskMap::new();
        let ranges = [declared(4, (16, 32), (48000, 48000))];
        let found = probing_with(&device, &mut masks, &ranges).all_rates(4);
        assert_eq!(describe(&found[0]), (48000, 1, 16, 16));
        assert_eq!(describe(&found[1]), (48000, 1, 32, 32));
        // Every channel count is probed with all four integer candidates that
        // fit in the declared 16 to 32 bits, the float one is left out.
        for channels in 1..=4 {
            let queries = device.queries_for(48000, channels);
            assert_eq!(queries.len(), 4);
            assert!(
                queries
                    .iter()
                    .all(|q| q.candidate.sample_type == SampleType::Int)
            );
        }
    }

    #[test]
    fn every_rate_is_in_the_combined_list() {
        let mut combined: Vec<usize> = FAMILY_48_RATES
            .iter()
            .chain(FAMILY_44_RATES)
            .chain(REMAINING_RATES)
            .copied()
            .collect();
        combined.sort_unstable();
        assert_eq!(combined, ALL_RATES);
    }

    #[test]
    fn the_full_scan_of_a_device_without_support_finds_nothing() {
        let device = FakeDevice::new(&[], 0, &[]);
        let mut masks = ChannelMaskMap::new();
        assert!(probing(&device, &mut masks).all_rates(2).is_empty());
        assert!(masks.is_empty());
        // Nothing was found, so the low rates are probed with the full channel range.
        assert!(
            device
                .queries()
                .iter()
                .any(|q| q.samplerate == 32000 && q.channels == 2)
        );
    }
}
