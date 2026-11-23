#[inline]
pub unsafe fn ActivateAudioInterfaceAsync<P0, P3>(
    deviceinterfacepath: P0,
    riid: *const IID,
    activationparams: Option<*const PROPVARIANT>,
    completionhandler: P3,
) -> windows_core::Result<IActivateAudioInterfaceAsyncOperation>
where
    P0: windows_core::Param<windows_core::PCWSTR>,
    P3: windows_core::Param<IActivateAudioInterfaceCompletionHandler>,
{
    windows_core::link!("mmdevapi.dll" "system" fn ActivateAudioInterfaceAsync(deviceinterfacepath : windows_core::PCWSTR, riid : *const IID, activationparams : *const PROPVARIANT, completionhandler : *mut core::ffi::c_void, activationoperation : *mut *mut core::ffi::c_void) -> windows_core::HRESULT);
    unsafe {
        let mut result__ = core::mem::zeroed();
        ActivateAudioInterfaceAsync(
            deviceinterfacepath.param().abi(),
            riid,
            activationparams.unwrap_or(core::mem::zeroed()) as _,
            completionhandler.param().abi(),
            &mut result__,
        )
        .and_then(|| windows_core::imp::from_abi(result__))
    }
}
#[inline]
pub unsafe fn CloseHandle(hobject: HANDLE) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn CloseHandle(hobject : HANDLE) -> windows_core::BOOL);
    unsafe { CloseHandle(hobject) }
}
#[inline]
pub unsafe fn CoCreateInstance<P1, T>(
    rclsid: *const IID,
    punkouter: P1,
    dwclscontext: DWORD,
) -> windows_core::Result<T>
where
    P1: windows_core::Param<LPUNKNOWN>,
    T: windows_core::Interface,
{
    windows_core::link!("ole32.dll" "system" fn CoCreateInstance(rclsid : *const IID, punkouter : *mut core::ffi::c_void, dwclscontext : DWORD, riid : *const IID, ppv : *mut LPVOID) -> windows_core::HRESULT);
    let mut result__ = core::ptr::null_mut();
    unsafe {
        CoCreateInstance(
            rclsid,
            punkouter.param().abi(),
            dwclscontext,
            &T::IID,
            &mut result__,
        )
        .and_then(|| windows_core::imp::from_abi(result__))
    }
}
#[inline]
pub unsafe fn CoInitializeEx(pvreserved: Option<LPVOID>, dwcoinit: DWORD) -> windows_core::HRESULT {
    windows_core::link!("ole32.dll" "system" fn CoInitializeEx(pvreserved : LPVOID, dwcoinit : DWORD) -> windows_core::HRESULT);
    unsafe { CoInitializeEx(pvreserved.unwrap_or(core::mem::zeroed()) as _, dwcoinit) }
}
#[inline]
pub unsafe fn CoTaskMemFree(pv: LPVOID) {
    windows_core::link!("ole32.dll" "system" fn CoTaskMemFree(pv : LPVOID));
    unsafe { CoTaskMemFree(pv as _) }
}
#[inline]
pub unsafe fn CoUninitialize() {
    windows_core::link!("ole32.dll" "system" fn CoUninitialize());
    unsafe { CoUninitialize() }
}
#[inline]
pub unsafe fn CreateEventA<P3>(
    lpeventattributes: Option<LPSECURITY_ATTRIBUTES>,
    bmanualreset: bool,
    binitialstate: bool,
    lpname: P3,
) -> HANDLE
where
    P3: windows_core::Param<windows_core::PCSTR>,
{
    windows_core::link!("kernel32.dll" "system" fn CreateEventA(lpeventattributes : LPSECURITY_ATTRIBUTES, bmanualreset : windows_core::BOOL, binitialstate : windows_core::BOOL, lpname : windows_core::PCSTR) -> HANDLE);
    unsafe {
        CreateEventA(
            lpeventattributes.unwrap_or(core::mem::zeroed()) as _,
            bmanualreset.into(),
            binitialstate.into(),
            lpname.param().abi(),
        )
    }
}
#[inline]
pub unsafe fn CreateFileW<P0>(
    lpfilename: P0,
    dwdesiredaccess: DWORD,
    dwsharemode: DWORD,
    lpsecurityattributes: Option<LPSECURITY_ATTRIBUTES>,
    dwcreationdisposition: DWORD,
    dwflagsandattributes: DWORD,
    htemplatefile: Option<HANDLE>,
) -> HANDLE
where
    P0: windows_core::Param<windows_core::PCWSTR>,
{
    windows_core::link!("kernel32.dll" "system" fn CreateFileW(lpfilename : windows_core::PCWSTR, dwdesiredaccess : DWORD, dwsharemode : DWORD, lpsecurityattributes : LPSECURITY_ATTRIBUTES, dwcreationdisposition : DWORD, dwflagsandattributes : DWORD, htemplatefile : HANDLE) -> HANDLE);
    unsafe {
        CreateFileW(
            lpfilename.param().abi(),
            dwdesiredaccess,
            dwsharemode,
            lpsecurityattributes.unwrap_or(core::mem::zeroed()) as _,
            dwcreationdisposition,
            dwflagsandattributes,
            htemplatefile.unwrap_or(core::mem::zeroed()) as _,
        )
    }
}
#[inline]
pub unsafe fn DeviceIoControl(
    hdevice: HANDLE,
    dwiocontrolcode: DWORD,
    lpinbuffer: Option<LPVOID>,
    ninbuffersize: DWORD,
    lpoutbuffer: Option<LPVOID>,
    noutbuffersize: DWORD,
    lpbytesreturned: Option<LPDWORD>,
    lpoverlapped: Option<LPOVERLAPPED>,
) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn DeviceIoControl(hdevice : HANDLE, dwiocontrolcode : DWORD, lpinbuffer : LPVOID, ninbuffersize : DWORD, lpoutbuffer : LPVOID, noutbuffersize : DWORD, lpbytesreturned : LPDWORD, lpoverlapped : LPOVERLAPPED) -> windows_core::BOOL);
    unsafe {
        DeviceIoControl(
            hdevice,
            dwiocontrolcode,
            lpinbuffer.unwrap_or(core::mem::zeroed()) as _,
            ninbuffersize,
            lpoutbuffer.unwrap_or(core::mem::zeroed()) as _,
            noutbuffersize,
            lpbytesreturned.unwrap_or(core::mem::zeroed()) as _,
            lpoverlapped.unwrap_or(core::mem::zeroed()) as _,
        )
    }
}
#[inline]
pub unsafe fn PropVariantClear(pvar: *mut PROPVARIANT) -> windows_core::HRESULT {
    windows_core::link!("ole32.dll" "system" fn PropVariantClear(pvar : *mut PROPVARIANT) -> windows_core::HRESULT);
    unsafe { PropVariantClear(pvar) }
}
#[inline]
pub unsafe fn PropVariantToStringAlloc(
    propvar: *const PROPVARIANT,
    ppszout: *mut windows_core::PWSTR,
) -> windows_core::HRESULT {
    windows_core::link!("propsys.dll" "system" fn PropVariantToStringAlloc(propvar : *const PROPVARIANT, ppszout : *mut windows_core::PWSTR) -> windows_core::HRESULT);
    unsafe { PropVariantToStringAlloc(propvar, ppszout as _) }
}
#[inline]
pub unsafe fn StringFromIID(rclsid: *const IID, lplpsz: *mut LPOLESTR) -> windows_core::HRESULT {
    windows_core::link!("ole32.dll" "system" fn StringFromIID(rclsid : *const IID, lplpsz : *mut LPOLESTR) -> windows_core::HRESULT);
    unsafe { StringFromIID(rclsid, lplpsz as _) }
}
#[inline]
pub unsafe fn WaitForSingleObject(hhandle: HANDLE, dwmilliseconds: DWORD) -> DWORD {
    windows_core::link!("kernel32.dll" "system" fn WaitForSingleObject(hhandle : HANDLE, dwmilliseconds : DWORD) -> DWORD);
    unsafe { WaitForSingleObject(hhandle, dwmilliseconds) }
}
pub const AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY: _AUDCLNT_BUFFERFLAGS = 1;
pub const AUDCLNT_BUFFERFLAGS_SILENT: _AUDCLNT_BUFFERFLAGS = 2;
pub const AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR: _AUDCLNT_BUFFERFLAGS = 4;
pub const AUDCLNT_E_ALREADY_INITIALIZED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890002_u32 as _);
pub const AUDCLNT_E_BUFDURATION_PERIOD_NOT_EQUAL: windows_core::HRESULT =
    windows_core::HRESULT(0x88890013_u32 as _);
pub const AUDCLNT_E_BUFFER_ERROR: windows_core::HRESULT =
    windows_core::HRESULT(0x88890018_u32 as _);
pub const AUDCLNT_E_BUFFER_OPERATION_PENDING: windows_core::HRESULT =
    windows_core::HRESULT(0x8889000B_u32 as _);
pub const AUDCLNT_E_BUFFER_SIZE_ERROR: windows_core::HRESULT =
    windows_core::HRESULT(0x88890016_u32 as _);
pub const AUDCLNT_E_BUFFER_SIZE_NOT_ALIGNED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890019_u32 as _);
pub const AUDCLNT_E_BUFFER_TOO_LARGE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890006_u32 as _);
pub const AUDCLNT_E_CPUUSAGE_EXCEEDED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890017_u32 as _);
pub const AUDCLNT_E_DEVICE_INVALIDATED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890004_u32 as _);
pub const AUDCLNT_E_DEVICE_IN_USE: windows_core::HRESULT =
    windows_core::HRESULT(0x8889000A_u32 as _);
pub const AUDCLNT_E_EFFECT_NOT_AVAILABLE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890041_u32 as _);
pub const AUDCLNT_E_EFFECT_STATE_READ_ONLY: windows_core::HRESULT =
    windows_core::HRESULT(0x88890042_u32 as _);
pub const AUDCLNT_E_ENDPOINT_CREATE_FAILED: windows_core::HRESULT =
    windows_core::HRESULT(0x8889000F_u32 as _);
pub const AUDCLNT_E_ENDPOINT_OFFLOAD_NOT_CAPABLE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890022_u32 as _);
pub const AUDCLNT_E_ENGINE_FORMAT_LOCKED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890029_u32 as _);
pub const AUDCLNT_E_ENGINE_PERIODICITY_LOCKED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890028_u32 as _);
pub const AUDCLNT_E_EVENTHANDLE_NOT_EXPECTED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890011_u32 as _);
pub const AUDCLNT_E_EVENTHANDLE_NOT_SET: windows_core::HRESULT =
    windows_core::HRESULT(0x88890014_u32 as _);
pub const AUDCLNT_E_EXCLUSIVE_MODE_NOT_ALLOWED: windows_core::HRESULT =
    windows_core::HRESULT(0x8889000E_u32 as _);
pub const AUDCLNT_E_EXCLUSIVE_MODE_ONLY: windows_core::HRESULT =
    windows_core::HRESULT(0x88890012_u32 as _);
pub const AUDCLNT_E_HEADTRACKING_ENABLED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890030_u32 as _);
pub const AUDCLNT_E_HEADTRACKING_UNSUPPORTED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890040_u32 as _);
pub const AUDCLNT_E_INCORRECT_BUFFER_SIZE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890015_u32 as _);
pub const AUDCLNT_E_INVALID_DEVICE_PERIOD: windows_core::HRESULT =
    windows_core::HRESULT(0x88890020_u32 as _);
pub const AUDCLNT_E_INVALID_SIZE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890009_u32 as _);
pub const AUDCLNT_E_INVALID_STREAM_FLAG: windows_core::HRESULT =
    windows_core::HRESULT(0x88890021_u32 as _);
pub const AUDCLNT_E_NONOFFLOAD_MODE_ONLY: windows_core::HRESULT =
    windows_core::HRESULT(0x88890025_u32 as _);
pub const AUDCLNT_E_NOT_INITIALIZED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890001_u32 as _);
pub const AUDCLNT_E_NOT_STOPPED: windows_core::HRESULT = windows_core::HRESULT(0x88890005_u32 as _);
pub const AUDCLNT_E_OFFLOAD_MODE_ONLY: windows_core::HRESULT =
    windows_core::HRESULT(0x88890024_u32 as _);
pub const AUDCLNT_E_OUT_OF_OFFLOAD_RESOURCES: windows_core::HRESULT =
    windows_core::HRESULT(0x88890023_u32 as _);
pub const AUDCLNT_E_OUT_OF_ORDER: windows_core::HRESULT =
    windows_core::HRESULT(0x88890007_u32 as _);
pub const AUDCLNT_E_RAW_MODE_UNSUPPORTED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890027_u32 as _);
pub const AUDCLNT_E_RESOURCES_INVALIDATED: windows_core::HRESULT =
    windows_core::HRESULT(0x88890026_u32 as _);
pub const AUDCLNT_E_SERVICE_NOT_RUNNING: windows_core::HRESULT =
    windows_core::HRESULT(0x88890010_u32 as _);
pub const AUDCLNT_E_THREAD_NOT_REGISTERED: windows_core::HRESULT =
    windows_core::HRESULT(0x8889000C_u32 as _);
pub const AUDCLNT_E_UNSUPPORTED_FORMAT: windows_core::HRESULT =
    windows_core::HRESULT(0x88890008_u32 as _);
pub const AUDCLNT_E_WRONG_ENDPOINT_TYPE: windows_core::HRESULT =
    windows_core::HRESULT(0x88890003_u32 as _);
pub type AUDCLNT_SHAREMODE = i32;
pub const AUDCLNT_SHAREMODE_EXCLUSIVE: AUDCLNT_SHAREMODE = 1;
pub const AUDCLNT_SHAREMODE_SHARED: AUDCLNT_SHAREMODE = 0;
pub const AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM: u32 = 2147483648;
pub const AUDCLNT_STREAMFLAGS_EVENTCALLBACK: i32 = 262144;
pub const AUDCLNT_STREAMFLAGS_LOOPBACK: i32 = 131072;
pub const AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY: i32 = 134217728;
pub type AUDCLNT_STREAMOPTIONS = i32;
pub const AUDCLNT_STREAMOPTIONS_AMBISONICS: AUDCLNT_STREAMOPTIONS = 4;
pub const AUDCLNT_STREAMOPTIONS_MATCH_FORMAT: AUDCLNT_STREAMOPTIONS = 2;
pub const AUDCLNT_STREAMOPTIONS_NONE: AUDCLNT_STREAMOPTIONS = 0;
pub const AUDCLNT_STREAMOPTIONS_RAW: AUDCLNT_STREAMOPTIONS = 1;
pub const AUDCLNT_S_BUFFER_EMPTY: SCODE = 0x8890001_u32 as _;
pub const AUDCLNT_S_POSITION_STALLED: SCODE = 0x8890003_u32 as _;
pub const AUDCLNT_S_THREAD_ALREADY_REGISTERED: SCODE = 0x8890002_u32 as _;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct AUDIOCLIENT_ACTIVATION_PARAMS {
    pub ActivationType: AUDIOCLIENT_ACTIVATION_TYPE,
    pub Anonymous: AUDIOCLIENT_ACTIVATION_PARAMS_0,
}
impl Default for AUDIOCLIENT_ACTIVATION_PARAMS {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union AUDIOCLIENT_ACTIVATION_PARAMS_0 {
    pub ProcessLoopbackParams: AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS,
}
impl Default for AUDIOCLIENT_ACTIVATION_PARAMS_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
pub type AUDIOCLIENT_ACTIVATION_TYPE = i32;
pub const AUDIOCLIENT_ACTIVATION_TYPE_PROCESS_LOOPBACK: AUDIOCLIENT_ACTIVATION_TYPE = 1;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct AUDIOCLIENT_PROCESS_LOOPBACK_PARAMS {
    pub TargetProcessId: DWORD,
    pub ProcessLoopbackMode: PROCESS_LOOPBACK_MODE,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct AUDIO_EFFECT {
    pub id: windows_core::GUID,
    pub canSetState: windows_core::BOOL,
    pub state: AUDIO_EFFECT_STATE,
}
pub type AUDIO_EFFECT_STATE = i32;
pub const AUDIO_EFFECT_STATE_ON: AUDIO_EFFECT_STATE = 1;
pub type AUDIO_STREAM_CATEGORY = i32;
pub const AudioCategory_Alerts: AUDIO_STREAM_CATEGORY = 4;
pub const AudioCategory_Communications: AUDIO_STREAM_CATEGORY = 3;
pub const AudioCategory_FarFieldSpeech: AUDIO_STREAM_CATEGORY = 12;
pub const AudioCategory_ForegroundOnlyMedia: AUDIO_STREAM_CATEGORY = 1;
pub const AudioCategory_GameChat: AUDIO_STREAM_CATEGORY = 8;
pub const AudioCategory_GameEffects: AUDIO_STREAM_CATEGORY = 6;
pub const AudioCategory_GameMedia: AUDIO_STREAM_CATEGORY = 7;
pub const AudioCategory_Media: AUDIO_STREAM_CATEGORY = 11;
pub const AudioCategory_Movie: AUDIO_STREAM_CATEGORY = 10;
pub const AudioCategory_Other: AUDIO_STREAM_CATEGORY = 0;
pub const AudioCategory_SoundEffects: AUDIO_STREAM_CATEGORY = 5;
pub const AudioCategory_Speech: AUDIO_STREAM_CATEGORY = 9;
pub const AudioCategory_UniformSpeech: AUDIO_STREAM_CATEGORY = 13;
pub const AudioCategory_VoiceTyping: AUDIO_STREAM_CATEGORY = 14;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct AudioClientProperties {
    pub cbSize: UINT32,
    pub bIsOffload: windows_core::BOOL,
    pub eCategory: AUDIO_STREAM_CATEGORY,
    pub Options: AUDCLNT_STREAMOPTIONS,
}
pub type AudioSessionDisconnectReason = i32;
pub type AudioSessionState = i32;
pub const AudioSessionStateActive: AudioSessionState = 1;
pub const AudioSessionStateExpired: AudioSessionState = 2;
pub const AudioSessionStateInactive: AudioSessionState = 0;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct BLOB {
    pub cbSize: ULONG,
    pub pBlobData: *mut BYTE,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct BSTRBLOB {
    pub cbSize: ULONG,
    pub pData: *mut BYTE,
}
pub type BYTE = u8;
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABOOL {
    pub cElems: ULONG,
    pub pElems: *mut VARIANT_BOOL,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::BSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CABSTRBLOB {
    pub cElems: ULONG,
    pub pElems: *mut BSTRBLOB,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAC {
    pub cElems: ULONG,
    pub pElems: *mut CHAR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACLIPDATA {
    pub cElems: ULONG,
    pub pElems: *mut CLIPDATA,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACLSID {
    pub cElems: ULONG,
    pub pElems: *mut CLSID,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CACY {
    pub cElems: ULONG,
    pub pElems: *mut CY,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CADATE {
    pub cElems: ULONG,
    pub pElems: *mut DATE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CADBL {
    pub cElems: ULONG,
    pub pElems: *mut DOUBLE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAFILETIME {
    pub cElems: ULONG,
    pub pElems: *mut FILETIME,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAFLT {
    pub cElems: ULONG,
    pub pElems: *mut FLOAT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAH {
    pub cElems: ULONG,
    pub pElems: *mut LARGE_INTEGER,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAI {
    pub cElems: ULONG,
    pub pElems: *mut SHORT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAL {
    pub cElems: ULONG,
    pub pElems: *mut LONG,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CALPSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::PSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CALPWSTR {
    pub cElems: ULONG,
    pub pElems: *mut windows_core::PWSTR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAPROPVARIANT {
    pub cElems: ULONG,
    pub pElems: *mut PROPVARIANT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CASCODE {
    pub cElems: ULONG,
    pub pElems: *mut SCODE,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUB {
    pub cElems: ULONG,
    pub pElems: *mut UCHAR,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUH {
    pub cElems: ULONG,
    pub pElems: *mut ULARGE_INTEGER,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUI {
    pub cElems: ULONG,
    pub pElems: *mut USHORT,
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct CAUL {
    pub cElems: ULONG,
    pub pElems: *mut ULONG,
}
pub type CHAR = i8;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CLIPDATA {
    pub cbSize: ULONG,
    pub ulClipFmt: LONG,
    pub pClipData: *mut BYTE,
}
pub type CLSCTX = i32;
pub const CLSCTX_ALL: CLSCTX = 23;
pub type CLSID = windows_core::GUID;
pub type COINIT = i32;
pub const COINIT_APARTMENTTHREADED: COINIT = 2;
pub const COINIT_MULTITHREADED: COINIT = 0;
#[repr(C)]
#[derive(Clone, Copy)]
pub union CY {
    pub Anonymous: CY_0,
    pub int64: LONGLONG,
}
impl Default for CY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CY_0 {
    pub Lo: ULONG,
    pub Hi: LONG,
}
pub const Connector: PartType = 0;
pub type ConnectorType = i32;
pub type DATE = f64;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DECIMAL {
    pub wReserved: USHORT,
    pub Anonymous: DECIMAL_0,
    pub Hi32: ULONG,
    pub Anonymous2: DECIMAL_1,
}
impl Default for DECIMAL {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union DECIMAL_0 {
    pub Anonymous: DECIMAL_0_0,
    pub signscale: USHORT,
}
impl Default for DECIMAL_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_0_0 {
    pub scale: BYTE,
    pub sign: BYTE,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union DECIMAL_1 {
    pub Anonymous: DECIMAL_1_0,
    pub Lo64: ULONGLONG,
}
impl Default for DECIMAL_1 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_1_0 {
    pub Lo32: ULONG,
    pub Mid32: ULONG,
}
pub const DEVICE_STATE_ACTIVE: i32 = 1;
pub const DEVICE_STATE_DISABLED: i32 = 2;
pub const DEVICE_STATE_NOTPRESENT: i32 = 4;
pub const DEVICE_STATE_UNPLUGGED: i32 = 8;
pub const DEVINTERFACE_AUDIO_CAPTURE: windows_core::GUID =
    windows_core::GUID::from_u128(0x2eef81be_33fa_4800_9670_1cd474972c3f);
pub const DEVINTERFACE_AUDIO_RENDER: windows_core::GUID =
    windows_core::GUID::from_u128(0xe6327cad_dcec_4949_ae8a_991e976a79d2);
pub type DOUBLE = f64;
pub type DWORD = u32;
pub type DataFlow = i32;
pub const DisconnectReasonDeviceRemoval: AudioSessionDisconnectReason = 0;
pub const DisconnectReasonExclusiveModeOverride: AudioSessionDisconnectReason = 5;
pub const DisconnectReasonFormatChanged: AudioSessionDisconnectReason = 2;
pub const DisconnectReasonServerShutdown: AudioSessionDisconnectReason = 1;
pub const DisconnectReasonSessionDisconnected: AudioSessionDisconnectReason = 4;
pub const DisconnectReasonSessionLogoff: AudioSessionDisconnectReason = 3;
pub type EDataFlow = i32;
pub const ENDPOINT_HARDWARE_SUPPORT_METER: i32 = 4;
pub const ENDPOINT_HARDWARE_SUPPORT_MUTE: i32 = 2;
pub const ENDPOINT_HARDWARE_SUPPORT_VOLUME: i32 = 1;
pub const ERROR_INSUFFICIENT_BUFFER: i32 = 122;
pub const ERROR_MORE_DATA: i32 = 234;
pub type ERole = i32;
pub const E_INVALIDARG: windows_core::HRESULT = windows_core::HRESULT(0x80070057_u32 as _);
pub const E_NOINTERFACE: windows_core::HRESULT = windows_core::HRESULT(0x80004002_u32 as _);
pub const E_POINTER: windows_core::HRESULT = windows_core::HRESULT(0x80004003_u32 as _);
pub const FALSE: i32 = 0;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct FILETIME {
    pub dwLowDateTime: DWORD,
    pub dwHighDateTime: DWORD,
}
pub const FILE_SHARE_READ: i32 = 1;
pub const FILE_SHARE_WRITE: i32 = 2;
pub type FLOAT = f32;
pub const GENERIC_READ: u32 = 2147483648;
pub const GENERIC_WRITE: i32 = 1073741824;
pub type HANDLE = *mut core::ffi::c_void;
windows_core::imp::define_interface!(
    IAcousticEchoCancellationControl,
    IAcousticEchoCancellationControl_Vtbl,
    0xf4ae25b5_aaa3_437d_b6b3_dbbe2d0e9549
);
windows_core::imp::interface_hierarchy!(IAcousticEchoCancellationControl, windows_core::IUnknown);
impl IAcousticEchoCancellationControl {
    pub unsafe fn SetEchoCancellationRenderEndpoint<P0>(
        &self,
        endpointid: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).SetEchoCancellationRenderEndpoint)(
                windows_core::Interface::as_raw(self),
                endpointid.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IAcousticEchoCancellationControl_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub SetEchoCancellationRenderEndpoint: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IActivateAudioInterfaceAsyncOperation,
    IActivateAudioInterfaceAsyncOperation_Vtbl,
    0x72a22d78_cde4_431d_b8cc_843a71199b6d
);
windows_core::imp::interface_hierarchy!(
    IActivateAudioInterfaceAsyncOperation,
    windows_core::IUnknown
);
impl IActivateAudioInterfaceAsyncOperation {
    pub unsafe fn GetActivateResult(
        &self,
        activateresult: *mut windows_core::HRESULT,
        activatedinterface: *mut Option<windows_core::IUnknown>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetActivateResult)(
                windows_core::Interface::as_raw(self),
                activateresult as _,
                activatedinterface as _,
            )
        }
    }
}
#[repr(C)]
pub struct IActivateAudioInterfaceAsyncOperation_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetActivateResult: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::HRESULT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IActivateAudioInterfaceCompletionHandler,
    IActivateAudioInterfaceCompletionHandler_Vtbl,
    0x41d949ab_9862_444a_80f6_c261334da5eb
);
windows_core::imp::interface_hierarchy!(
    IActivateAudioInterfaceCompletionHandler,
    windows_core::IUnknown
);
impl IActivateAudioInterfaceCompletionHandler {
    pub unsafe fn ActivateCompleted<P0>(&self, activateoperation: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IActivateAudioInterfaceAsyncOperation>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).ActivateCompleted)(
                windows_core::Interface::as_raw(self),
                activateoperation.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IActivateAudioInterfaceCompletionHandler_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub ActivateCompleted: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
pub trait IActivateAudioInterfaceCompletionHandler_Impl: windows_core::IUnknownImpl {
    fn ActivateCompleted(
        &self,
        activateoperation: windows_core::Ref<IActivateAudioInterfaceAsyncOperation>,
    ) -> windows_core::HRESULT;
}
impl IActivateAudioInterfaceCompletionHandler_Vtbl {
    pub const fn new<
        Identity: IActivateAudioInterfaceCompletionHandler_Impl,
        const OFFSET: isize,
    >() -> Self {
        unsafe extern "system" fn ActivateCompleted<
            Identity: IActivateAudioInterfaceCompletionHandler_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            activateoperation: *mut core::ffi::c_void,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IActivateAudioInterfaceCompletionHandler_Impl::ActivateCompleted(
                    this,
                    core::mem::transmute_copy(&activateoperation),
                )
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            ActivateCompleted: ActivateCompleted::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IActivateAudioInterfaceCompletionHandler as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IActivateAudioInterfaceCompletionHandler {}
windows_core::imp::define_interface!(
    IAudioCaptureClient,
    IAudioCaptureClient_Vtbl,
    0xc8adbd64_e71e_48a0_a4de_185c395cd317
);
windows_core::imp::interface_hierarchy!(IAudioCaptureClient, windows_core::IUnknown);
impl IAudioCaptureClient {
    pub unsafe fn GetBuffer(
        &self,
        ppdata: *mut *mut BYTE,
        pnumframestoread: *mut UINT32,
        pdwflags: *mut DWORD,
        pu64deviceposition: Option<*mut UINT64>,
        pu64qpcposition: Option<*mut UINT64>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetBuffer)(
                windows_core::Interface::as_raw(self),
                ppdata as _,
                pnumframestoread as _,
                pdwflags as _,
                pu64deviceposition.unwrap_or(core::mem::zeroed()) as _,
                pu64qpcposition.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn ReleaseBuffer(&self, numframesread: UINT32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).ReleaseBuffer)(
                windows_core::Interface::as_raw(self),
                numframesread,
            )
        }
    }
    pub unsafe fn GetNextPacketSize(
        &self,
        pnumframesinnextpacket: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetNextPacketSize)(
                windows_core::Interface::as_raw(self),
                pnumframesinnextpacket as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioCaptureClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetBuffer: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut BYTE,
        *mut UINT32,
        *mut DWORD,
        *mut UINT64,
        *mut UINT64,
    ) -> windows_core::HRESULT,
    pub ReleaseBuffer:
        unsafe extern "system" fn(*mut core::ffi::c_void, UINT32) -> windows_core::HRESULT,
    pub GetNextPacketSize:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT32) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClient,
    IAudioClient_Vtbl,
    0x1cb9ad4c_dbfa_4c32_b178_c2f568a703b2
);
windows_core::imp::interface_hierarchy!(IAudioClient, windows_core::IUnknown);
impl IAudioClient {
    pub unsafe fn Initialize(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        streamflags: DWORD,
        hnsbufferduration: REFERENCE_TIME,
        hnsperiodicity: REFERENCE_TIME,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<LPCGUID>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Initialize)(
                windows_core::Interface::as_raw(self),
                sharemode,
                streamflags,
                hnsbufferduration,
                hnsperiodicity,
                pformat,
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetBufferSize(&self, pnumbufferframes: *mut UINT32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetBufferSize)(
                windows_core::Interface::as_raw(self),
                pnumbufferframes as _,
            )
        }
    }
    pub unsafe fn GetStreamLatency(
        &self,
        phnslatency: *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetStreamLatency)(
                windows_core::Interface::as_raw(self),
                phnslatency as _,
            )
        }
    }
    pub unsafe fn GetCurrentPadding(
        &self,
        pnumpaddingframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCurrentPadding)(
                windows_core::Interface::as_raw(self),
                pnumpaddingframes as _,
            )
        }
    }
    pub unsafe fn IsFormatSupported(
        &self,
        sharemode: AUDCLNT_SHAREMODE,
        pformat: *const WAVEFORMATEX,
        ppclosestmatch: Option<*mut *mut WAVEFORMATEX>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).IsFormatSupported)(
                windows_core::Interface::as_raw(self),
                sharemode,
                pformat,
                ppclosestmatch.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetMixFormat(
        &self,
        ppdeviceformat: *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMixFormat)(
                windows_core::Interface::as_raw(self),
                ppdeviceformat as _,
            )
        }
    }
    pub unsafe fn GetDevicePeriod(
        &self,
        phnsdefaultdeviceperiod: Option<*mut REFERENCE_TIME>,
        phnsminimumdeviceperiod: Option<*mut REFERENCE_TIME>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDevicePeriod)(
                windows_core::Interface::as_raw(self),
                phnsdefaultdeviceperiod.unwrap_or(core::mem::zeroed()) as _,
                phnsminimumdeviceperiod.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn Start(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Start)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn Stop(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Stop)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn Reset(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Reset)(windows_core::Interface::as_raw(self))
        }
    }
    pub unsafe fn SetEventHandle(&self, eventhandle: HANDLE) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetEventHandle)(
                windows_core::Interface::as_raw(self),
                eventhandle,
            )
        }
    }
    pub unsafe fn GetService<T>(&self) -> windows_core::Result<T>
    where
        T: windows_core::Interface,
    {
        let mut result__ = core::ptr::null_mut();
        unsafe {
            (windows_core::Interface::vtable(self).GetService)(
                windows_core::Interface::as_raw(self),
                &T::IID,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IAudioClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Initialize: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDCLNT_SHAREMODE,
        DWORD,
        REFERENCE_TIME,
        REFERENCE_TIME,
        *const WAVEFORMATEX,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetBufferSize:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT32) -> windows_core::HRESULT,
    pub GetStreamLatency: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
    pub GetCurrentPadding:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT32) -> windows_core::HRESULT,
    pub IsFormatSupported: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDCLNT_SHAREMODE,
        *const WAVEFORMATEX,
        *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT,
    pub GetMixFormat: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut WAVEFORMATEX,
    ) -> windows_core::HRESULT,
    pub GetDevicePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut REFERENCE_TIME,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
    pub Start: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub Stop: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub Reset: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub SetEventHandle:
        unsafe extern "system" fn(*mut core::ffi::c_void, HANDLE) -> windows_core::HRESULT,
    pub GetService: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const IID,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClient2,
    IAudioClient2_Vtbl,
    0x726778cd_f60a_4eda_82de_e47610cd78aa
);
impl core::ops::Deref for IAudioClient2 {
    type Target = IAudioClient;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(IAudioClient2, windows_core::IUnknown, IAudioClient);
impl IAudioClient2 {
    pub unsafe fn IsOffloadCapable(
        &self,
        category: AUDIO_STREAM_CATEGORY,
        pboffloadcapable: *mut windows_core::BOOL,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).IsOffloadCapable)(
                windows_core::Interface::as_raw(self),
                category,
                pboffloadcapable as _,
            )
        }
    }
    pub unsafe fn SetClientProperties(
        &self,
        pproperties: *const AudioClientProperties,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetClientProperties)(
                windows_core::Interface::as_raw(self),
                pproperties,
            )
        }
    }
    pub unsafe fn GetBufferSizeLimits(
        &self,
        pformat: *const WAVEFORMATEX,
        beventdriven: bool,
        phnsminbufferduration: *mut REFERENCE_TIME,
        phnsmaxbufferduration: *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetBufferSizeLimits)(
                windows_core::Interface::as_raw(self),
                pformat,
                beventdriven.into(),
                phnsminbufferduration as _,
                phnsmaxbufferduration as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioClient2_Vtbl {
    pub base__: IAudioClient_Vtbl,
    pub IsOffloadCapable: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDIO_STREAM_CATEGORY,
        *mut windows_core::BOOL,
    ) -> windows_core::HRESULT,
    pub SetClientProperties: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const AudioClientProperties,
    ) -> windows_core::HRESULT,
    pub GetBufferSizeLimits: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const WAVEFORMATEX,
        windows_core::BOOL,
        *mut REFERENCE_TIME,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClient3,
    IAudioClient3_Vtbl,
    0x7ed4ee07_8e67_4cd4_8c1a_2b7a5987ad42
);
impl core::ops::Deref for IAudioClient3 {
    type Target = IAudioClient2;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(
    IAudioClient3,
    windows_core::IUnknown,
    IAudioClient,
    IAudioClient2
);
impl IAudioClient3 {
    pub unsafe fn GetSharedModeEnginePeriod(
        &self,
        pformat: *const WAVEFORMATEX,
        pdefaultperiodinframes: *mut UINT32,
        pfundamentalperiodinframes: *mut UINT32,
        pminperiodinframes: *mut UINT32,
        pmaxperiodinframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSharedModeEnginePeriod)(
                windows_core::Interface::as_raw(self),
                pformat,
                pdefaultperiodinframes as _,
                pfundamentalperiodinframes as _,
                pminperiodinframes as _,
                pmaxperiodinframes as _,
            )
        }
    }
    pub unsafe fn GetCurrentSharedModeEnginePeriod(
        &self,
        ppformat: *mut *mut WAVEFORMATEX,
        pcurrentperiodinframes: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCurrentSharedModeEnginePeriod)(
                windows_core::Interface::as_raw(self),
                ppformat as _,
                pcurrentperiodinframes as _,
            )
        }
    }
    pub unsafe fn InitializeSharedAudioStream(
        &self,
        streamflags: DWORD,
        periodinframes: UINT32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<LPCGUID>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).InitializeSharedAudioStream)(
                windows_core::Interface::as_raw(self),
                streamflags,
                periodinframes,
                pformat,
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioClient3_Vtbl {
    pub base__: IAudioClient2_Vtbl,
    pub GetSharedModeEnginePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const WAVEFORMATEX,
        *mut UINT32,
        *mut UINT32,
        *mut UINT32,
        *mut UINT32,
    ) -> windows_core::HRESULT,
    pub GetCurrentSharedModeEnginePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut WAVEFORMATEX,
        *mut UINT32,
    ) -> windows_core::HRESULT,
    pub InitializeSharedAudioStream: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        UINT32,
        *const WAVEFORMATEX,
        LPCGUID,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClock,
    IAudioClock_Vtbl,
    0xcd63314f_3fba_4a1b_812c_ef96358728e7
);
windows_core::imp::interface_hierarchy!(IAudioClock, windows_core::IUnknown);
impl IAudioClock {
    pub unsafe fn GetFrequency(&self, pu64frequency: *mut UINT64) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetFrequency)(
                windows_core::Interface::as_raw(self),
                pu64frequency as _,
            )
        }
    }
    pub unsafe fn GetPosition(
        &self,
        pu64position: *mut UINT64,
        pu64qpcposition: Option<*mut UINT64>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetPosition)(
                windows_core::Interface::as_raw(self),
                pu64position as _,
                pu64qpcposition.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetCharacteristics(
        &self,
        pdwcharacteristics: *mut DWORD,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCharacteristics)(
                windows_core::Interface::as_raw(self),
                pdwcharacteristics as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioClock_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetFrequency:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT64) -> windows_core::HRESULT,
    pub GetPosition: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut UINT64,
        *mut UINT64,
    ) -> windows_core::HRESULT,
    pub GetCharacteristics:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioEffectsChangedNotificationClient,
    IAudioEffectsChangedNotificationClient_Vtbl,
    0xa5ded44f_3c5d_4b2b_bd1e_5dc1ee20bbf6
);
windows_core::imp::interface_hierarchy!(
    IAudioEffectsChangedNotificationClient,
    windows_core::IUnknown
);
#[repr(C)]
pub struct IAudioEffectsChangedNotificationClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    OnAudioEffectsChanged: usize,
}
windows_core::imp::define_interface!(
    IAudioEffectsManager,
    IAudioEffectsManager_Vtbl,
    0x4460b3ae_4b44_4527_8676_7548a8acd260
);
windows_core::imp::interface_hierarchy!(IAudioEffectsManager, windows_core::IUnknown);
impl IAudioEffectsManager {
    pub unsafe fn RegisterAudioEffectsChangedNotificationCallback<P0>(
        &self,
        client: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioEffectsChangedNotificationClient>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterAudioEffectsChangedNotificationCallback)(
                windows_core::Interface::as_raw(self),
                client.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterAudioEffectsChangedNotificationCallback<P0>(
        &self,
        client: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioEffectsChangedNotificationClient>,
    {
        unsafe {
            (windows_core::Interface::vtable(self)
                .UnregisterAudioEffectsChangedNotificationCallback)(
                windows_core::Interface::as_raw(self),
                client.param().abi(),
            )
        }
    }
    pub unsafe fn GetAudioEffects(
        &self,
        effects: *mut *mut AUDIO_EFFECT,
        numeffects: *mut UINT32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetAudioEffects)(
                windows_core::Interface::as_raw(self),
                effects as _,
                numeffects as _,
            )
        }
    }
    pub unsafe fn SetAudioEffectState(
        &self,
        effectid: windows_core::GUID,
        state: AUDIO_EFFECT_STATE,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetAudioEffectState)(
                windows_core::Interface::as_raw(self),
                effectid,
                state,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioEffectsManager_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub RegisterAudioEffectsChangedNotificationCallback:
        unsafe extern "system" fn(
            *mut core::ffi::c_void,
            *mut core::ffi::c_void,
        ) -> windows_core::HRESULT,
    pub UnregisterAudioEffectsChangedNotificationCallback:
        unsafe extern "system" fn(
            *mut core::ffi::c_void,
            *mut core::ffi::c_void,
        ) -> windows_core::HRESULT,
    pub GetAudioEffects: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut AUDIO_EFFECT,
        *mut UINT32,
    ) -> windows_core::HRESULT,
    pub SetAudioEffectState: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::GUID,
        AUDIO_EFFECT_STATE,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioMeterInformation,
    IAudioMeterInformation_Vtbl,
    0xc02216f6_8c67_4b5b_9d00_d008e73e0064
);
windows_core::imp::interface_hierarchy!(IAudioMeterInformation, windows_core::IUnknown);
impl IAudioMeterInformation {
    pub unsafe fn GetPeakValue(&self, pfpeak: *mut f32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetPeakValue)(
                windows_core::Interface::as_raw(self),
                pfpeak as _,
            )
        }
    }
    pub unsafe fn GetMeteringChannelCount(
        &self,
        pnchannelcount: *mut UINT,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetMeteringChannelCount)(
                windows_core::Interface::as_raw(self),
                pnchannelcount as _,
            )
        }
    }
    pub unsafe fn GetChannelsPeakValues(
        &self,
        u32channelcount: UINT32,
        afpeakvalues: *mut f32,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetChannelsPeakValues)(
                windows_core::Interface::as_raw(self),
                u32channelcount,
                afpeakvalues as _,
            )
        }
    }
    pub unsafe fn QueryHardwareSupport(
        &self,
        pdwhardwaresupportmask: *mut DWORD,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).QueryHardwareSupport)(
                windows_core::Interface::as_raw(self),
                pdwhardwaresupportmask as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioMeterInformation_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetPeakValue:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut f32) -> windows_core::HRESULT,
    pub GetMeteringChannelCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetChannelsPeakValues: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT32,
        *mut f32,
    ) -> windows_core::HRESULT,
    pub QueryHardwareSupport:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioRenderClient,
    IAudioRenderClient_Vtbl,
    0xf294acfc_3146_4483_a7bf_addca7c260e2
);
windows_core::imp::interface_hierarchy!(IAudioRenderClient, windows_core::IUnknown);
impl IAudioRenderClient {
    pub unsafe fn GetBuffer(
        &self,
        numframesrequested: UINT32,
        ppdata: *mut *mut BYTE,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetBuffer)(
                windows_core::Interface::as_raw(self),
                numframesrequested,
                ppdata as _,
            )
        }
    }
    pub unsafe fn ReleaseBuffer(
        &self,
        numframeswritten: UINT32,
        dwflags: DWORD,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).ReleaseBuffer)(
                windows_core::Interface::as_raw(self),
                numframeswritten,
                dwflags,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioRenderClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetBuffer: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT32,
        *mut *mut BYTE,
    ) -> windows_core::HRESULT,
    pub ReleaseBuffer:
        unsafe extern "system" fn(*mut core::ffi::c_void, UINT32, DWORD) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionControl,
    IAudioSessionControl_Vtbl,
    0xf4b1a599_7266_4319_a8ca_e70acb11e8cd
);
windows_core::imp::interface_hierarchy!(IAudioSessionControl, windows_core::IUnknown);
impl IAudioSessionControl {
    pub unsafe fn GetState(&self, pretval: *mut AudioSessionState) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetState)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn GetDisplayName(
        &self,
        pretval: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDisplayName)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn SetDisplayName<P0>(
        &self,
        value: P0,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).SetDisplayName)(
                windows_core::Interface::as_raw(self),
                value.param().abi(),
                eventcontext,
            )
        }
    }
    pub unsafe fn GetIconPath(&self, pretval: *mut windows_core::PWSTR) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetIconPath)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn SetIconPath<P0>(&self, value: P0, eventcontext: LPCGUID) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).SetIconPath)(
                windows_core::Interface::as_raw(self),
                value.param().abi(),
                eventcontext,
            )
        }
    }
    pub unsafe fn GetGroupingParam(
        &self,
        pretval: *mut windows_core::GUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetGroupingParam)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn SetGroupingParam(
        &self,
        r#override: LPCGUID,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetGroupingParam)(
                windows_core::Interface::as_raw(self),
                r#override,
                eventcontext,
            )
        }
    }
    pub unsafe fn RegisterAudioSessionNotification<P0>(
        &self,
        newnotifications: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioSessionEvents>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterAudioSessionNotification)(
                windows_core::Interface::as_raw(self),
                newnotifications.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterAudioSessionNotification<P0>(
        &self,
        newnotifications: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioSessionEvents>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterAudioSessionNotification)(
                windows_core::Interface::as_raw(self),
                newnotifications.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IAudioSessionControl_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetState: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut AudioSessionState,
    ) -> windows_core::HRESULT,
    pub GetDisplayName: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub SetDisplayName: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetIconPath: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub SetIconPath: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub GetGroupingParam: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub SetGroupingParam: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        LPCGUID,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub RegisterAudioSessionNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub UnregisterAudioSessionNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionControl2,
    IAudioSessionControl2_Vtbl,
    0xbfb7ff88_7239_4fc9_8fa2_07c950be9c6d
);
impl core::ops::Deref for IAudioSessionControl2 {
    type Target = IAudioSessionControl;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(
    IAudioSessionControl2,
    windows_core::IUnknown,
    IAudioSessionControl
);
impl IAudioSessionControl2 {
    pub unsafe fn GetSessionIdentifier(
        &self,
        pretval: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSessionIdentifier)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn GetSessionInstanceIdentifier(
        &self,
        pretval: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSessionInstanceIdentifier)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn GetProcessId(&self, pretval: *mut DWORD) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetProcessId)(
                windows_core::Interface::as_raw(self),
                pretval as _,
            )
        }
    }
    pub unsafe fn IsSystemSoundsSession(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).IsSystemSoundsSession)(
                windows_core::Interface::as_raw(self),
            )
        }
    }
    pub unsafe fn SetDuckingPreference(&self, optout: bool) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetDuckingPreference)(
                windows_core::Interface::as_raw(self),
                optout.into(),
            )
        }
    }
}
#[repr(C)]
pub struct IAudioSessionControl2_Vtbl {
    pub base__: IAudioSessionControl_Vtbl,
    pub GetSessionIdentifier: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetSessionInstanceIdentifier: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetProcessId:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
    pub IsSystemSoundsSession:
        unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub SetDuckingPreference: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::BOOL,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionEnumerator,
    IAudioSessionEnumerator_Vtbl,
    0xe2f5bb11_0570_40ca_acdd_3aa01277dee8
);
windows_core::imp::interface_hierarchy!(IAudioSessionEnumerator, windows_core::IUnknown);
impl IAudioSessionEnumerator {
    pub unsafe fn GetCount(&self, sessioncount: *mut i32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                sessioncount as _,
            )
        }
    }
    pub unsafe fn GetSession(
        &self,
        sessioncount: i32,
        session: *mut Option<IAudioSessionControl>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSession)(
                windows_core::Interface::as_raw(self),
                sessioncount,
                session as _,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioSessionEnumerator_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut i32) -> windows_core::HRESULT,
    pub GetSession: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        i32,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionEvents,
    IAudioSessionEvents_Vtbl,
    0x24918acc_64b3_37c1_8ca9_74a66e9957a8
);
windows_core::imp::interface_hierarchy!(IAudioSessionEvents, windows_core::IUnknown);
impl IAudioSessionEvents {
    pub unsafe fn OnDisplayNameChanged<P0>(
        &self,
        newdisplayname: P0,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnDisplayNameChanged)(
                windows_core::Interface::as_raw(self),
                newdisplayname.param().abi(),
                eventcontext,
            )
        }
    }
    pub unsafe fn OnIconPathChanged<P0>(
        &self,
        newiconpath: P0,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnIconPathChanged)(
                windows_core::Interface::as_raw(self),
                newiconpath.param().abi(),
                eventcontext,
            )
        }
    }
    pub unsafe fn OnSimpleVolumeChanged(
        &self,
        newvolume: f32,
        newmute: bool,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnSimpleVolumeChanged)(
                windows_core::Interface::as_raw(self),
                newvolume,
                newmute.into(),
                eventcontext,
            )
        }
    }
    pub unsafe fn OnChannelVolumeChanged(
        &self,
        channelcount: DWORD,
        newchannelvolumearray: *const f32,
        changedchannel: DWORD,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnChannelVolumeChanged)(
                windows_core::Interface::as_raw(self),
                channelcount,
                newchannelvolumearray,
                changedchannel,
                eventcontext,
            )
        }
    }
    pub unsafe fn OnGroupingParamChanged(
        &self,
        newgroupingparam: LPCGUID,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnGroupingParamChanged)(
                windows_core::Interface::as_raw(self),
                newgroupingparam,
                eventcontext,
            )
        }
    }
    pub unsafe fn OnStateChanged(&self, newstate: AudioSessionState) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnStateChanged)(
                windows_core::Interface::as_raw(self),
                newstate,
            )
        }
    }
    pub unsafe fn OnSessionDisconnected(
        &self,
        disconnectreason: AudioSessionDisconnectReason,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnSessionDisconnected)(
                windows_core::Interface::as_raw(self),
                disconnectreason,
            )
        }
    }
}
#[repr(C)]
pub struct IAudioSessionEvents_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub OnDisplayNameChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub OnIconPathChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub OnSimpleVolumeChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        f32,
        windows_core::BOOL,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub OnChannelVolumeChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        *const f32,
        DWORD,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub OnGroupingParamChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        LPCGUID,
        LPCGUID,
    ) -> windows_core::HRESULT,
    pub OnStateChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AudioSessionState,
    ) -> windows_core::HRESULT,
    pub OnSessionDisconnected: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AudioSessionDisconnectReason,
    ) -> windows_core::HRESULT,
}
pub trait IAudioSessionEvents_Impl: windows_core::IUnknownImpl {
    fn OnDisplayNameChanged(
        &self,
        newdisplayname: &windows_core::PCWSTR,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT;
    fn OnIconPathChanged(
        &self,
        newiconpath: &windows_core::PCWSTR,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT;
    fn OnSimpleVolumeChanged(
        &self,
        newvolume: f32,
        newmute: windows_core::BOOL,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT;
    fn OnChannelVolumeChanged(
        &self,
        channelcount: DWORD,
        newchannelvolumearray: *const f32,
        changedchannel: DWORD,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT;
    fn OnGroupingParamChanged(
        &self,
        newgroupingparam: LPCGUID,
        eventcontext: LPCGUID,
    ) -> windows_core::HRESULT;
    fn OnStateChanged(&self, newstate: AudioSessionState) -> windows_core::HRESULT;
    fn OnSessionDisconnected(
        &self,
        disconnectreason: AudioSessionDisconnectReason,
    ) -> windows_core::HRESULT;
}
impl IAudioSessionEvents_Vtbl {
    pub const fn new<Identity: IAudioSessionEvents_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn OnDisplayNameChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            newdisplayname: windows_core::PCWSTR,
            eventcontext: LPCGUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnDisplayNameChanged(
                    this,
                    &*(&raw const newdisplayname).cast(),
                    core::mem::transmute_copy(&eventcontext),
                )
            }
        }
        unsafe extern "system" fn OnIconPathChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            newiconpath: windows_core::PCWSTR,
            eventcontext: LPCGUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnIconPathChanged(
                    this,
                    &*(&raw const newiconpath).cast(),
                    core::mem::transmute_copy(&eventcontext),
                )
            }
        }
        unsafe extern "system" fn OnSimpleVolumeChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            newvolume: f32,
            newmute: windows_core::BOOL,
            eventcontext: LPCGUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnSimpleVolumeChanged(
                    this,
                    core::mem::transmute_copy(&newvolume),
                    core::mem::transmute_copy(&newmute),
                    core::mem::transmute_copy(&eventcontext),
                )
            }
        }
        unsafe extern "system" fn OnChannelVolumeChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            channelcount: DWORD,
            newchannelvolumearray: *const f32,
            changedchannel: DWORD,
            eventcontext: LPCGUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnChannelVolumeChanged(
                    this,
                    core::mem::transmute_copy(&channelcount),
                    core::mem::transmute_copy(&newchannelvolumearray),
                    core::mem::transmute_copy(&changedchannel),
                    core::mem::transmute_copy(&eventcontext),
                )
            }
        }
        unsafe extern "system" fn OnGroupingParamChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            newgroupingparam: LPCGUID,
            eventcontext: LPCGUID,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnGroupingParamChanged(
                    this,
                    core::mem::transmute_copy(&newgroupingparam),
                    core::mem::transmute_copy(&eventcontext),
                )
            }
        }
        unsafe extern "system" fn OnStateChanged<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            newstate: AudioSessionState,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnStateChanged(this, core::mem::transmute_copy(&newstate))
            }
        }
        unsafe extern "system" fn OnSessionDisconnected<
            Identity: IAudioSessionEvents_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            disconnectreason: AudioSessionDisconnectReason,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IAudioSessionEvents_Impl::OnSessionDisconnected(
                    this,
                    core::mem::transmute_copy(&disconnectreason),
                )
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            OnDisplayNameChanged: OnDisplayNameChanged::<Identity, OFFSET>,
            OnIconPathChanged: OnIconPathChanged::<Identity, OFFSET>,
            OnSimpleVolumeChanged: OnSimpleVolumeChanged::<Identity, OFFSET>,
            OnChannelVolumeChanged: OnChannelVolumeChanged::<Identity, OFFSET>,
            OnGroupingParamChanged: OnGroupingParamChanged::<Identity, OFFSET>,
            OnStateChanged: OnStateChanged::<Identity, OFFSET>,
            OnSessionDisconnected: OnSessionDisconnected::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IAudioSessionEvents as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IAudioSessionEvents {}
windows_core::imp::define_interface!(
    IAudioSessionManager,
    IAudioSessionManager_Vtbl,
    0xbfa971f1_4d5e_40bb_935e_967039bfbee4
);
windows_core::imp::interface_hierarchy!(IAudioSessionManager, windows_core::IUnknown);
impl IAudioSessionManager {
    pub unsafe fn GetAudioSessionControl(
        &self,
        audiosessionguid: Option<LPCGUID>,
        streamflags: DWORD,
    ) -> windows_core::Result<IAudioSessionControl> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetAudioSessionControl)(
                windows_core::Interface::as_raw(self),
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
                streamflags,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn GetSimpleAudioVolume(
        &self,
        audiosessionguid: Option<LPCGUID>,
        streamflags: DWORD,
    ) -> windows_core::Result<ISimpleAudioVolume> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSimpleAudioVolume)(
                windows_core::Interface::as_raw(self),
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
                streamflags,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IAudioSessionManager_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetAudioSessionControl: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        LPCGUID,
        DWORD,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetSimpleAudioVolume: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        LPCGUID,
        DWORD,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionManager2,
    IAudioSessionManager2_Vtbl,
    0x77aa99a0_1bd6_484f_8bc7_2c654c9a9b6f
);
impl core::ops::Deref for IAudioSessionManager2 {
    type Target = IAudioSessionManager;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(
    IAudioSessionManager2,
    windows_core::IUnknown,
    IAudioSessionManager
);
impl IAudioSessionManager2 {
    pub unsafe fn GetSessionEnumerator(&self) -> windows_core::Result<IAudioSessionEnumerator> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSessionEnumerator)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn RegisterSessionNotification<P0>(
        &self,
        sessionnotification: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioSessionNotification>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterSessionNotification)(
                windows_core::Interface::as_raw(self),
                sessionnotification.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterSessionNotification<P0>(
        &self,
        sessionnotification: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioSessionNotification>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterSessionNotification)(
                windows_core::Interface::as_raw(self),
                sessionnotification.param().abi(),
            )
        }
    }
    pub unsafe fn RegisterDuckNotification<P0, P1>(
        &self,
        sessionid: P0,
        ducknotification: P1,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
        P1: windows_core::Param<IAudioVolumeDuckNotification>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterDuckNotification)(
                windows_core::Interface::as_raw(self),
                sessionid.param().abi(),
                ducknotification.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterDuckNotification<P0>(
        &self,
        ducknotification: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IAudioVolumeDuckNotification>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterDuckNotification)(
                windows_core::Interface::as_raw(self),
                ducknotification.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IAudioSessionManager2_Vtbl {
    pub base__: IAudioSessionManager_Vtbl,
    pub GetSessionEnumerator: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub RegisterSessionNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub UnregisterSessionNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub RegisterDuckNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub UnregisterDuckNotification: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionNotification,
    IAudioSessionNotification_Vtbl,
    0x641dd20b_4d41_49cc_aba3_174b9477bb08
);
windows_core::imp::interface_hierarchy!(IAudioSessionNotification, windows_core::IUnknown);
#[repr(C)]
pub struct IAudioSessionNotification_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    OnSessionCreated: usize,
}
windows_core::imp::define_interface!(
    IAudioVolumeDuckNotification,
    IAudioVolumeDuckNotification_Vtbl,
    0xc3b284d4_6d39_4359_b3cf_b56ddb3bb39c
);
windows_core::imp::interface_hierarchy!(IAudioVolumeDuckNotification, windows_core::IUnknown);
#[repr(C)]
pub struct IAudioVolumeDuckNotification_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    OnVolumeDuckNotification: usize,
    OnVolumeUnduckNotification: usize,
}
windows_core::imp::define_interface!(
    IConnector,
    IConnector_Vtbl,
    0x9c2c4058_23f5_41de_877a_df3af236a09e
);
windows_core::imp::interface_hierarchy!(IConnector, windows_core::IUnknown);
impl IConnector {
    pub unsafe fn GetType(&self, ptype: *mut ConnectorType) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetType)(
                windows_core::Interface::as_raw(self),
                ptype as _,
            )
        }
    }
    pub unsafe fn GetDataFlow(&self, pflow: *mut DataFlow) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDataFlow)(
                windows_core::Interface::as_raw(self),
                pflow as _,
            )
        }
    }
    pub unsafe fn ConnectTo<P0>(&self, pconnectto: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<Self>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).ConnectTo)(
                windows_core::Interface::as_raw(self),
                pconnectto.param().abi(),
            )
        }
    }
    pub unsafe fn Disconnect(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Disconnect)(windows_core::Interface::as_raw(
                self,
            ))
        }
    }
    pub unsafe fn IsConnected(
        &self,
        pbconnected: *mut windows_core::BOOL,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).IsConnected)(
                windows_core::Interface::as_raw(self),
                pbconnected as _,
            )
        }
    }
    pub unsafe fn GetConnectedTo(&self, ppconto: *mut Option<Self>) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetConnectedTo)(
                windows_core::Interface::as_raw(self),
                ppconto as _,
            )
        }
    }
    pub unsafe fn GetConnectorIdConnectedTo(
        &self,
        ppwstrconnectorid: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetConnectorIdConnectedTo)(
                windows_core::Interface::as_raw(self),
                ppwstrconnectorid as _,
            )
        }
    }
    pub unsafe fn GetDeviceIdConnectedTo(
        &self,
        ppwstrdeviceid: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDeviceIdConnectedTo)(
                windows_core::Interface::as_raw(self),
                ppwstrdeviceid as _,
            )
        }
    }
}
#[repr(C)]
pub struct IConnector_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetType: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut ConnectorType,
    ) -> windows_core::HRESULT,
    pub GetDataFlow:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DataFlow) -> windows_core::HRESULT,
    pub ConnectTo: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub Disconnect: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
    pub IsConnected: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::BOOL,
    ) -> windows_core::HRESULT,
    pub GetConnectedTo: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetConnectorIdConnectedTo: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetDeviceIdConnectedTo: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IControlChangeNotify,
    IControlChangeNotify_Vtbl,
    0xa09513ed_c709_4d21_bd7b_5f34c47f3947
);
windows_core::imp::interface_hierarchy!(IControlChangeNotify, windows_core::IUnknown);
#[repr(C)]
pub struct IControlChangeNotify_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    OnNotify: usize,
}
windows_core::imp::define_interface!(
    IControlInterface,
    IControlInterface_Vtbl,
    0x45d37c3f_5140_444a_ae24_400789f3cbf3
);
windows_core::imp::interface_hierarchy!(IControlInterface, windows_core::IUnknown);
#[repr(C)]
pub struct IControlInterface_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetName: usize,
    GetIID: usize,
}
windows_core::imp::define_interface!(
    IDeviceTopology,
    IDeviceTopology_Vtbl,
    0x2a07407e_6497_4a18_9787_32f79bd0d98f
);
windows_core::imp::interface_hierarchy!(IDeviceTopology, windows_core::IUnknown);
impl IDeviceTopology {
    pub unsafe fn GetConnectorCount(&self, pcount: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetConnectorCount)(
                windows_core::Interface::as_raw(self),
                pcount as _,
            )
        }
    }
    pub unsafe fn GetConnector(
        &self,
        nindex: UINT,
        ppconnector: *mut Option<IConnector>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetConnector)(
                windows_core::Interface::as_raw(self),
                nindex,
                ppconnector as _,
            )
        }
    }
    pub unsafe fn GetSubunitCount(&self, pcount: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSubunitCount)(
                windows_core::Interface::as_raw(self),
                pcount as _,
            )
        }
    }
    pub unsafe fn GetSubunit(&self, nindex: UINT) -> windows_core::Result<ISubunit> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSubunit)(
                windows_core::Interface::as_raw(self),
                nindex,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn GetPartById(&self, nid: UINT) -> windows_core::Result<IPart> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetPartById)(
                windows_core::Interface::as_raw(self),
                nid,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn GetDeviceId(
        &self,
        ppwstrdeviceid: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDeviceId)(
                windows_core::Interface::as_raw(self),
                ppwstrdeviceid as _,
            )
        }
    }
    pub unsafe fn GetSignalPath<P0, P1>(
        &self,
        pipartfrom: P0,
        pipartto: P1,
        brejectmixedpaths: bool,
    ) -> windows_core::Result<IPartsList>
    where
        P0: windows_core::Param<IPart>,
        P1: windows_core::Param<IPart>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSignalPath)(
                windows_core::Interface::as_raw(self),
                pipartfrom.param().abi(),
                pipartto.param().abi(),
                brejectmixedpaths.into(),
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IDeviceTopology_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetConnectorCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetConnector: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetSubunitCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetSubunit: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetPartById: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetDeviceId: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetSignalPath: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
        windows_core::BOOL,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IDispatch,
    IDispatch_Vtbl,
    0x00020400_0000_0000_c000_000000000046
);
windows_core::imp::interface_hierarchy!(IDispatch, windows_core::IUnknown);
#[repr(C)]
pub struct IDispatch_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetTypeInfoCount: usize,
    GetTypeInfo: usize,
    GetIDsOfNames: usize,
    Invoke: usize,
}
pub type IID = windows_core::GUID;
windows_core::imp::define_interface!(
    IMMDevice,
    IMMDevice_Vtbl,
    0xd666063f_1587_4e43_81f1_b948e807363f
);
windows_core::imp::interface_hierarchy!(IMMDevice, windows_core::IUnknown);
impl IMMDevice {
    pub unsafe fn Activate<T>(
        &self,
        dwclsctx: DWORD,
        pactivationparams: Option<*const PROPVARIANT>,
    ) -> windows_core::Result<T>
    where
        T: windows_core::Interface,
    {
        let mut result__ = core::ptr::null_mut();
        unsafe {
            (windows_core::Interface::vtable(self).Activate)(
                windows_core::Interface::as_raw(self),
                &T::IID,
                dwclsctx,
                pactivationparams.unwrap_or(core::mem::zeroed()) as _,
                &mut result__,
            )
            .and_then(|| windows_core::imp::from_abi(result__))
        }
    }
    pub unsafe fn OpenPropertyStore(
        &self,
        stgmaccess: DWORD,
        ppproperties: *mut Option<IPropertyStore>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OpenPropertyStore)(
                windows_core::Interface::as_raw(self),
                stgmaccess,
                ppproperties as _,
            )
        }
    }
    pub unsafe fn GetId(&self, ppstrid: *mut windows_core::PWSTR) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetId)(
                windows_core::Interface::as_raw(self),
                ppstrid as _,
            )
        }
    }
    pub unsafe fn GetState(&self, pdwstate: *mut DWORD) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetState)(
                windows_core::Interface::as_raw(self),
                pdwstate as _,
            )
        }
    }
}
#[repr(C)]
pub struct IMMDevice_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Activate: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const IID,
        DWORD,
        *const PROPVARIANT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub OpenPropertyStore: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetId: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetState:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IMMDeviceCollection,
    IMMDeviceCollection_Vtbl,
    0x0bd7a1be_7a1a_44db_8397_cc5392387b5e
);
windows_core::imp::interface_hierarchy!(IMMDeviceCollection, windows_core::IUnknown);
impl IMMDeviceCollection {
    pub unsafe fn GetCount(&self, pcdevices: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                pcdevices as _,
            )
        }
    }
    pub unsafe fn Item(
        &self,
        ndevice: UINT,
        ppdevice: *mut Option<IMMDevice>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Item)(
                windows_core::Interface::as_raw(self),
                ndevice,
                ppdevice as _,
            )
        }
    }
}
#[repr(C)]
pub struct IMMDeviceCollection_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub Item: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IMMDeviceEnumerator,
    IMMDeviceEnumerator_Vtbl,
    0xa95664d2_9614_4f35_a746_de8db63617e6
);
windows_core::imp::interface_hierarchy!(IMMDeviceEnumerator, windows_core::IUnknown);
impl IMMDeviceEnumerator {
    pub unsafe fn EnumAudioEndpoints(
        &self,
        dataflow: EDataFlow,
        dwstatemask: DWORD,
        ppdevices: *mut Option<IMMDeviceCollection>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).EnumAudioEndpoints)(
                windows_core::Interface::as_raw(self),
                dataflow,
                dwstatemask,
                ppdevices as _,
            )
        }
    }
    pub unsafe fn GetDefaultAudioEndpoint(
        &self,
        dataflow: EDataFlow,
        role: ERole,
        ppendpoint: *mut Option<IMMDevice>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDefaultAudioEndpoint)(
                windows_core::Interface::as_raw(self),
                dataflow,
                role,
                ppendpoint as _,
            )
        }
    }
    pub unsafe fn GetDevice<P0>(
        &self,
        pwstrid: P0,
        ppdevice: *mut Option<IMMDevice>,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).GetDevice)(
                windows_core::Interface::as_raw(self),
                pwstrid.param().abi(),
                ppdevice as _,
            )
        }
    }
    pub unsafe fn RegisterEndpointNotificationCallback<P0>(
        &self,
        pclient: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IMMNotificationClient>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterEndpointNotificationCallback)(
                windows_core::Interface::as_raw(self),
                pclient.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterEndpointNotificationCallback<P0>(
        &self,
        pclient: P0,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IMMNotificationClient>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterEndpointNotificationCallback)(
                windows_core::Interface::as_raw(self),
                pclient.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IMMDeviceEnumerator_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub EnumAudioEndpoints: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        EDataFlow,
        DWORD,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetDefaultAudioEndpoint: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        EDataFlow,
        ERole,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetDevice: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub RegisterEndpointNotificationCallback: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    )
        -> windows_core::HRESULT,
    pub UnregisterEndpointNotificationCallback: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    )
        -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IMMEndpoint,
    IMMEndpoint_Vtbl,
    0x1be09788_6894_4089_8586_9a2a6c265ac5
);
windows_core::imp::interface_hierarchy!(IMMEndpoint, windows_core::IUnknown);
impl IMMEndpoint {
    pub unsafe fn GetDataFlow(&self, pdataflow: *mut EDataFlow) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetDataFlow)(
                windows_core::Interface::as_raw(self),
                pdataflow as _,
            )
        }
    }
}
#[repr(C)]
pub struct IMMEndpoint_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetDataFlow:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut EDataFlow) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IMMNotificationClient,
    IMMNotificationClient_Vtbl,
    0x7991eec9_7e89_4d85_8390_6c703cec60c0
);
windows_core::imp::interface_hierarchy!(IMMNotificationClient, windows_core::IUnknown);
impl IMMNotificationClient {
    pub unsafe fn OnDeviceStateChanged<P0>(
        &self,
        pwstrdeviceid: P0,
        dwnewstate: DWORD,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnDeviceStateChanged)(
                windows_core::Interface::as_raw(self),
                pwstrdeviceid.param().abi(),
                dwnewstate,
            )
        }
    }
    pub unsafe fn OnDeviceAdded<P0>(&self, pwstrdeviceid: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnDeviceAdded)(
                windows_core::Interface::as_raw(self),
                pwstrdeviceid.param().abi(),
            )
        }
    }
    pub unsafe fn OnDeviceRemoved<P0>(&self, pwstrdeviceid: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnDeviceRemoved)(
                windows_core::Interface::as_raw(self),
                pwstrdeviceid.param().abi(),
            )
        }
    }
    pub unsafe fn OnDefaultDeviceChanged<P2>(
        &self,
        flow: EDataFlow,
        role: ERole,
        pwstrdefaultdeviceid: P2,
    ) -> windows_core::HRESULT
    where
        P2: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnDefaultDeviceChanged)(
                windows_core::Interface::as_raw(self),
                flow,
                role,
                pwstrdefaultdeviceid.param().abi(),
            )
        }
    }
    pub unsafe fn OnPropertyValueChanged<P0>(
        &self,
        pwstrdeviceid: P0,
        key: PROPERTYKEY,
    ) -> windows_core::HRESULT
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).OnPropertyValueChanged)(
                windows_core::Interface::as_raw(self),
                pwstrdeviceid.param().abi(),
                key,
            )
        }
    }
}
#[repr(C)]
pub struct IMMNotificationClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub OnDeviceStateChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        DWORD,
    ) -> windows_core::HRESULT,
    pub OnDeviceAdded: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
    pub OnDeviceRemoved: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
    pub OnDefaultDeviceChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        EDataFlow,
        ERole,
        windows_core::PCWSTR,
    ) -> windows_core::HRESULT,
    pub OnPropertyValueChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        PROPERTYKEY,
    ) -> windows_core::HRESULT,
}
pub trait IMMNotificationClient_Impl: windows_core::IUnknownImpl {
    fn OnDeviceStateChanged(
        &self,
        pwstrdeviceid: &windows_core::PCWSTR,
        dwnewstate: DWORD,
    ) -> windows_core::HRESULT;
    fn OnDeviceAdded(&self, pwstrdeviceid: &windows_core::PCWSTR) -> windows_core::HRESULT;
    fn OnDeviceRemoved(&self, pwstrdeviceid: &windows_core::PCWSTR) -> windows_core::HRESULT;
    fn OnDefaultDeviceChanged(
        &self,
        flow: EDataFlow,
        role: ERole,
        pwstrdefaultdeviceid: &windows_core::PCWSTR,
    ) -> windows_core::HRESULT;
    fn OnPropertyValueChanged(
        &self,
        pwstrdeviceid: &windows_core::PCWSTR,
        key: &PROPERTYKEY,
    ) -> windows_core::HRESULT;
}
impl IMMNotificationClient_Vtbl {
    pub const fn new<Identity: IMMNotificationClient_Impl, const OFFSET: isize>() -> Self {
        unsafe extern "system" fn OnDeviceStateChanged<
            Identity: IMMNotificationClient_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            pwstrdeviceid: windows_core::PCWSTR,
            dwnewstate: DWORD,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IMMNotificationClient_Impl::OnDeviceStateChanged(
                    this,
                    &*(&raw const pwstrdeviceid).cast(),
                    core::mem::transmute_copy(&dwnewstate),
                )
            }
        }
        unsafe extern "system" fn OnDeviceAdded<
            Identity: IMMNotificationClient_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            pwstrdeviceid: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IMMNotificationClient_Impl::OnDeviceAdded(this, &*(&raw const pwstrdeviceid).cast())
            }
        }
        unsafe extern "system" fn OnDeviceRemoved<
            Identity: IMMNotificationClient_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            pwstrdeviceid: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IMMNotificationClient_Impl::OnDeviceRemoved(
                    this,
                    &*(&raw const pwstrdeviceid).cast(),
                )
            }
        }
        unsafe extern "system" fn OnDefaultDeviceChanged<
            Identity: IMMNotificationClient_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            flow: EDataFlow,
            role: ERole,
            pwstrdefaultdeviceid: windows_core::PCWSTR,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IMMNotificationClient_Impl::OnDefaultDeviceChanged(
                    this,
                    core::mem::transmute_copy(&flow),
                    core::mem::transmute_copy(&role),
                    &*(&raw const pwstrdefaultdeviceid).cast(),
                )
            }
        }
        unsafe extern "system" fn OnPropertyValueChanged<
            Identity: IMMNotificationClient_Impl,
            const OFFSET: isize,
        >(
            this: *mut core::ffi::c_void,
            pwstrdeviceid: windows_core::PCWSTR,
            key: PROPERTYKEY,
        ) -> windows_core::HRESULT {
            unsafe {
                let this: &Identity =
                    &*((this as *const *const ()).offset(OFFSET) as *const Identity);
                IMMNotificationClient_Impl::OnPropertyValueChanged(
                    this,
                    &*(&raw const pwstrdeviceid).cast(),
                    &*(&raw const key).cast(),
                )
            }
        }
        Self {
            base__: windows_core::IUnknown_Vtbl::new::<Identity, OFFSET>(),
            OnDeviceStateChanged: OnDeviceStateChanged::<Identity, OFFSET>,
            OnDeviceAdded: OnDeviceAdded::<Identity, OFFSET>,
            OnDeviceRemoved: OnDeviceRemoved::<Identity, OFFSET>,
            OnDefaultDeviceChanged: OnDefaultDeviceChanged::<Identity, OFFSET>,
            OnPropertyValueChanged: OnPropertyValueChanged::<Identity, OFFSET>,
        }
    }
    pub fn matches(iid: &windows_core::GUID) -> bool {
        iid == &<IMMNotificationClient as windows_core::Interface>::IID
    }
}
impl windows_core::RuntimeName for IMMNotificationClient {}
pub type INT = i32;
pub const INVALID_HANDLE_VALUE: HANDLE = -1 as _;
pub const IOCTL_KS_PROPERTY: i32 = 3080195;
windows_core::imp::define_interface!(IPart, IPart_Vtbl, 0xae2de0e4_5bca_4f2d_aa46_5d13f8fdb3a9);
windows_core::imp::interface_hierarchy!(IPart, windows_core::IUnknown);
impl IPart {
    pub unsafe fn GetName(&self, ppwstrname: *mut windows_core::PWSTR) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetName)(
                windows_core::Interface::as_raw(self),
                ppwstrname as _,
            )
        }
    }
    pub unsafe fn GetLocalId(&self, pnid: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetLocalId)(
                windows_core::Interface::as_raw(self),
                pnid as _,
            )
        }
    }
    pub unsafe fn GetGlobalId(
        &self,
        ppwstrglobalid: *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetGlobalId)(
                windows_core::Interface::as_raw(self),
                ppwstrglobalid as _,
            )
        }
    }
    pub unsafe fn GetPartType(&self, pparttype: *mut PartType) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetPartType)(
                windows_core::Interface::as_raw(self),
                pparttype as _,
            )
        }
    }
    pub unsafe fn GetSubType(&self, psubtype: *mut windows_core::GUID) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetSubType)(
                windows_core::Interface::as_raw(self),
                psubtype as _,
            )
        }
    }
    pub unsafe fn GetControlInterfaceCount(&self, pcount: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetControlInterfaceCount)(
                windows_core::Interface::as_raw(self),
                pcount as _,
            )
        }
    }
    pub unsafe fn GetControlInterface(
        &self,
        nindex: UINT,
        ppinterfacedesc: *mut Option<IControlInterface>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetControlInterface)(
                windows_core::Interface::as_raw(self),
                nindex,
                ppinterfacedesc as _,
            )
        }
    }
    pub unsafe fn EnumPartsIncoming(
        &self,
        ppparts: *mut Option<IPartsList>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).EnumPartsIncoming)(
                windows_core::Interface::as_raw(self),
                ppparts as _,
            )
        }
    }
    pub unsafe fn EnumPartsOutgoing(
        &self,
        ppparts: *mut Option<IPartsList>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).EnumPartsOutgoing)(
                windows_core::Interface::as_raw(self),
                ppparts as _,
            )
        }
    }
    pub unsafe fn GetTopologyObject(
        &self,
        pptopology: *mut Option<IDeviceTopology>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetTopologyObject)(
                windows_core::Interface::as_raw(self),
                pptopology as _,
            )
        }
    }
    pub unsafe fn Activate<T>(
        &self,
        dwclscontext: DWORD,
        result__: *mut Option<T>,
    ) -> windows_core::Result<()>
    where
        T: windows_core::Interface,
    {
        unsafe {
            (windows_core::Interface::vtable(self).Activate)(
                windows_core::Interface::as_raw(self),
                dwclscontext,
                &T::IID,
                result__ as *mut _ as *mut _,
            )
            .ok()
        }
    }
    pub unsafe fn RegisterControlChangeCallback<P1>(
        &self,
        riid: *const windows_core::GUID,
        pnotify: P1,
    ) -> windows_core::HRESULT
    where
        P1: windows_core::Param<IControlChangeNotify>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).RegisterControlChangeCallback)(
                windows_core::Interface::as_raw(self),
                riid,
                pnotify.param().abi(),
            )
        }
    }
    pub unsafe fn UnregisterControlChangeCallback<P0>(&self, pnotify: P0) -> windows_core::HRESULT
    where
        P0: windows_core::Param<IControlChangeNotify>,
    {
        unsafe {
            (windows_core::Interface::vtable(self).UnregisterControlChangeCallback)(
                windows_core::Interface::as_raw(self),
                pnotify.param().abi(),
            )
        }
    }
}
#[repr(C)]
pub struct IPart_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetName: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetLocalId:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetGlobalId: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetPartType:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut PartType) -> windows_core::HRESULT,
    pub GetSubType: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub GetControlInterfaceCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetControlInterface: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub EnumPartsIncoming: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub EnumPartsOutgoing: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetTopologyObject: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub Activate: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        *const IID,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub RegisterControlChangeCallback: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub UnregisterControlChangeCallback: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IPartsList,
    IPartsList_Vtbl,
    0x6daa848c_5eb0_45cc_aea5_998a2cda1ffb
);
windows_core::imp::interface_hierarchy!(IPartsList, windows_core::IUnknown);
impl IPartsList {
    pub unsafe fn GetCount(&self, pcount: *mut UINT) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                pcount as _,
            )
        }
    }
    pub unsafe fn GetPart(
        &self,
        nindex: UINT,
        pppart: *mut Option<IPart>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetPart)(
                windows_core::Interface::as_raw(self),
                nindex,
                pppart as _,
            )
        }
    }
}
#[repr(C)]
pub struct IPartsList_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut UINT) -> windows_core::HRESULT,
    pub GetPart: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        UINT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IPropertyStore,
    IPropertyStore_Vtbl,
    0x886d8eeb_8cf2_4446_8d02_cdba1dbdcf99
);
windows_core::imp::interface_hierarchy!(IPropertyStore, windows_core::IUnknown);
impl IPropertyStore {
    pub unsafe fn GetCount(&self, cprops: *mut DWORD) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                cprops as _,
            )
        }
    }
    pub unsafe fn GetAt(&self, iprop: DWORD, pkey: *mut PROPERTYKEY) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetAt)(
                windows_core::Interface::as_raw(self),
                iprop,
                pkey as _,
            )
        }
    }
    pub unsafe fn GetValue(
        &self,
        key: *const PROPERTYKEY,
        pv: *mut PROPVARIANT,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetValue)(
                windows_core::Interface::as_raw(self),
                key,
                pv,
            )
        }
    }
    pub unsafe fn SetValue(
        &self,
        key: *const PROPERTYKEY,
        propvar: *const PROPVARIANT,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).SetValue)(
                windows_core::Interface::as_raw(self),
                key,
                propvar,
            )
        }
    }
    pub unsafe fn Commit(&self) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Commit)(windows_core::Interface::as_raw(self))
        }
    }
}
#[repr(C)]
pub struct IPropertyStore_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut DWORD) -> windows_core::HRESULT,
    pub GetAt: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        DWORD,
        *mut PROPERTYKEY,
    ) -> windows_core::HRESULT,
    pub GetValue: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const PROPERTYKEY,
        *mut PROPVARIANT,
    ) -> windows_core::HRESULT,
    pub SetValue: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const PROPERTYKEY,
        *const PROPVARIANT,
    ) -> windows_core::HRESULT,
    pub Commit: unsafe extern "system" fn(*mut core::ffi::c_void) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    ISequentialStream,
    ISequentialStream_Vtbl,
    0x0c733a30_2a1c_11ce_ade5_00aa0044773d
);
windows_core::imp::interface_hierarchy!(ISequentialStream, windows_core::IUnknown);
#[repr(C)]
pub struct ISequentialStream_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    Read: usize,
    Write: usize,
}
windows_core::imp::define_interface!(
    ISimpleAudioVolume,
    ISimpleAudioVolume_Vtbl,
    0x87ce5498_68d6_44e5_9215_6da47ef883d8
);
windows_core::imp::interface_hierarchy!(ISimpleAudioVolume, windows_core::IUnknown);
#[repr(C)]
pub struct ISimpleAudioVolume_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    SetMasterVolume: usize,
    GetMasterVolume: usize,
    SetMute: usize,
    GetMute: usize,
}
windows_core::imp::define_interface!(
    IStorage,
    IStorage_Vtbl,
    0x0000000b_0000_0000_c000_000000000046
);
windows_core::imp::interface_hierarchy!(IStorage, windows_core::IUnknown);
#[repr(C)]
pub struct IStorage_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    CreateStream: usize,
    OpenStream: usize,
    CreateStorage: usize,
    OpenStorage: usize,
    CopyTo: usize,
    MoveElementTo: usize,
    Commit: usize,
    Revert: usize,
    EnumElements: usize,
    DestroyElement: usize,
    RenameElement: usize,
    SetElementTimes: usize,
    SetClass: usize,
    SetStateBits: usize,
    Stat: usize,
}
windows_core::imp::define_interface!(
    IStream,
    IStream_Vtbl,
    0x0000000c_0000_0000_c000_000000000046
);
impl core::ops::Deref for IStream {
    type Target = ISequentialStream;
    fn deref(&self) -> &Self::Target {
        unsafe { core::mem::transmute(self) }
    }
}
windows_core::imp::interface_hierarchy!(IStream, windows_core::IUnknown, ISequentialStream);
#[repr(C)]
pub struct IStream_Vtbl {
    pub base__: ISequentialStream_Vtbl,
    Seek: usize,
    SetSize: usize,
    CopyTo: usize,
    Commit: usize,
    Revert: usize,
    LockRegion: usize,
    UnlockRegion: usize,
    Stat: usize,
    Clone: usize,
}
windows_core::imp::define_interface!(
    ISubunit,
    ISubunit_Vtbl,
    0x82149a85_dba6_4487_86bb_ea8f7fefcc71
);
windows_core::imp::interface_hierarchy!(ISubunit, windows_core::IUnknown);
#[repr(C)]
pub struct ISubunit_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union KSDATAFORMAT {
    pub Anonymous: KSDATAFORMAT_0,
    pub Alignment: LONGLONG,
}
impl Default for KSDATAFORMAT {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSDATAFORMAT_0 {
    pub FormatSize: ULONG,
    pub Flags: ULONG,
    pub SampleSize: ULONG,
    pub Reserved: ULONG,
    pub MajorFormat: windows_core::GUID,
    pub SubFormat: windows_core::GUID,
    pub Specifier: windows_core::GUID,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub struct KSIDENTIFIER {
    pub Anonymous: KSIDENTIFIER_0,
}
impl Default for KSIDENTIFIER {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union KSIDENTIFIER_0 {
    pub Anonymous: KSIDENTIFIER_0_0,
    pub Alignment: LONGLONG,
}
impl Default for KSIDENTIFIER_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSIDENTIFIER_0_0 {
    pub Set: windows_core::GUID,
    pub Id: ULONG,
    pub Flags: ULONG,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSMULTIPLE_ITEM {
    pub Size: ULONG,
    pub Count: ULONG,
}
pub type KSPIN_DATAFLOW = i32;
pub const KSPIN_DATAFLOW_IN: KSPIN_DATAFLOW = 1;
pub const KSPIN_DATAFLOW_OUT: KSPIN_DATAFLOW = 2;
pub type KSPROPERTY = KSIDENTIFIER;
pub type KSPROPERTY_PIN = i32;
pub const KSPROPERTY_PIN_CATEGORY: KSPROPERTY_PIN = 11;
pub const KSPROPERTY_PIN_CINSTANCES: KSPROPERTY_PIN = 0;
pub const KSPROPERTY_PIN_COMMUNICATION: KSPROPERTY_PIN = 7;
pub const KSPROPERTY_PIN_CONSTRAINEDDATARANGES: KSPROPERTY_PIN = 13;
pub const KSPROPERTY_PIN_CTYPES: KSPROPERTY_PIN = 1;
pub const KSPROPERTY_PIN_DATAFLOW: KSPROPERTY_PIN = 2;
pub const KSPROPERTY_PIN_DATAINTERSECTION: KSPROPERTY_PIN = 4;
pub const KSPROPERTY_PIN_DATARANGES: KSPROPERTY_PIN = 3;
pub const KSPROPERTY_PIN_GLOBALCINSTANCES: KSPROPERTY_PIN = 8;
pub const KSPROPERTY_PIN_INTERFACES: KSPROPERTY_PIN = 5;
pub const KSPROPERTY_PIN_MEDIUMS: KSPROPERTY_PIN = 6;
pub const KSPROPERTY_PIN_MODEDATAFORMATS: KSPROPERTY_PIN = 16;
pub const KSPROPERTY_PIN_NAME: KSPROPERTY_PIN = 12;
pub const KSPROPERTY_PIN_NECESSARYINSTANCES: KSPROPERTY_PIN = 9;
pub const KSPROPERTY_PIN_PHYSICALCONNECTION: KSPROPERTY_PIN = 10;
pub const KSPROPERTY_PIN_PROPOSEDATAFORMAT: KSPROPERTY_PIN = 14;
pub const KSPROPERTY_PIN_PROPOSEDATAFORMAT2: KSPROPERTY_PIN = 15;
pub const KSPROPERTY_TYPE_GET: i32 = 1;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct KSP_PIN {
    pub Property: KSPROPERTY,
    pub PinId: ULONG,
    pub Anonymous: KSP_PIN_0,
}
impl Default for KSP_PIN {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union KSP_PIN_0 {
    pub Reserved: ULONG,
    pub Flags: ULONG,
}
impl Default for KSP_PIN_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union LARGE_INTEGER {
    pub Anonymous: LARGE_INTEGER_0,
    pub u: LARGE_INTEGER_1,
    pub QuadPart: LONGLONG,
}
impl Default for LARGE_INTEGER {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct LARGE_INTEGER_0 {
    pub LowPart: DWORD,
    pub HighPart: LONG,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct LARGE_INTEGER_1 {
    pub LowPart: DWORD,
    pub HighPart: LONG,
}
pub type LONG = i32;
pub type LONGLONG = i64;
pub type LPCGUID = *const windows_core::GUID;
pub type LPDWORD = *mut u32;
pub type LPOLESTR = *mut u16;
pub type LPOVERLAPPED = *mut OVERLAPPED;
pub type LPSAFEARRAY = *mut SAFEARRAY;
pub type LPSECURITY_ATTRIBUTES = *mut SECURITY_ATTRIBUTES;
pub type LPUNKNOWN = windows_core::IUnknown;
pub type LPVERSIONEDSTREAM = *mut VERSIONEDSTREAM;
pub type LPVOID = *mut core::ffi::c_void;
pub const MMDeviceEnumerator: windows_core::GUID =
    windows_core::GUID::from_u128(0xbcde0395_e52f_467c_8e3d_c4579291692e);
pub const OPEN_EXISTING: i32 = 3;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct OVERLAPPED {
    pub Internal: ULONG_PTR,
    pub InternalHigh: ULONG_PTR,
    pub Anonymous: OVERLAPPED_0,
    pub hEvent: HANDLE,
}
impl Default for OVERLAPPED {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union OVERLAPPED_0 {
    pub Anonymous: OVERLAPPED_0_0,
    pub Pointer: PVOID,
}
impl Default for OVERLAPPED_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct OVERLAPPED_0_0 {
    pub Offset: DWORD,
    pub OffsetHigh: DWORD,
}
pub const PKEY_AudioEngine_DeviceFormat: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows_core::GUID::from_u128(0xf19f064d_082c_4e27_bc73_6882a1bb8e4c),
    pid: 0,
};
pub const PKEY_DeviceInterface_FriendlyName: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows_core::GUID::from_u128(0x026e516e_b814_414b_83cd_856d6fef4822),
    pid: 2,
};
pub const PKEY_Device_DeviceDesc: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows_core::GUID::from_u128(0xa45c254e_df1c_4efd_8020_67d146a850e0),
    pid: 2,
};
pub const PKEY_Device_FriendlyName: PROPERTYKEY = PROPERTYKEY {
    fmtid: windows_core::GUID::from_u128(0xa45c254e_df1c_4efd_8020_67d146a850e0),
    pid: 14,
};
pub type PROCESS_LOOPBACK_MODE = i32;
pub const PROCESS_LOOPBACK_MODE_EXCLUDE_TARGET_PROCESS_TREE: PROCESS_LOOPBACK_MODE = 1;
pub const PROCESS_LOOPBACK_MODE_INCLUDE_TARGET_PROCESS_TREE: PROCESS_LOOPBACK_MODE = 0;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct PROPERTYKEY {
    pub fmtid: windows_core::GUID,
    pub pid: DWORD,
}
pub type PROPVARIANT = tagPROPVARIANT;
pub type PROPVAR_PAD1 = u16;
pub type PROPVAR_PAD2 = u16;
pub type PROPVAR_PAD3 = u16;
pub type PVOID = *mut core::ffi::c_void;
pub type PartType = i32;
pub type REFERENCE_TIME = i64;
#[repr(C, packed(8))]
#[derive(Clone, Copy)]
pub struct SAFEARRAY {
    pub cDims: USHORT,
    pub fFeatures: USHORT,
    pub cbElements: ULONG,
    pub cLocks: ULONG,
    pub pvData: PVOID,
    pub rgsabound: [SAFEARRAYBOUND; 1],
}
impl Default for SAFEARRAY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
#[derive(Clone, Copy, Default)]
pub struct SAFEARRAYBOUND {
    pub cElements: ULONG,
    pub lLbound: LONG,
}
pub type SCODE = i32;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct SECURITY_ATTRIBUTES {
    pub nLength: DWORD,
    pub lpSecurityDescriptor: LPVOID,
    pub bInheritHandle: windows_core::BOOL,
}
pub type SHORT = i16;
pub const SPEAKER_BACK_CENTER: i32 = 256;
pub const SPEAKER_BACK_LEFT: i32 = 16;
pub const SPEAKER_BACK_RIGHT: i32 = 32;
pub const SPEAKER_FRONT_CENTER: i32 = 4;
pub const SPEAKER_FRONT_LEFT: i32 = 1;
pub const SPEAKER_FRONT_LEFT_OF_CENTER: i32 = 64;
pub const SPEAKER_FRONT_RIGHT: i32 = 2;
pub const SPEAKER_FRONT_RIGHT_OF_CENTER: i32 = 128;
pub const SPEAKER_LOW_FREQUENCY: i32 = 8;
pub const SPEAKER_SIDE_LEFT: i32 = 512;
pub const SPEAKER_SIDE_RIGHT: i32 = 1024;
pub const SPEAKER_TOP_BACK_CENTER: i32 = 65536;
pub const SPEAKER_TOP_BACK_LEFT: i32 = 32768;
pub const SPEAKER_TOP_BACK_RIGHT: i32 = 131072;
pub const SPEAKER_TOP_CENTER: i32 = 2048;
pub const SPEAKER_TOP_FRONT_CENTER: i32 = 8192;
pub const SPEAKER_TOP_FRONT_LEFT: i32 = 4096;
pub const SPEAKER_TOP_FRONT_RIGHT: i32 = 16384;
pub const STGM_READ: i32 = 0;
pub const S_OK: windows_core::HRESULT = windows_core::HRESULT(0x0_u32 as _);
pub type UCHAR = u8;
pub type UINT = u32;
pub type UINT32 = u32;
pub type UINT64 = u64;
#[repr(C)]
#[derive(Clone, Copy)]
pub union ULARGE_INTEGER {
    pub Anonymous: ULARGE_INTEGER_0,
    pub u: ULARGE_INTEGER_1,
    pub QuadPart: ULONGLONG,
}
impl Default for ULARGE_INTEGER {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ULARGE_INTEGER_0 {
    pub LowPart: DWORD,
    pub HighPart: DWORD,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct ULARGE_INTEGER_1 {
    pub LowPart: DWORD,
    pub HighPart: DWORD,
}
pub type ULONG = u32;
pub type ULONGLONG = u64;
pub type ULONG_PTR = usize;
pub type USHORT = u16;
pub type VARENUM = i32;
pub type VARIANT_BOOL = i16;
pub type VARTYPE = u16;
#[repr(C, packed(8))]
#[derive(Default)]
pub struct VERSIONEDSTREAM {
    pub guidVersion: windows_core::GUID,
    pub pStream: core::mem::ManuallyDrop<Option<IStream>>,
}
pub const VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK: windows_core::PCWSTR =
    windows_core::w!("VAD\\Process_Loopback");
pub const VT_BLOB: VARENUM = 65;
pub const WAIT_OBJECT_0: windows_core::NTSTATUS = windows_core::NTSTATUS(0x0_u32 as _);
#[repr(C, packed(1))]
#[derive(Clone, Copy, Default)]
pub struct WAVEFORMATEX {
    pub wFormatTag: WORD,
    pub nChannels: WORD,
    pub nSamplesPerSec: DWORD,
    pub nAvgBytesPerSec: DWORD,
    pub nBlockAlign: WORD,
    pub wBitsPerSample: WORD,
    pub cbSize: WORD,
}
#[repr(C, packed(1))]
#[derive(Clone, Copy)]
pub struct WAVEFORMATEXTENSIBLE {
    pub Format: WAVEFORMATEX,
    pub Samples: WAVEFORMATEXTENSIBLE_0,
    pub dwChannelMask: DWORD,
    pub SubFormat: windows_core::GUID,
}
impl Default for WAVEFORMATEXTENSIBLE {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(1))]
#[derive(Clone, Copy)]
pub union WAVEFORMATEXTENSIBLE_0 {
    pub wValidBitsPerSample: WORD,
    pub wSamplesPerBlock: WORD,
    pub wReserved: WORD,
}
impl Default for WAVEFORMATEXTENSIBLE_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
pub const WAVE_FORMAT_EXTENSIBLE: i32 = 65534;
pub const WAVE_FORMAT_IEEE_FLOAT: i32 = 3;
pub const WAVE_FORMAT_PCM: i32 = 1;
pub type WORD = u16;
pub type _AUDCLNT_BUFFERFLAGS = i32;
pub const eAll: EDataFlow = 2;
pub const eCapture: EDataFlow = 1;
pub const eCommunications: ERole = 2;
pub const eConsole: ERole = 0;
pub const eMultimedia: ERole = 1;
pub const eRender: EDataFlow = 0;
#[repr(C, packed(8))]
pub struct tagPROPVARIANT {
    pub Anonymous: tagPROPVARIANT_0,
}
impl Clone for tagPROPVARIANT {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub union tagPROPVARIANT_0 {
    pub Anonymous: core::mem::ManuallyDrop<tagPROPVARIANT_0_0>,
    pub decVal: DECIMAL,
}
impl Clone for tagPROPVARIANT_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub struct tagPROPVARIANT_0_0 {
    pub vt: VARTYPE,
    pub wReserved1: PROPVAR_PAD1,
    pub wReserved2: PROPVAR_PAD2,
    pub wReserved3: PROPVAR_PAD3,
    pub Anonymous: tagPROPVARIANT_0_0_0,
}
impl Clone for tagPROPVARIANT_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C, packed(8))]
pub union tagPROPVARIANT_0_0_0 {
    pub cVal: CHAR,
    pub bVal: UCHAR,
    pub iVal: SHORT,
    pub uiVal: USHORT,
    pub lVal: LONG,
    pub ulVal: ULONG,
    pub intVal: INT,
    pub uintVal: UINT,
    pub hVal: LARGE_INTEGER,
    pub uhVal: ULARGE_INTEGER,
    pub fltVal: FLOAT,
    pub dblVal: DOUBLE,
    pub boolVal: VARIANT_BOOL,
    pub __OBSOLETE__VARIANT_BOOL: VARIANT_BOOL,
    pub scode: SCODE,
    pub cyVal: CY,
    pub date: DATE,
    pub filetime: FILETIME,
    pub puuid: *mut CLSID,
    pub pclipdata: *mut CLIPDATA,
    pub bstrVal: core::mem::ManuallyDrop<windows_core::BSTR>,
    pub bstrblobVal: BSTRBLOB,
    pub blob: BLOB,
    pub pszVal: windows_core::PSTR,
    pub pwszVal: windows_core::PWSTR,
    pub punkVal: core::mem::ManuallyDrop<Option<windows_core::IUnknown>>,
    pub pdispVal: core::mem::ManuallyDrop<Option<IDispatch>>,
    pub pStream: core::mem::ManuallyDrop<Option<IStream>>,
    pub pStorage: core::mem::ManuallyDrop<Option<IStorage>>,
    pub pVersionedStream: LPVERSIONEDSTREAM,
    pub parray: LPSAFEARRAY,
    pub cac: CAC,
    pub caub: CAUB,
    pub cai: CAI,
    pub caui: CAUI,
    pub cal: CAL,
    pub caul: CAUL,
    pub cah: CAH,
    pub cauh: CAUH,
    pub caflt: CAFLT,
    pub cadbl: CADBL,
    pub cabool: CABOOL,
    pub cascode: CASCODE,
    pub cacy: CACY,
    pub cadate: CADATE,
    pub cafiletime: CAFILETIME,
    pub cauuid: CACLSID,
    pub caclipdata: CACLIPDATA,
    pub cabstr: CABSTR,
    pub cabstrblob: CABSTRBLOB,
    pub calpstr: CALPSTR,
    pub calpwstr: CALPWSTR,
    pub capropvar: CAPROPVARIANT,
    pub pcVal: *mut CHAR,
    pub pbVal: *mut UCHAR,
    pub piVal: *mut SHORT,
    pub puiVal: *mut USHORT,
    pub plVal: *mut LONG,
    pub pulVal: *mut ULONG,
    pub pintVal: *mut INT,
    pub puintVal: *mut UINT,
    pub pfltVal: *mut FLOAT,
    pub pdblVal: *mut DOUBLE,
    pub pboolVal: *mut VARIANT_BOOL,
    pub pdecVal: *mut DECIMAL,
    pub pscode: *mut SCODE,
    pub pcyVal: *mut CY,
    pub pdate: *mut DATE,
    pub pbstrVal: *mut windows_core::BSTR,
    pub ppunkVal: *mut Option<windows_core::IUnknown>,
    pub ppdispVal: *mut Option<IDispatch>,
    pub pparray: *mut LPSAFEARRAY,
    pub pvarVal: *mut PROPVARIANT,
}
impl Clone for tagPROPVARIANT_0_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for tagPROPVARIANT_0_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
