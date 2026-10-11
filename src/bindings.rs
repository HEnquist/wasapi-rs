#[inline]
pub unsafe fn ActivateAudioInterfaceAsync<P0, P3>(
    deviceinterfacepath: P0,
    riid: *const windows_core::GUID,
    activationparams: Option<*const PROPVARIANT>,
    completionhandler: P3,
) -> windows_core::Result<IActivateAudioInterfaceAsyncOperation>
where
    P0: windows_core::Param<windows_core::PCWSTR>,
    P3: windows_core::Param<IActivateAudioInterfaceCompletionHandler>,
{
    windows_core::link!("mmdevapi.dll" "system" fn ActivateAudioInterfaceAsync(deviceinterfacepath : windows_core::PCWSTR, riid : *const windows_core::GUID, activationparams : *const PROPVARIANT, completionhandler : *mut core::ffi::c_void, activationoperation : *mut *mut core::ffi::c_void) -> windows_core::HRESULT);
    unsafe {
        let mut result__ = core::mem::zeroed();
        ActivateAudioInterfaceAsync(
            deviceinterfacepath.param().abi(),
            riid,
            activationparams.unwrap_or(core::mem::zeroed()) as _,
            completionhandler.param().abi(),
            &mut result__,
        )
        .and_then(|| windows_core::imp::Type::from_abi(result__))
    }
}
#[inline]
pub unsafe fn CloseHandle(hobject: HANDLE) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn CloseHandle(hobject : HANDLE) -> windows_core::BOOL);
    unsafe { CloseHandle(hobject) }
}
#[inline]
pub unsafe fn CoCreateInstance<P1, T>(
    rclsid: *const windows_core::GUID,
    punkouter: P1,
    dwclscontext: u32,
) -> windows_core::Result<T>
where
    P1: windows_core::Param<windows_core::IUnknown>,
    T: windows_core::Interface,
{
    windows_core::link!("ole32.dll" "system" fn CoCreateInstance(rclsid : *const windows_core::GUID, punkouter : *mut core::ffi::c_void, dwclscontext : u32, riid : *const windows_core::GUID, ppv : *mut *mut core::ffi::c_void) -> windows_core::HRESULT);
    let mut result__ = core::ptr::null_mut();
    unsafe {
        CoCreateInstance(
            rclsid,
            punkouter.param().abi(),
            dwclscontext,
            &T::IID,
            &mut result__,
        )
        .and_then(|| windows_core::imp::Type::from_abi(result__))
    }
}
#[inline]
pub unsafe fn CoInitializeEx(
    pvreserved: Option<*const core::ffi::c_void>,
    dwcoinit: u32,
) -> windows_core::HRESULT {
    windows_core::link!("ole32.dll" "system" fn CoInitializeEx(pvreserved : *const core::ffi::c_void, dwcoinit : u32) -> windows_core::HRESULT);
    unsafe { CoInitializeEx(pvreserved.unwrap_or(core::mem::zeroed()) as _, dwcoinit) }
}
#[inline]
pub unsafe fn CoTaskMemFree(pv: *mut core::ffi::c_void) {
    windows_core::link!("ole32.dll" "system" fn CoTaskMemFree(pv : *mut core::ffi::c_void));
    unsafe { CoTaskMemFree(pv as _) }
}
#[inline]
pub unsafe fn CoUninitialize() {
    windows_core::link!("ole32.dll" "system" fn CoUninitialize());
    unsafe { CoUninitialize() }
}
#[inline]
pub unsafe fn CreateEventA<P3>(
    lpeventattributes: Option<*const SECURITY_ATTRIBUTES>,
    bmanualreset: bool,
    binitialstate: bool,
    lpname: P3,
) -> HANDLE
where
    P3: windows_core::Param<windows_core::PCSTR>,
{
    windows_core::link!("kernel32.dll" "system" fn CreateEventA(lpeventattributes : *const SECURITY_ATTRIBUTES, bmanualreset : windows_core::BOOL, binitialstate : windows_core::BOOL, lpname : windows_core::PCSTR) -> HANDLE);
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
    dwdesiredaccess: u32,
    dwsharemode: u32,
    lpsecurityattributes: Option<*const SECURITY_ATTRIBUTES>,
    dwcreationdisposition: u32,
    dwflagsandattributes: u32,
    htemplatefile: Option<HANDLE>,
) -> HANDLE
where
    P0: windows_core::Param<windows_core::PCWSTR>,
{
    windows_core::link!("kernel32.dll" "system" fn CreateFileW(lpfilename : windows_core::PCWSTR, dwdesiredaccess : u32, dwsharemode : u32, lpsecurityattributes : *const SECURITY_ATTRIBUTES, dwcreationdisposition : u32, dwflagsandattributes : u32, htemplatefile : HANDLE) -> HANDLE);
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
    dwiocontrolcode: u32,
    lpinbuffer: Option<*const core::ffi::c_void>,
    ninbuffersize: u32,
    lpoutbuffer: Option<*mut core::ffi::c_void>,
    noutbuffersize: u32,
    lpbytesreturned: Option<*mut u32>,
    lpoverlapped: Option<*mut OVERLAPPED>,
) -> windows_core::BOOL {
    windows_core::link!("kernel32.dll" "system" fn DeviceIoControl(hdevice : HANDLE, dwiocontrolcode : u32, lpinbuffer : *const core::ffi::c_void, ninbuffersize : u32, lpoutbuffer : *mut core::ffi::c_void, noutbuffersize : u32, lpbytesreturned : *mut u32, lpoverlapped : *mut OVERLAPPED) -> windows_core::BOOL);
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
) -> windows_core::Result<windows_core::PWSTR> {
    windows_core::link!("propsys.dll" "system" fn PropVariantToStringAlloc(propvar : *const PROPVARIANT, ppszout : *mut windows_core::PWSTR) -> windows_core::HRESULT);
    unsafe {
        let mut result__ = core::mem::zeroed();
        PropVariantToStringAlloc(propvar, &mut result__).map(|| result__)
    }
}
#[inline]
pub unsafe fn WaitForSingleObject(hhandle: HANDLE, dwmilliseconds: u32) -> u32 {
    windows_core::link!("kernel32.dll" "system" fn WaitForSingleObject(hhandle : HANDLE, dwmilliseconds : u32) -> u32);
    unsafe { WaitForSingleObject(hhandle, dwmilliseconds) }
}
pub const AUDCLNT_BUFFERFLAGS_DATA_DISCONTINUITY: _AUDCLNT_BUFFERFLAGS = 1;
pub const AUDCLNT_BUFFERFLAGS_SILENT: _AUDCLNT_BUFFERFLAGS = 2;
pub const AUDCLNT_BUFFERFLAGS_TIMESTAMP_ERROR: _AUDCLNT_BUFFERFLAGS = 4;
pub type AUDCLNT_SHAREMODE = i32;
pub const AUDCLNT_SHAREMODE_EXCLUSIVE: AUDCLNT_SHAREMODE = 1;
pub const AUDCLNT_SHAREMODE_SHARED: AUDCLNT_SHAREMODE = 0;
pub const AUDCLNT_STREAMFLAGS_AUTOCONVERTPCM: u32 = 2147483648;
pub const AUDCLNT_STREAMFLAGS_EVENTCALLBACK: i32 = 262144;
pub const AUDCLNT_STREAMFLAGS_LOOPBACK: i32 = 131072;
pub const AUDCLNT_STREAMFLAGS_SRC_DEFAULT_QUALITY: i32 = 134217728;
pub type AUDCLNT_STREAMOPTIONS = u32;
pub const AUDCLNT_STREAMOPTIONS_AMBISONICS: AUDCLNT_STREAMOPTIONS = 4;
pub const AUDCLNT_STREAMOPTIONS_MATCH_FORMAT: AUDCLNT_STREAMOPTIONS = 2;
pub const AUDCLNT_STREAMOPTIONS_NONE: AUDCLNT_STREAMOPTIONS = 0;
pub const AUDCLNT_STREAMOPTIONS_POST_VOLUME_LOOPBACK: AUDCLNT_STREAMOPTIONS = 8;
pub const AUDCLNT_STREAMOPTIONS_RAW: AUDCLNT_STREAMOPTIONS = 1;
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
    pub TargetProcessId: u32,
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
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct AUDIO_EFFECT_TYPE_ACOUSTIC_ECHO_CANCELLATION(pub u8);
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
    pub cbSize: u32,
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
    pub cbSize: u32,
    pub pBlobData: *mut u8,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct BSTRBLOB {
    pub cbSize: u32,
    pub pData: *mut u8,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CABOOL {
    pub cElems: u32,
    pub pElems: *mut VARIANT_BOOL,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CABSTR {
    pub cElems: u32,
    pub pElems: *mut windows_core::BSTR,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CABSTRBLOB {
    pub cElems: u32,
    pub pElems: *mut BSTRBLOB,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAC {
    pub cElems: u32,
    pub pElems: *mut i8,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CACLIPDATA {
    pub cElems: u32,
    pub pElems: *mut CLIPDATA,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CACLSID {
    pub cElems: u32,
    pub pElems: *mut windows_core::GUID,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CACY {
    pub cElems: u32,
    pub pElems: *mut CY,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CADATE {
    pub cElems: u32,
    pub pElems: *mut f64,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CADBL {
    pub cElems: u32,
    pub pElems: *mut f64,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAFILETIME {
    pub cElems: u32,
    pub pElems: *mut FILETIME,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAFLT {
    pub cElems: u32,
    pub pElems: *mut f32,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAH {
    pub cElems: u32,
    pub pElems: *mut i64,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAI {
    pub cElems: u32,
    pub pElems: *mut i16,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAL {
    pub cElems: u32,
    pub pElems: *mut i32,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CALPSTR {
    pub cElems: u32,
    pub pElems: *mut windows_core::PSTR,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CALPWSTR {
    pub cElems: u32,
    pub pElems: *mut windows_core::PWSTR,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAPROPVARIANT {
    pub cElems: u32,
    pub pElems: *mut PROPVARIANT,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CASCODE {
    pub cElems: u32,
    pub pElems: *mut SCODE,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAUB {
    pub cElems: u32,
    pub pElems: *mut u8,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAUH {
    pub cElems: u32,
    pub pElems: *mut u64,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAUI {
    pub cElems: u32,
    pub pElems: *mut u16,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CAUL {
    pub cElems: u32,
    pub pElems: *mut u32,
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CLIPDATA {
    pub cbSize: u32,
    pub ulClipFmt: i32,
    pub pClipData: *mut u8,
}
pub const CLSCTX_ALL: i32 = 23;
pub type COINIT = i32;
pub const COINIT_APARTMENTTHREADED: COINIT = 2;
pub const COINIT_MULTITHREADED: COINIT = 0;
#[repr(C)]
#[derive(Clone, Copy)]
pub union CY {
    pub Anonymous: CY_0,
    pub int64: i64,
}
impl Default for CY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct CY_0 {
    pub Lo: u32,
    pub Hi: i32,
}
pub const Connector: PartType = 0;
pub type ConnectorType = i32;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct DECIMAL {
    pub wReserved: u16,
    pub Anonymous: DECIMAL_0,
    pub Hi32: u32,
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
    pub signscale: u16,
}
impl Default for DECIMAL_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_0_0 {
    pub scale: u8,
    pub sign: u8,
}
#[repr(C)]
#[derive(Clone, Copy)]
pub union DECIMAL_1 {
    pub Anonymous: DECIMAL_1_0,
    pub Lo64: u64,
}
impl Default for DECIMAL_1 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct DECIMAL_1_0 {
    pub Lo32: u32,
    pub Mid32: u32,
}
pub type DataFlow = i32;
pub const DisconnectReasonDeviceRemoval: AudioSessionDisconnectReason = 0;
pub const DisconnectReasonExclusiveModeOverride: AudioSessionDisconnectReason = 5;
pub const DisconnectReasonFormatChanged: AudioSessionDisconnectReason = 2;
pub const DisconnectReasonServerShutdown: AudioSessionDisconnectReason = 1;
pub const DisconnectReasonSessionDisconnected: AudioSessionDisconnectReason = 4;
pub const DisconnectReasonSessionLogoff: AudioSessionDisconnectReason = 3;
pub type EDataFlow = i32;
pub const EDataFlow_enum_count: EDataFlow = 3;
pub const ENDPOINT_HARDWARE_SUPPORT_METER: i32 = 4;
pub const ENDPOINT_HARDWARE_SUPPORT_MUTE: i32 = 2;
pub const ENDPOINT_HARDWARE_SUPPORT_VOLUME: i32 = 1;
pub const ERROR_INSUFFICIENT_BUFFER: i32 = 122;
pub const ERROR_MORE_DATA: i32 = 234;
pub type ERole = i32;
pub const ERole_enum_count: ERole = 3;
pub const E_INVALIDARG: windows_core::HRESULT = windows_core::HRESULT(0x80070057_u32 as _);
pub const E_NOINTERFACE: windows_core::HRESULT = windows_core::HRESULT(0x80004002_u32 as _);
pub const FALSE: i32 = 0;
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct FILETIME {
    pub dwLowDateTime: u32,
    pub dwHighDateTime: u32,
}
pub const FILE_SHARE_READ: i32 = 1;
pub const FILE_SHARE_WRITE: i32 = 2;
pub const GENERIC_READ: u32 = 2147483648;
pub const GENERIC_WRITE: i32 = 1073741824;
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct HANDLE(pub *mut core::ffi::c_void);
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
                core::mem::transmute(activatedinterface),
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
windows_core::imp::define_interface!(
    IAudioCaptureClient,
    IAudioCaptureClient_Vtbl,
    0xc8adbd64_e71e_48a0_a4de_185c395cd317
);
windows_core::imp::interface_hierarchy!(IAudioCaptureClient, windows_core::IUnknown);
impl IAudioCaptureClient {
    pub unsafe fn GetBuffer(
        &self,
        ppdata: *mut *mut u8,
        pnumframestoread: *mut u32,
        pdwflags: *mut u32,
        pu64deviceposition: Option<*mut u64>,
        pu64qpcposition: Option<*mut u64>,
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
    pub unsafe fn ReleaseBuffer(&self, numframesread: u32) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).ReleaseBuffer)(
                windows_core::Interface::as_raw(self),
                numframesread,
            )
        }
    }
    pub unsafe fn GetNextPacketSize(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetNextPacketSize)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
pub struct IAudioCaptureClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetBuffer: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut u8,
        *mut u32,
        *mut u32,
        *mut u64,
        *mut u64,
    ) -> windows_core::HRESULT,
    pub ReleaseBuffer:
        unsafe extern "system" fn(*mut core::ffi::c_void, u32) -> windows_core::HRESULT,
    pub GetNextPacketSize:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
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
        streamflags: u32,
        hnsbufferduration: REFERENCE_TIME,
        hnsperiodicity: REFERENCE_TIME,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<*const windows_core::GUID>,
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
    pub unsafe fn GetBufferSize(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetBufferSize)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetStreamLatency(&self) -> windows_core::Result<REFERENCE_TIME> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetStreamLatency)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetCurrentPadding(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetCurrentPadding)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
    pub unsafe fn GetMixFormat(&self) -> windows_core::Result<*mut WAVEFORMATEX> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetMixFormat)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IAudioClient_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Initialize: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        AUDCLNT_SHAREMODE,
        u32,
        REFERENCE_TIME,
        REFERENCE_TIME,
        *const WAVEFORMATEX,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub GetBufferSize:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub GetStreamLatency: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut REFERENCE_TIME,
    ) -> windows_core::HRESULT,
    pub GetCurrentPadding:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
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
        *const windows_core::GUID,
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
    ) -> windows_core::Result<windows_core::BOOL> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).IsOffloadCapable)(
                windows_core::Interface::as_raw(self),
                category,
                &mut result__,
            )
            .map(|| result__)
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
        pdefaultperiodinframes: *mut u32,
        pfundamentalperiodinframes: *mut u32,
        pminperiodinframes: *mut u32,
        pmaxperiodinframes: *mut u32,
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
        pcurrentperiodinframes: *mut u32,
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
        streamflags: u32,
        periodinframes: u32,
        pformat: *const WAVEFORMATEX,
        audiosessionguid: Option<*const windows_core::GUID>,
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
        *mut u32,
        *mut u32,
        *mut u32,
        *mut u32,
    ) -> windows_core::HRESULT,
    pub GetCurrentSharedModeEnginePeriod: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut *mut WAVEFORMATEX,
        *mut u32,
    ) -> windows_core::HRESULT,
    pub InitializeSharedAudioStream: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        u32,
        *const WAVEFORMATEX,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioClock,
    IAudioClock_Vtbl,
    0xcd63314f_3fba_4a1b_812c_ef96358728e7
);
windows_core::imp::interface_hierarchy!(IAudioClock, windows_core::IUnknown);
impl IAudioClock {
    pub unsafe fn GetFrequency(&self) -> windows_core::Result<u64> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetFrequency)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetPosition(
        &self,
        pu64position: *mut u64,
        pu64qpcposition: Option<*mut u64>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).GetPosition)(
                windows_core::Interface::as_raw(self),
                pu64position as _,
                pu64qpcposition.unwrap_or(core::mem::zeroed()) as _,
            )
        }
    }
    pub unsafe fn GetCharacteristics(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetCharacteristics)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
pub struct IAudioClock_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetFrequency:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u64) -> windows_core::HRESULT,
    pub GetPosition: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut u64,
        *mut u64,
    ) -> windows_core::HRESULT,
    pub GetCharacteristics:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
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
        numeffects: *mut u32,
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
        *mut u32,
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
    pub unsafe fn GetPeakValue(&self) -> windows_core::Result<f32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetPeakValue)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetMeteringChannelCount(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetMeteringChannelCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetChannelsPeakValues(&self, u32channelcount: u32) -> windows_core::Result<f32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetChannelsPeakValues)(
                windows_core::Interface::as_raw(self),
                u32channelcount,
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn QueryHardwareSupport(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).QueryHardwareSupport)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
pub struct IAudioMeterInformation_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetPeakValue:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut f32) -> windows_core::HRESULT,
    pub GetMeteringChannelCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub GetChannelsPeakValues:
        unsafe extern "system" fn(*mut core::ffi::c_void, u32, *mut f32) -> windows_core::HRESULT,
    pub QueryHardwareSupport:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioRenderClient,
    IAudioRenderClient_Vtbl,
    0xf294acfc_3146_4483_a7bf_addca7c260e2
);
windows_core::imp::interface_hierarchy!(IAudioRenderClient, windows_core::IUnknown);
impl IAudioRenderClient {
    pub unsafe fn GetBuffer(&self, numframesrequested: u32) -> windows_core::Result<*mut u8> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetBuffer)(
                windows_core::Interface::as_raw(self),
                numframesrequested,
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn ReleaseBuffer(
        &self,
        numframeswritten: u32,
        dwflags: u32,
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
        u32,
        *mut *mut u8,
    ) -> windows_core::HRESULT,
    pub ReleaseBuffer:
        unsafe extern "system" fn(*mut core::ffi::c_void, u32, u32) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IAudioSessionControl,
    IAudioSessionControl_Vtbl,
    0xf4b1a599_7266_4319_a8ca_e70acb11e8cd
);
windows_core::imp::interface_hierarchy!(IAudioSessionControl, windows_core::IUnknown);
impl IAudioSessionControl {
    pub unsafe fn GetState(&self) -> windows_core::Result<AudioSessionState> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetState)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetDisplayName(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDisplayName)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn SetDisplayName<P0>(
        &self,
        value: P0,
        eventcontext: *const windows_core::GUID,
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
    pub unsafe fn GetIconPath(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetIconPath)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn SetIconPath<P0>(
        &self,
        value: P0,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::HRESULT
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
    pub unsafe fn GetGroupingParam(&self) -> windows_core::Result<windows_core::GUID> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetGroupingParam)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn SetGroupingParam(
        &self,
        r#override: *const windows_core::GUID,
        eventcontext: *const windows_core::GUID,
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
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub GetIconPath: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub SetIconPath: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub GetGroupingParam: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub SetGroupingParam: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        *const windows_core::GUID,
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
    pub unsafe fn GetSessionIdentifier(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSessionIdentifier)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetSessionInstanceIdentifier(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSessionInstanceIdentifier)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetProcessId(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetProcessId)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
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
    pub unsafe fn GetCount(&self) -> windows_core::Result<i32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetSession(
        &self,
        sessioncount: i32,
    ) -> windows_core::Result<IAudioSessionControl> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSession)(
                windows_core::Interface::as_raw(self),
                sessioncount,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
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
        eventcontext: *const windows_core::GUID,
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
        eventcontext: *const windows_core::GUID,
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
        eventcontext: *const windows_core::GUID,
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
        newchannelvolumearray: &[f32],
        changedchannel: u32,
        eventcontext: *const windows_core::GUID,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).OnChannelVolumeChanged)(
                windows_core::Interface::as_raw(self),
                newchannelvolumearray.len().try_into().unwrap(),
                newchannelvolumearray.as_ptr(),
                changedchannel,
                eventcontext,
            )
        }
    }
    pub unsafe fn OnGroupingParamChanged(
        &self,
        newgroupingparam: *const windows_core::GUID,
        eventcontext: *const windows_core::GUID,
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
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub OnIconPathChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        windows_core::PCWSTR,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub OnSimpleVolumeChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        f32,
        windows_core::BOOL,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub OnChannelVolumeChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *const f32,
        u32,
        *const windows_core::GUID,
    ) -> windows_core::HRESULT,
    pub OnGroupingParamChanged: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        *const windows_core::GUID,
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
windows_core::imp::define_interface!(
    IAudioSessionManager,
    IAudioSessionManager_Vtbl,
    0xbfa971f1_4d5e_40bb_935e_967039bfbee4
);
windows_core::imp::interface_hierarchy!(IAudioSessionManager, windows_core::IUnknown);
impl IAudioSessionManager {
    pub unsafe fn GetAudioSessionControl(
        &self,
        audiosessionguid: Option<*const windows_core::GUID>,
        streamflags: u32,
    ) -> windows_core::Result<IAudioSessionControl> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetAudioSessionControl)(
                windows_core::Interface::as_raw(self),
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
                streamflags,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetSimpleAudioVolume(
        &self,
        audiosessionguid: Option<*const windows_core::GUID>,
        streamflags: u32,
    ) -> windows_core::Result<ISimpleAudioVolume> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSimpleAudioVolume)(
                windows_core::Interface::as_raw(self),
                audiosessionguid.unwrap_or(core::mem::zeroed()) as _,
                streamflags,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IAudioSessionManager_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetAudioSessionControl: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        u32,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetSimpleAudioVolume: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        u32,
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
            .and_then(|| windows_core::imp::Type::from_abi(result__))
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
    pub unsafe fn GetType(&self) -> windows_core::Result<ConnectorType> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetType)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetDataFlow(&self) -> windows_core::Result<DataFlow> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDataFlow)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
    pub unsafe fn IsConnected(&self) -> windows_core::Result<windows_core::BOOL> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).IsConnected)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetConnectedTo(&self) -> windows_core::Result<Self> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetConnectedTo)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetConnectorIdConnectedTo(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetConnectorIdConnectedTo)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetDeviceIdConnectedTo(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDeviceIdConnectedTo)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
    pub unsafe fn GetConnectorCount(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetConnectorCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetConnector(&self, nindex: u32) -> windows_core::Result<IConnector> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetConnector)(
                windows_core::Interface::as_raw(self),
                nindex,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetSubunitCount(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSubunitCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetSubunit(&self, nindex: u32) -> windows_core::Result<ISubunit> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSubunit)(
                windows_core::Interface::as_raw(self),
                nindex,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetPartById(&self, nid: u32) -> windows_core::Result<IPart> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetPartById)(
                windows_core::Interface::as_raw(self),
                nid,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetDeviceId(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDeviceId)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IDeviceTopology_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetConnectorCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub GetConnector: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetSubunitCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub GetSubunit: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetPartById: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
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
windows_core::imp::define_interface!(
    IMMDevice,
    IMMDevice_Vtbl,
    0xd666063f_1587_4e43_81f1_b948e807363f
);
windows_core::imp::interface_hierarchy!(IMMDevice, windows_core::IUnknown);
impl IMMDevice {
    pub unsafe fn Activate<T>(
        &self,
        dwclsctx: u32,
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
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn OpenPropertyStore(
        &self,
        stgmaccess: u32,
    ) -> windows_core::Result<IPropertyStore> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).OpenPropertyStore)(
                windows_core::Interface::as_raw(self),
                stgmaccess,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetId(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetId)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetState(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetState)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
}
#[repr(C)]
pub struct IMMDevice_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub Activate: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *const windows_core::GUID,
        u32,
        *const PROPVARIANT,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub OpenPropertyStore: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
        *mut *mut core::ffi::c_void,
    ) -> windows_core::HRESULT,
    pub GetId: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        *mut windows_core::PWSTR,
    ) -> windows_core::HRESULT,
    pub GetState:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
}
windows_core::imp::define_interface!(
    IMMDeviceCollection,
    IMMDeviceCollection_Vtbl,
    0x0bd7a1be_7a1a_44db_8397_cc5392387b5e
);
windows_core::imp::interface_hierarchy!(IMMDeviceCollection, windows_core::IUnknown);
impl IMMDeviceCollection {
    pub unsafe fn GetCount(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn Item(&self, ndevice: u32) -> windows_core::Result<IMMDevice> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).Item)(
                windows_core::Interface::as_raw(self),
                ndevice,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
}
#[repr(C)]
pub struct IMMDeviceCollection_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    pub GetCount:
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub Item: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
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
        dwstatemask: u32,
    ) -> windows_core::Result<IMMDeviceCollection> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).EnumAudioEndpoints)(
                windows_core::Interface::as_raw(self),
                dataflow,
                dwstatemask,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetDefaultAudioEndpoint(
        &self,
        dataflow: EDataFlow,
        role: ERole,
    ) -> windows_core::Result<IMMDevice> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDefaultAudioEndpoint)(
                windows_core::Interface::as_raw(self),
                dataflow,
                role,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetDevice<P0>(&self, pwstrid: P0) -> windows_core::Result<IMMDevice>
    where
        P0: windows_core::Param<windows_core::PCWSTR>,
    {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDevice)(
                windows_core::Interface::as_raw(self),
                pwstrid.param().abi(),
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
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
        u32,
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
    pub unsafe fn GetDataFlow(&self) -> windows_core::Result<EDataFlow> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetDataFlow)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
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
        dwnewstate: u32,
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
        u32,
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
pub const IOCTL_KS_PROPERTY: i32 = 3080195;
windows_core::imp::define_interface!(IPart, IPart_Vtbl, 0xae2de0e4_5bca_4f2d_aa46_5d13f8fdb3a9);
windows_core::imp::interface_hierarchy!(IPart, windows_core::IUnknown);
impl IPart {
    pub unsafe fn GetName(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetName)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetLocalId(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetLocalId)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetGlobalId(&self) -> windows_core::Result<windows_core::PWSTR> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetGlobalId)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetPartType(&self) -> windows_core::Result<PartType> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetPartType)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetSubType(&self) -> windows_core::Result<windows_core::GUID> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetSubType)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetControlInterfaceCount(&self) -> windows_core::Result<u32> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetControlInterfaceCount)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .map(|| result__)
        }
    }
    pub unsafe fn GetControlInterface(
        &self,
        nindex: u32,
    ) -> windows_core::Result<IControlInterface> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetControlInterface)(
                windows_core::Interface::as_raw(self),
                nindex,
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn EnumPartsIncoming(&self) -> windows_core::Result<IPartsList> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).EnumPartsIncoming)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn EnumPartsOutgoing(&self) -> windows_core::Result<IPartsList> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).EnumPartsOutgoing)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn GetTopologyObject(&self) -> windows_core::Result<IDeviceTopology> {
        unsafe {
            let mut result__ = core::mem::zeroed();
            (windows_core::Interface::vtable(self).GetTopologyObject)(
                windows_core::Interface::as_raw(self),
                &mut result__,
            )
            .and_then(|| windows_core::imp::Type::from_abi(result__))
        }
    }
    pub unsafe fn Activate(
        &self,
        dwclscontext: u32,
        refiid: *const windows_core::GUID,
        ppvobject: Option<*mut *mut core::ffi::c_void>,
    ) -> windows_core::HRESULT {
        unsafe {
            (windows_core::Interface::vtable(self).Activate)(
                windows_core::Interface::as_raw(self),
                dwclscontext,
                refiid,
                ppvobject.unwrap_or(core::mem::zeroed()) as _,
            )
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
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
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
        unsafe extern "system" fn(*mut core::ffi::c_void, *mut u32) -> windows_core::HRESULT,
    pub GetControlInterface: unsafe extern "system" fn(
        *mut core::ffi::c_void,
        u32,
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
        u32,
        *const windows_core::GUID,
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
#[repr(C)]
pub struct IPartsList_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetCount: usize,
    GetPart: usize,
}
windows_core::imp::define_interface!(
    IPropertyStore,
    IPropertyStore_Vtbl,
    0x886d8eeb_8cf2_4446_8d02_cdba1dbdcf99
);
windows_core::imp::interface_hierarchy!(IPropertyStore, windows_core::IUnknown);
#[repr(C)]
pub struct IPropertyStore_Vtbl {
    pub base__: windows_core::IUnknown_Vtbl,
    GetCount: usize,
    GetAt: usize,
    GetValue: usize,
    SetValue: usize,
    Commit: usize,
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
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSDATAFORMAT_SUBTYPE_IEEE_FLOAT(pub u8);
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSDATAFORMAT_SUBTYPE_PCM(pub u8);
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSDATAFORMAT_TYPE_AUDIO(pub u8);
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSMULTIPLE_ITEM {
    pub Size: u32,
    pub Count: u32,
}
pub type KSPIN_DATAFLOW = i32;
pub const KSPIN_DATAFLOW_IN: KSPIN_DATAFLOW = 1;
pub const KSPIN_DATAFLOW_OUT: KSPIN_DATAFLOW = 2;
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
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct KSPROPSETID_Pin(pub u8);
pub type LPSAFEARRAY = *mut SAFEARRAY;
pub type LPVERSIONEDSTREAM = *mut VERSIONEDSTREAM;
pub const MMDeviceEnumerator: windows_core::GUID =
    windows_core::GUID::from_u128(0xbcde0395_e52f_467c_8e3d_c4579291692e);
pub const OPEN_EXISTING: i32 = 3;
#[repr(C)]
#[derive(Clone, Copy)]
pub struct OVERLAPPED {
    pub Internal: usize,
    pub InternalHigh: usize,
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
    pub Pointer: *mut core::ffi::c_void,
}
impl Default for OVERLAPPED_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct OVERLAPPED_0_0 {
    pub Offset: u32,
    pub OffsetHigh: u32,
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
    pub pid: u32,
}
#[repr(C)]
pub struct PROPVARIANT {
    pub Anonymous: PROPVARIANT_0,
}
impl Clone for PROPVARIANT {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for PROPVARIANT {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
pub union PROPVARIANT_0 {
    pub Anonymous: core::mem::ManuallyDrop<PROPVARIANT_0_0>,
    pub decVal: DECIMAL,
}
impl Clone for PROPVARIANT_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for PROPVARIANT_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
pub struct PROPVARIANT_0_0 {
    pub vt: VARTYPE,
    pub wReserved1: PROPVAR_PAD1,
    pub wReserved2: PROPVAR_PAD2,
    pub wReserved3: PROPVAR_PAD3,
    pub Anonymous: PROPVARIANT_0_0_0,
}
impl Clone for PROPVARIANT_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for PROPVARIANT_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
pub union PROPVARIANT_0_0_0 {
    pub cVal: i8,
    pub bVal: u8,
    pub iVal: i16,
    pub uiVal: u16,
    pub lVal: i32,
    pub ulVal: u32,
    pub intVal: i32,
    pub uintVal: u32,
    pub hVal: i64,
    pub uhVal: u64,
    pub fltVal: f32,
    pub dblVal: f64,
    pub boolVal: VARIANT_BOOL,
    pub __OBSOLETE__VARIANT_BOOL: VARIANT_BOOL,
    pub scode: SCODE,
    pub cyVal: CY,
    pub date: f64,
    pub filetime: FILETIME,
    pub puuid: *mut windows_core::GUID,
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
    pub pcVal: *mut i8,
    pub pbVal: *mut u8,
    pub piVal: *mut i16,
    pub puiVal: *mut u16,
    pub plVal: *mut i32,
    pub pulVal: *mut u32,
    pub pintVal: *mut i32,
    pub puintVal: *mut u32,
    pub pfltVal: *mut f32,
    pub pdblVal: *mut f64,
    pub pboolVal: *mut VARIANT_BOOL,
    pub pdecVal: *mut DECIMAL,
    pub pscode: *mut SCODE,
    pub pcyVal: *mut CY,
    pub pdate: *mut f64,
    pub pbstrVal: *mut windows_core::BSTR,
    pub ppunkVal: *mut Option<windows_core::IUnknown>,
    pub ppdispVal: *mut Option<IDispatch>,
    pub pparray: *mut LPSAFEARRAY,
    pub pvarVal: *mut PROPVARIANT,
}
impl Clone for PROPVARIANT_0_0_0 {
    fn clone(&self) -> Self {
        unsafe { core::mem::transmute_copy(self) }
    }
}
impl Default for PROPVARIANT_0_0_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct PROPVAR_PAD1(pub u16);
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct PROPVAR_PAD2(pub u16);
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct PROPVAR_PAD3(pub u16);
pub type PartType = i32;
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct REFERENCE_TIME(pub i64);
#[repr(C)]
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct SAFEARRAY {
    pub cDims: u16,
    pub fFeatures: u16,
    pub cbElements: u32,
    pub cLocks: u32,
    pub pvData: *mut core::ffi::c_void,
    pub rgsabound: [SAFEARRAYBOUND; 1],
}
impl Default for SAFEARRAY {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct SAFEARRAYBOUND {
    pub cElements: u32,
    pub lLbound: i32,
}
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct SCODE(pub i32);
#[repr(C)]
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub struct SECURITY_ATTRIBUTES {
    pub nLength: u32,
    pub lpSecurityDescriptor: *mut core::ffi::c_void,
    pub bInheritHandle: windows_core::BOOL,
}
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
pub type VARENUM = i32;
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct VARIANT_BOOL(pub i16);
#[repr(transparent)]
#[derive(Clone, Copy, Debug, PartialEq, Eq, Default)]
pub struct VARTYPE(pub u16);
#[repr(C)]
#[derive(Clone, Debug, Default, Eq, PartialEq)]
pub struct VERSIONEDSTREAM {
    pub guidVersion: windows_core::GUID,
    pub pStream: core::mem::ManuallyDrop<Option<IStream>>,
}
pub const VIRTUAL_AUDIO_DEVICE_PROCESS_LOOPBACK: windows_core::PCWSTR =
    windows_core::w!("VAD\\Process_Loopback");
pub const VT_BLOB: VARENUM = 65;
pub const WAIT_OBJECT_0: i32 = 0;
#[repr(C, packed(1))]
#[derive(Clone, Copy, Default)]
pub struct WAVEFORMATEX {
    pub wFormatTag: u16,
    pub nChannels: u16,
    pub nSamplesPerSec: u32,
    pub nAvgBytesPerSec: u32,
    pub nBlockAlign: u16,
    pub wBitsPerSample: u16,
    pub cbSize: u16,
}
#[repr(C, packed(1))]
#[derive(Clone, Copy)]
pub struct WAVEFORMATEXTENSIBLE {
    pub Format: WAVEFORMATEX,
    pub Samples: WAVEFORMATEXTENSIBLE_0,
    pub dwChannelMask: u32,
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
    pub wValidBitsPerSample: u16,
    pub wSamplesPerBlock: u16,
    pub wReserved: u16,
}
impl Default for WAVEFORMATEXTENSIBLE_0 {
    fn default() -> Self {
        unsafe { core::mem::zeroed() }
    }
}
pub const WAVE_FORMAT_EXTENSIBLE: i32 = 65534;
pub const WAVE_FORMAT_IEEE_FLOAT: i32 = 3;
pub const WAVE_FORMAT_PCM: i32 = 1;
pub type _AUDCLNT_BUFFERFLAGS = i32;
pub const eAll: EDataFlow = 2;
pub const eCapture: EDataFlow = 1;
pub const eCommunications: ERole = 2;
pub const eConsole: ERole = 0;
pub const eMultimedia: ERole = 1;
pub const eRender: EDataFlow = 0;
