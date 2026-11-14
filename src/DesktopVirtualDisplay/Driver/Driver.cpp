// Copyright (c) Microsoft Corporation.
// Copyright (c) 2026 Reddysagar Chenchu.
// Derived from Microsoft's IndirectDisplay sample. See ../LICENSE-MS-PL.txt and NOTICE.md.
#define NOMINMAX
#include <windows.h>
#include <wudfwdm.h>
#include <wdf.h>
#include <iddcx.h>
#include <dxgi1_5.h>
#include <d3d11.h>
#include <avrt.h>
#include <wrl.h>
#include <atomic>
#include <memory>
#include <new>
#include "../Shared/Protocol.h"

using Microsoft::WRL::ComPtr;
using namespace desktop_vdd;

namespace {
struct State {
    std::atomic_uint32_t flags{0};
    std::atomic_int32_t error{0};
    std::atomic_uint64_t frames{0};
    bool initialized = false; // WDF serializes device power callbacks.
    void Set(Flag f, bool value) noexcept {
        if (value) flags.fetch_or(f); else flags.fetch_and(~static_cast<std::uint32_t>(f));
    }
};
struct StateContext { std::shared_ptr<State>* state; };
WDF_DECLARE_CONTEXT_TYPE_WITH_NAME(StateContext, GetStateContext);

// A processor owns the swapchain after the assign callback returns success.
// It checks termination even under continuous frame delivery, and joins before destruction.
class Processor {
    IDDCX_SWAPCHAIN chain_;
    HANDLE available_;
    HANDLE stop_ = nullptr;
    HANDLE thread_ = nullptr;
    ComPtr<ID3D11Device> device_;
    std::shared_ptr<State> state_;

    static DWORD WINAPI ThreadMain(void* context) {
        static_cast<Processor*>(context)->Run();
        return 0;
    }
    void Run() noexcept {
        DWORD task = 0;
        HANDLE priority = AvSetMmThreadCharacteristicsW(L"Distribution", &task);
        const HRESULT result = Consume();
        if (FAILED(result)) state_->error.store(result);
        state_->Set(SwapChainActive, false);
        // Release the chain on both expected unassignment and DirectX failure.
        WdfObjectDelete(chain_);
        chain_ = nullptr;
        if (priority) AvRevertMmThreadCharacteristics(priority);
    }
    HRESULT Consume() noexcept {
        ComPtr<IDXGIDevice> dxgi;
        HRESULT hr = device_.As(&dxgi);
        if (FAILED(hr)) return hr;
        IDARG_IN_SWAPCHAINSETDEVICE input{};
        input.pDevice = dxgi.Get();
        hr = IddCxSwapChainSetDevice(chain_, &input);
        if (FAILED(hr)) return hr;
        state_->frames.store(0);
        state_->Set(SwapChainActive, true);
        for (;;) {
            if (WaitForSingleObject(stop_, 0) == WAIT_OBJECT_0) return S_OK;
            IDARG_OUT_RELEASEANDACQUIREBUFFER buffer{};
            hr = IddCxSwapChainReleaseAndAcquireBuffer(chain_, &buffer);
            if (hr == E_PENDING) {
                HANDLE handles[] = {stop_, available_};
                DWORD wait = WaitForMultipleObjects(2, handles, FALSE, INFINITE);
                if (wait == WAIT_OBJECT_0) return S_OK;
                if (wait == WAIT_OBJECT_0 + 1) continue;
                return HRESULT_FROM_WIN32(GetLastError());
            }
            if (FAILED(hr)) return hr;
            // A virtual sink consumes frames without encoding or copying their pixels.
            // Screenshots are a separate user-session operation, not a driver operation.
            ComPtr<IDXGIResource> surface;
            surface.Attach(buffer.MetaData.pSurface);
            surface.Reset();
            hr = IddCxSwapChainFinishedProcessingFrame(chain_);
            if (FAILED(hr)) return hr;
            state_->frames.fetch_add(1);
        }
    }
public:
    Processor(IDDCX_SWAPCHAIN chain, HANDLE available, std::shared_ptr<State> state)
        : chain_(chain), available_(available), state_(std::move(state)) {}
    HRESULT Start(LUID adapterLuid) noexcept {
        ComPtr<IDXGIFactory5> factory;
        HRESULT hr = CreateDXGIFactory2(0, IID_PPV_ARGS(&factory));
        if (FAILED(hr)) return hr;
        ComPtr<IDXGIAdapter1> adapter;
        hr = factory->EnumAdapterByLuid(adapterLuid, IID_PPV_ARGS(&adapter));
        if (FAILED(hr)) return hr;
        hr = D3D11CreateDevice(adapter.Get(), D3D_DRIVER_TYPE_UNKNOWN, nullptr,
            D3D11_CREATE_DEVICE_BGRA_SUPPORT, nullptr, 0, D3D11_SDK_VERSION,
            &device_, nullptr, nullptr);
        if (FAILED(hr)) return hr;
        stop_ = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (!stop_) return HRESULT_FROM_WIN32(GetLastError());
        thread_ = CreateThread(nullptr, 0, ThreadMain, this, 0, nullptr);
        if (!thread_) return HRESULT_FROM_WIN32(GetLastError());
        return S_OK;
    }
    ~Processor() {
        if (stop_) SetEvent(stop_);
        if (thread_) { WaitForSingleObject(thread_, INFINITE); CloseHandle(thread_); }
        if (stop_) CloseHandle(stop_);
    }
};
struct MonitorContext {
    std::shared_ptr<State>* state;
    Processor* processor;
};
WDF_DECLARE_CONTEXT_TYPE_WITH_NAME(MonitorContext, GetMonitorContext);

void FillSignal(DISPLAYCONFIG_VIDEO_SIGNAL_INFO& signal, bool monitor) noexcept {
    signal.activeSize = {Width, Height};
    // CEA 1080p60 timing: 2200 x 1125 total pixels, 148.5 MHz pixel clock.
    signal.totalSize = {2200, 1125};
    signal.pixelRate = 148500000;
    signal.hSyncFreq = {67500, 1};
    signal.vSyncFreq = {RefreshRate, 1};
    signal.AdditionalSignalInfo.videoStandard = 255;
    signal.AdditionalSignalInfo.vSyncFreqDivider = monitor ? 0 : 1;
    signal.scanLineOrdering = DISPLAYCONFIG_SCANLINE_ORDERING_PROGRESSIVE;
}
void DeleteState(WDFOBJECT object) {
    auto context = GetStateContext(object);
    delete context->state;
    context->state = nullptr;
}
}

extern "C" DRIVER_INITIALIZE DriverEntry;
EVT_WDF_DRIVER_DEVICE_ADD DeviceAdd;
EVT_WDF_DEVICE_D0_ENTRY DeviceD0Entry;
EVT_WDF_DEVICE_D0_EXIT DeviceD0Exit;
EVT_IDD_CX_ADAPTER_INIT_FINISHED AdapterFinished;
EVT_IDD_CX_ADAPTER_COMMIT_MODES CommitModes;
EVT_IDD_CX_PARSE_MONITOR_DESCRIPTION ParseDescription;
EVT_IDD_CX_MONITOR_GET_DEFAULT_DESCRIPTION_MODES DefaultModes;
EVT_IDD_CX_MONITOR_QUERY_TARGET_MODES TargetModes;
EVT_IDD_CX_MONITOR_ASSIGN_SWAPCHAIN AssignSwapChain;
EVT_IDD_CX_MONITOR_UNASSIGN_SWAPCHAIN UnassignSwapChain;
EVT_IDD_CX_DEVICE_IO_CONTROL DeviceIoControl;

extern "C" BOOL WINAPI DllMain(HINSTANCE, UINT, LPVOID) { return TRUE; }

_Use_decl_annotations_
extern "C" NTSTATUS DriverEntry(PDRIVER_OBJECT driver, PUNICODE_STRING registryPath) {
    WDF_DRIVER_CONFIG config;
    WDF_DRIVER_CONFIG_INIT(&config, DeviceAdd);
    return WdfDriverCreate(driver, registryPath, WDF_NO_OBJECT_ATTRIBUTES, &config, WDF_NO_HANDLE);
}

_Use_decl_annotations_
NTSTATUS DeviceAdd(WDFDRIVER driver, PWDFDEVICE_INIT init) {
    UNREFERENCED_PARAMETER(driver);
    WDF_PNPPOWER_EVENT_CALLBACKS power;
    WDF_PNPPOWER_EVENT_CALLBACKS_INIT(&power);
    power.EvtDeviceD0Entry = DeviceD0Entry;
    power.EvtDeviceD0Exit = DeviceD0Exit;
    WdfDeviceInitSetPnpPowerEventCallbacks(init, &power);

    IDD_CX_CLIENT_CONFIG config;
    IDD_CX_CLIENT_CONFIG_INIT(&config); // Dynamic size for the IddCx 1.3 minimum.
    config.EvtIddCxAdapterInitFinished = AdapterFinished;
    config.EvtIddCxAdapterCommitModes = CommitModes;
    config.EvtIddCxParseMonitorDescription = ParseDescription;
    config.EvtIddCxMonitorGetDefaultDescriptionModes = DefaultModes;
    config.EvtIddCxMonitorQueryTargetModes = TargetModes;
    config.EvtIddCxMonitorAssignSwapChain = AssignSwapChain;
    config.EvtIddCxMonitorUnassignSwapChain = UnassignSwapChain;
    config.EvtIddCxDeviceIoControl = DeviceIoControl;
    NTSTATUS status = IddCxDeviceInitConfig(init, &config);
    if (!NT_SUCCESS(status)) return status;

    WDF_OBJECT_ATTRIBUTES attributes;
    WDF_OBJECT_ATTRIBUTES_INIT_CONTEXT_TYPE(&attributes, StateContext);
    attributes.EvtCleanupCallback = DeleteState;
    WDFDEVICE device;
    status = WdfDeviceCreate(&init, &attributes, &device);
    if (!NT_SUCCESS(status)) return status;
    try { GetStateContext(device)->state = new std::shared_ptr<State>(std::make_shared<State>()); }
    catch (const std::bad_alloc&) { return STATUS_INSUFFICIENT_RESOURCES; }
    status = IddCxDeviceInitialize(device);
    if (!NT_SUCCESS(status)) return status;
    return WdfDeviceCreateDeviceInterface(device, &InterfaceId, nullptr);
}

_Use_decl_annotations_
NTSTATUS DeviceD0Entry(WDFDEVICE device, WDF_POWER_DEVICE_STATE previous) {
    UNREFERENCED_PARAMETER(previous);
    auto state = *GetStateContext(device)->state;
    state->Set(PoweredOn, true);
    if (state->initialized) return STATUS_SUCCESS;
    state->initialized = true;

    IDDCX_ADAPTER_CAPS caps{};
    caps.Size = sizeof(caps);
    caps.MaxMonitorsSupported = 1;
    caps.EndPointDiagnostics.Size = sizeof(caps.EndPointDiagnostics);
    caps.EndPointDiagnostics.GammaSupport = IDDCX_FEATURE_IMPLEMENTATION_NONE;
    caps.EndPointDiagnostics.TransmissionType = IDDCX_TRANSMISSION_TYPE_WIRED_OTHER;
    caps.EndPointDiagnostics.pEndPointFriendlyName = L"DesktopRuntime Virtual Display Prototype";
    caps.EndPointDiagnostics.pEndPointManufacturerName = L"Reddysagar Chenchu";
    caps.EndPointDiagnostics.pEndPointModelName = L"Console 1080p Prototype";
    IDDCX_ENDPOINT_VERSION version{};
    version.Size = sizeof(version);
    version.MajorVer = 0;
    version.MinorVer = 1;
    caps.EndPointDiagnostics.pFirmwareVersion = &version;
    caps.EndPointDiagnostics.pHardwareVersion = &version;

    WDF_OBJECT_ATTRIBUTES attributes;
    WDF_OBJECT_ATTRIBUTES_INIT_CONTEXT_TYPE(&attributes, StateContext);
    attributes.EvtCleanupCallback = DeleteState;
    IDARG_IN_ADAPTER_INIT input{};
    input.WdfDevice = device;
    input.pCaps = &caps;
    input.ObjectAttributes = &attributes;
    IDARG_OUT_ADAPTER_INIT output{};
    NTSTATUS status = IddCxAdapterInitAsync(&input, &output);
    if (!NT_SUCCESS(status)) { state->error.store(status); return status; }
    try { GetStateContext(output.AdapterObject)->state = new std::shared_ptr<State>(state); }
    catch (const std::bad_alloc&) { return STATUS_INSUFFICIENT_RESOURCES; }
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS DeviceD0Exit(WDFDEVICE device, WDF_POWER_DEVICE_STATE target) {
    UNREFERENCED_PARAMETER(target);
    (*GetStateContext(device)->state)->Set(PoweredOn, false);
    // IddCx owns adapter/monitor teardown; unassign and object cleanup join our frame thread.
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS AdapterFinished(IDDCX_ADAPTER adapter, const IDARG_IN_ADAPTER_INIT_FINISHED* input) {
    auto context = GetStateContext(adapter);
    if (!context->state) return STATUS_INSUFFICIENT_RESOURCES;
    auto state = *context->state;
    if (!NT_SUCCESS(input->AdapterInitStatus)) {
        state->error.store(input->AdapterInitStatus);
        return STATUS_SUCCESS;
    }
    state->Set(AdapterReady, true);
    IDDCX_MONITOR_INFO info{};
    info.Size = sizeof(info);
    info.MonitorType = DISPLAYCONFIG_OUTPUT_TECHNOLOGY_HDMI;
    info.ConnectorIndex = 0;
    info.MonitorContainerId = MonitorContainerId;
    info.MonitorDescription.Size = sizeof(info.MonitorDescription);
    info.MonitorDescription.Type = IDDCX_MONITOR_DESCRIPTION_TYPE_EDID;
    // Descriptorless virtual sink: DefaultModes supplies its single supported mode.
    info.MonitorDescription.DataSize = 0;
    info.MonitorDescription.pData = nullptr;

    WDF_OBJECT_ATTRIBUTES attributes;
    WDF_OBJECT_ATTRIBUTES_INIT_CONTEXT_TYPE(&attributes, MonitorContext);
    attributes.EvtCleanupCallback = [](WDFOBJECT object) {
        auto monitor = GetMonitorContext(object);
        delete monitor->processor;
        monitor->processor = nullptr;
        if (monitor->state) (*monitor->state)->Set(MonitorArrived, false);
        delete monitor->state;
        monitor->state = nullptr;
    };
    IDARG_IN_MONITORCREATE create{};
    create.ObjectAttributes = &attributes;
    create.pMonitorInfo = &info;
    IDARG_OUT_MONITORCREATE output{};
    NTSTATUS status = IddCxMonitorCreate(adapter, &create, &output);
    if (!NT_SUCCESS(status)) { state->error.store(status); return status; }
    auto monitor = GetMonitorContext(output.MonitorObject);
    try { monitor->state = new std::shared_ptr<State>(state); }
    catch (const std::bad_alloc&) {
        WdfObjectDelete(output.MonitorObject);
        state->error.store(STATUS_INSUFFICIENT_RESOURCES);
        return STATUS_INSUFFICIENT_RESOURCES;
    }
    IDARG_OUT_MONITORARRIVAL arrival{};
    status = IddCxMonitorArrival(output.MonitorObject, &arrival);
    if (!NT_SUCCESS(status)) { state->error.store(status); WdfObjectDelete(output.MonitorObject); }
    else state->Set(MonitorArrived, true);
    return status;
}

_Use_decl_annotations_
NTSTATUS CommitModes(IDDCX_ADAPTER adapter, const IDARG_IN_COMMITMODES* input) {
    auto state = *GetStateContext(adapter)->state;
    bool active = false;
    for (UINT index = 0; index < input->PathCount; ++index) {
        const auto& path = input->pPaths[index];
        if ((path.Flags & IDDCX_PATH_FLAGS_ACTIVE) != 0) {
            const auto& signal = path.TargetVideoSignalInfo;
            if (signal.activeSize.cx != Width || signal.activeSize.cy != Height ||
                signal.vSyncFreq.Denominator == 0 ||
                static_cast<UINT64>(signal.vSyncFreq.Numerator) !=
                    static_cast<UINT64>(RefreshRate) * signal.vSyncFreq.Denominator) {
                state->Set(PathActive, false);
                state->error.store(STATUS_INVALID_PARAMETER);
                return STATUS_INVALID_PARAMETER;
            }
            active = true;
        }
    }
    state->Set(PathActive, active);
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS ParseDescription(const IDARG_IN_PARSEMONITORDESCRIPTION* input, IDARG_OUT_PARSEMONITORDESCRIPTION* output) {
    UNREFERENCED_PARAMETER(input);
    output->MonitorModeBufferOutputCount = 0;
    output->PreferredMonitorModeIdx = 0;
    // This driver supplies no EDID. DefaultModes is the only supported mode source.
    return STATUS_INVALID_PARAMETER;
}

_Use_decl_annotations_
NTSTATUS DefaultModes(IDDCX_MONITOR monitor, const IDARG_IN_GETDEFAULTDESCRIPTIONMODES* input,
    IDARG_OUT_GETDEFAULTDESCRIPTIONMODES* output) {
    UNREFERENCED_PARAMETER(monitor);
    output->DefaultMonitorModeBufferOutputCount = 1;
    output->PreferredMonitorModeIdx = 0;
    if (input->DefaultMonitorModeBufferInputCount == 0) return STATUS_SUCCESS;
    if (!input->pDefaultMonitorModes) return STATUS_INVALID_PARAMETER;
    IDDCX_MONITOR_MODE mode{};
    mode.Size = sizeof(mode);
    mode.Origin = IDDCX_MONITOR_MODE_ORIGIN_DRIVER;
    FillSignal(mode.MonitorVideoSignalInfo, true);
    input->pDefaultMonitorModes[0] = mode;
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS TargetModes(IDDCX_MONITOR monitor, const IDARG_IN_QUERYTARGETMODES* input,
    IDARG_OUT_QUERYTARGETMODES* output) {
    UNREFERENCED_PARAMETER(monitor);
    output->TargetModeBufferOutputCount = 1;
    if (input->TargetModeBufferInputCount == 0) return STATUS_SUCCESS;
    if (!input->pTargetModes) return STATUS_INVALID_PARAMETER;
    IDDCX_TARGET_MODE mode{};
    mode.Size = sizeof(mode);
    FillSignal(mode.TargetVideoSignalInfo.targetVideoSignalInfo, false);
    input->pTargetModes[0] = mode;
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS AssignSwapChain(IDDCX_MONITOR monitorObject, const IDARG_IN_SETSWAPCHAIN* input) {
    auto monitor = GetMonitorContext(monitorObject);
    delete monitor->processor;
    monitor->processor = nullptr;
    auto processor = new (std::nothrow) Processor(input->hSwapChain, input->hNextSurfaceAvailable, *monitor->state);
    if (!processor) return STATUS_INSUFFICIENT_RESOURCES;
    HRESULT hr = processor->Start(input->RenderAdapterLuid);
    if (FAILED(hr)) {
        (*monitor->state)->error.store(hr);
        delete processor;
        // Ownership was not accepted: the OS releases the chain on failure.
        return STATUS_GRAPHICS_INDIRECT_DISPLAY_ABANDON_SWAPCHAIN;
    }
    monitor->processor = processor;
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
NTSTATUS UnassignSwapChain(IDDCX_MONITOR monitorObject) {
    auto monitor = GetMonitorContext(monitorObject);
    delete monitor->processor;
    monitor->processor = nullptr;
    (*monitor->state)->Set(SwapChainActive, false);
    return STATUS_SUCCESS;
}

_Use_decl_annotations_
VOID DeviceIoControl(WDFDEVICE device, WDFREQUEST request, size_t outputSize,
    size_t inputSize, ULONG code) {
    if (code != QueryStatus) { WdfRequestComplete(request, STATUS_INVALID_DEVICE_REQUEST); return; }
    if (inputSize != 0) { WdfRequestComplete(request, STATUS_INVALID_PARAMETER); return; }
    if (outputSize < sizeof(Status)) { WdfRequestComplete(request, STATUS_BUFFER_TOO_SMALL); return; }
    Status* output = nullptr;
    NTSTATUS status = WdfRequestRetrieveOutputBuffer(request, sizeof(Status), reinterpret_cast<void**>(&output), nullptr);
    if (!NT_SUCCESS(status)) { WdfRequestComplete(request, status); return; }
    auto state = *GetStateContext(device)->state;
    *output = {sizeof(Status), ProtocolVersion, state->flags.load(), state->error.load(),
        Width, Height, RefreshRate, 0, state->frames.load()};
    WdfRequestCompleteWithInformation(request, STATUS_SUCCESS, sizeof(Status));
}
