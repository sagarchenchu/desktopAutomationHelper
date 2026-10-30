// Copyright (c) 2026 Reddysagar Chenchu.
#define NOMINMAX
#include <windows.h>
#include <swdevice.h>
#include <setupapi.h>
#include <cfgmgr32.h>
#include <sddl.h>
#include <iostream>
#include <memory>
#include <string>
#include <vector>
#include "../Shared/Protocol.h"
#include "Options.h"

using namespace desktop_vdd;
namespace {
struct Handle {
    HANDLE value = nullptr;
    ~Handle() { if (value && value != INVALID_HANDLE_VALUE) CloseHandle(value); }
    Handle() = default;
    Handle(const Handle&) = delete;
    Handle& operator=(const Handle&) = delete;
};
struct DeviceSet {
    HDEVINFO value = INVALID_HANDLE_VALUE;
    ~DeviceSet() { if (value != INVALID_HANDLE_VALUE) SetupDiDestroyDeviceInfoList(value); }
};
struct SoftwareDevice {
    HSWDEVICE value = nullptr;
    void Close() { if (value) { SwDeviceClose(value); value = nullptr; } }
    ~SoftwareDevice() { Close(); }
};
struct CallbackState {
    HANDLE completed = nullptr; // Kept alive until process exit, including late callbacks.
    HRESULT result = E_PENDING;
    wchar_t instance[MAX_DEVICE_ID_LEN]{};
};
HANDLE cancelEvent = nullptr;
BOOL WINAPI OnConsoleEvent(DWORD event) {
    if (event == CTRL_C_EVENT || event == CTRL_BREAK_EVENT) {
        if (cancelEvent) SetEvent(cancelEvent);
        return TRUE;
    }
    return FALSE; // OS closes the owned software-device handle on process termination.
}
void CALLBACK Created(HSWDEVICE, HRESULT result, void* context, PCWSTR instance) {
    auto state = static_cast<CallbackState*>(context);
    state->result = result;
    if (instance) wcsncpy_s(state->instance, instance, _TRUNCATE);
    SetEvent(state->completed);
}
void Error(const char* event, unsigned long code) {
    std::cout << "{\"schemaVersion\":1,\"event\":\"" << event << "\",\"code\":" << code << "}" << std::endl;
}
void PrintStatus(const Status& status, const char* event) {
    DWORD session = 0;
    ProcessIdToSessionId(GetCurrentProcessId(), &session);
    std::cout << "{\"schemaVersion\":1,\"event\":\"" << event << "\",\"pipelineReady\":"
        << (IsPipelineReady(status) ? "true" : "false") << ",\"flags\":" << status.flags
        << ",\"lastDriverError\":" << status.lastError << ",\"completedFrames\":" << status.completedFrames
        << ",\"width\":" << status.width << ",\"height\":" << status.height
        << ",\"refreshRate\":" << status.refreshRate << ",\"callingSessionId\":" << session
        << ",\"activeConsoleSessionId\":" << WTSGetActiveConsoleSessionId()
        << ",\"inputVerified\":false,\"captureVerified\":false,\"automationVerified\":false}" << std::endl;
}
bool IsAdministrator() {
    SID_IDENTIFIER_AUTHORITY authority = SECURITY_NT_AUTHORITY;
    PSID group = nullptr;
    if (!AllocateAndInitializeSid(&authority, 2, SECURITY_BUILTIN_DOMAIN_RID,
        DOMAIN_ALIAS_RID_ADMINS, 0, 0, 0, 0, 0, 0, &group)) return false;
    BOOL member = FALSE;
    CheckTokenMembership(nullptr, group, &member);
    FreeSid(group);
    return member != FALSE;
}
// Resolve the interface by GUID and (for hold) the actual PnP ID from the callback.
// Never assume that a friendly name, interface enumeration order or derived PnP ID is unique.
bool ReadStatus(const wchar_t* ownedInstance, Status& status, DWORD& error) {
    DeviceSet set;
    set.value = SetupDiGetClassDevsW(&InterfaceId, nullptr, nullptr, DIGCF_PRESENT | DIGCF_DEVICEINTERFACE);
    if (set.value == INVALID_HANDLE_VALUE) { error = GetLastError(); return false; }
    bool found = false;
    for (DWORD index = 0;; ++index) {
        SP_DEVICE_INTERFACE_DATA item{};
        item.cbSize = sizeof(item);
        if (!SetupDiEnumDeviceInterfaces(set.value, nullptr, &InterfaceId, index, &item)) {
            error = GetLastError();
            if (error != ERROR_NO_MORE_ITEMS) return false;
            break;
        }
        DWORD required = 0;
        SetupDiGetDeviceInterfaceDetailW(set.value, &item, nullptr, 0, &required, nullptr);
        if (required < sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W) || required > 65536) { error = ERROR_INVALID_DATA; return false; }
        std::vector<BYTE> bytes(required);
        auto detail = reinterpret_cast<SP_DEVICE_INTERFACE_DETAIL_DATA_W*>(bytes.data());
        detail->cbSize = sizeof(SP_DEVICE_INTERFACE_DETAIL_DATA_W);
        SP_DEVINFO_DATA device{};
        device.cbSize = sizeof(device);
        if (!SetupDiGetDeviceInterfaceDetailW(set.value, &item, detail, required, nullptr, &device)) { error = GetLastError(); return false; }
        wchar_t instance[MAX_DEVICE_ID_LEN]{};
        if (!SetupDiGetDeviceInstanceIdW(set.value, &device, instance, MAX_DEVICE_ID_LEN, nullptr)) { error = GetLastError(); return false; }
        if (ownedInstance && _wcsicmp(ownedInstance, instance) != 0) continue;
        if (found) { error = ERROR_DUP_NAME; return false; }
        Handle file;
        file.value = CreateFileW(detail->DevicePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE,
            nullptr, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, nullptr);
        if (file.value == INVALID_HANDLE_VALUE) { error = GetLastError(); return false; }
        Handle done;
        done.value = CreateEventW(nullptr, TRUE, FALSE, nullptr);
        if (!done.value) { error = GetLastError(); return false; }
        OVERLAPPED operation{};
        operation.hEvent = done.value;
        DWORD returned = 0;
        Status candidate{};
        if (!DeviceIoControl(file.value, QueryStatus, nullptr, 0, &candidate, sizeof(candidate), &returned, &operation)) {
            error = GetLastError();
            if (error != ERROR_IO_PENDING) return false;
            if (WaitForSingleObject(done.value, 2000) != WAIT_OBJECT_0) {
                CancelIoEx(file.value, &operation);
                // Keep the OVERLAPPED and output buffer alive until cancellation completes.
                GetOverlappedResult(file.value, &operation, &returned, TRUE);
                error = WAIT_TIMEOUT;
                return false;
            }
            if (!GetOverlappedResult(file.value, &operation, &returned, FALSE)) { error = GetLastError(); return false; }
        }
        if (returned != sizeof(Status) || candidate.size != sizeof(Status) || candidate.version != ProtocolVersion) {
            error = ERROR_REVISION_MISMATCH; return false;
        }
        status = candidate;
        found = true;
    }
    error = found ? ERROR_SUCCESS : ERROR_NOT_FOUND;
    return found;
}
int Hold(const Options& options) {
    if (!IsAdministrator()) { Error("administrator-required", ERROR_ACCESS_DENIED); return 2; }
    Handle cancel;
    cancel.value = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!cancel.value) { Error("event-create-failed", GetLastError()); return 2; }
    cancelEvent = cancel.value;
    if (!SetConsoleCtrlHandler(OnConsoleEvent, TRUE)) { Error("handler-registration-failed", GetLastError()); cancelEvent = nullptr; return 2; }
    struct HandlerGuard { ~HandlerGuard() { SetConsoleCtrlHandler(OnConsoleEvent, FALSE); cancelEvent = nullptr; } } handlerGuard;
    // Process-lifetime storage also covers a late async callback after startup timeout.
    static CallbackState callback;
    callback.completed = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    if (!callback.completed) { Error("event-create-failed", GetLastError()); return 2; }
    const wchar_t hardwareIds[] = L"DesktopRuntimeVddPrototype\0"; // MULTI_SZ: explicit + implicit NUL.
    SW_DEVICE_CREATE_INFO info{};
    info.cbSize = sizeof(info);
    info.pszInstanceId = InstanceId;
    info.pszzHardwareIds = hardwareIds;
    info.pszDeviceDescription = L"DesktopRuntime Virtual Display Prototype";
    info.CapabilityFlags = SWDeviceCapabilitiesRemovable | SWDeviceCapabilitiesSilentInstall | SWDeviceCapabilitiesDriverRequired;
    SoftwareDevice device;
    HRESULT hr = SwDeviceCreate(Enumerator, L"HTREE\\ROOT\\0", &info, 0, nullptr, Created, &callback, &device.value);
    if (FAILED(hr)) { Error("device-create-failed", static_cast<ULONG>(hr)); return 2; }
    HANDLE waits[] = {cancel.value, callback.completed};
    DWORD wait = WaitForMultipleObjects(2, waits, FALSE, options.startupSeconds * 1000);
    int outcome = 2;
    bool enumerated = wait == WAIT_OBJECT_0 + 1 && SUCCEEDED(callback.result);
    if (!enumerated) {
        Error(wait == WAIT_OBJECT_0 ? "cancelled" : "enumeration-failed", wait == WAIT_OBJECT_0 + 1 ? static_cast<ULONG>(callback.result) : wait);
        outcome = wait == WAIT_OBJECT_0 ? 130 : 2;
    } else {
        // Enumeration is not display readiness. Require a committed path and completed frame.
        const ULONGLONG startupDeadline = GetTickCount64() + options.startupSeconds * 1000ULL;
        Status status{};
        DWORD error = ERROR_NOT_FOUND;
        bool ready = false;
        while (WaitForSingleObject(cancel.value, 0) != WAIT_OBJECT_0 && GetTickCount64() < startupDeadline) {
            if (ReadStatus(callback.instance, status, error) && IsPipelineReady(status)) { ready = true; break; }
            WaitForSingleObject(cancel.value, 200);
        }
        if (ready) {
            PrintStatus(status, "display-pipeline-ready");
            outcome = 0;
            const ULONGLONG deadline = GetTickCount64() + options.seconds * 1000ULL;
            while (GetTickCount64() < deadline) {
                if (WaitForSingleObject(cancel.value, 500) == WAIT_OBJECT_0) { outcome = 130; break; }
                if (!ReadStatus(callback.instance, status, error) || !IsPipelineReady(status)) {
                    Error("display-pipeline-lost", error); outcome = 2; break;
                }
            }
            PrintStatus(status, "hold-ended");
        } else {
            PrintStatus(status, "display-not-ready");
            Error("readiness-failed", error);
            outcome = WaitForSingleObject(cancel.value, 0) == WAIT_OBJECT_0 ? 130 : 2;
        }
    }
    device.Close();
    // Cleanup queries only the created devnode. This is not package uninstallation.
    if (enumerated) {
        bool removed = false;
        const ULONGLONG deadline = GetTickCount64() + 10000;
        while (GetTickCount64() < deadline) {
            DEVINST instance = 0;
            CONFIGRET result = CM_Locate_DevNodeW(&instance, callback.instance, CM_LOCATE_DEVNODE_NORMAL);
            if (result == CR_NO_SUCH_DEVNODE) { removed = true; break; }
            if (result != CR_SUCCESS) break;
            Sleep(100);
        }
        std::cout << "{\"schemaVersion\":1,\"event\":\"cleanup\",\"deviceRemovalObserved\":"
            << (removed ? "true" : "false") << ",\"primaryExitCode\":" << outcome << "}" << std::endl;
        if (!removed && outcome == 0) outcome = 3;
    } else {
        std::cout << "{\"schemaVersion\":1,\"event\":\"cleanup\",\"deviceRemovalObserved\":null,\"primaryExitCode\":" << outcome << "}" << std::endl;
    }
    return outcome;
}
}
int wmain(int argc, wchar_t** argv) {
    try {
        Options options = ParseOptions(std::vector<std::wstring>(argv + 1, argv + argc));
        if (options.command == Command::Help) {
            std::cout << "DesktopVddCtl status\nDesktopVddCtl hold [--seconds 1..3600] [--startup-timeout 1..120]\n"
                "hold requires elevation and an already installed, trusted prototype driver package.\n"
                "It creates a console display only; it does not log in, unlock, launch tests or configure RDP/Citrix.\n"
                "Exit: 0 pipeline ready / hold completed, 2 unavailable, 3 cleanup unconfirmed, 64 usage, 130 cancelled.\n";
            return 0;
        }
        if (options.command == Command::Hold) return Hold(options);
        Status status{};
        DWORD error = ERROR_NOT_FOUND;
        if (!ReadStatus(nullptr, status, error)) { Error("status-unavailable", error); return 2; }
        PrintStatus(status, "status");
        return IsPipelineReady(status) ? 0 : 2;
    } catch (const std::wstring& error) {
        std::wcerr << error << L"\nUse --help for syntax.\n";
        return 64;
    } catch (const std::exception& error) {
        std::cerr << error.what() << '\n';
        return 2;
    }
}
