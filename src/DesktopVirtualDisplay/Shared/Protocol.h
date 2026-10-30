// Copyright (c) 2026 Reddysagar Chenchu.
#pragma once
#include <cstdint>

namespace desktop_vdd {
constexpr std::uint32_t ProtocolVersion = 1;
constexpr std::uint32_t Width = 1920;
constexpr std::uint32_t Height = 1080;
constexpr std::uint32_t RefreshRate = 60;
enum Flag : std::uint32_t {
    AdapterReady = 1, MonitorArrived = 2, PathActive = 4, SwapChainActive = 8,
    PoweredOn = 16
};
struct Status {
    std::uint32_t size;
    std::uint32_t version;
    std::uint32_t flags;
    std::int32_t lastError; // NTSTATUS or HRESULT; diagnostic only, not Win32 GetLastError.
    std::uint32_t width;
    std::uint32_t height;
    std::uint32_t refreshRate;
    std::uint32_t reserved;
    std::uint64_t completedFrames;
};
static_assert(sizeof(Status) == 40, "The status ABI must remain fixed.");
inline bool IsPipelineReady(const Status& s) noexcept {
    constexpr auto required = AdapterReady | MonitorArrived | PathActive | SwapChainActive | PoweredOn;
    return s.size == sizeof(Status) && s.version == ProtocolVersion &&
        (s.flags & required) == required && s.completedFrames > 0 &&
        s.width == Width && s.height == Height && s.refreshRate == RefreshRate;
}
// Pipeline readiness deliberately says nothing about input, screenshots or test success.
}

#ifdef _WIN32
#include <guiddef.h>
#include <winioctl.h>
namespace desktop_vdd {
inline constexpr GUID InterfaceId = {0x5854e0bf, 0x5fb8, 0x42cf, {0x8e, 0x17, 0x9d, 0x3b, 0x74, 0x34, 0x51, 0x23}};
inline constexpr GUID MonitorContainerId = {0x94ab07db, 0x0e98, 0x4fd3, {0x98, 0x69, 0x83, 0x84, 0x7c, 0x13, 0x3f, 0x38}};
inline constexpr wchar_t HardwareId[] = L"DesktopRuntimeVddPrototype";
inline constexpr wchar_t Enumerator[] = L"DesktopRuntimeVddPrototype";
inline constexpr wchar_t InstanceId[] = L"Console";
constexpr DWORD QueryStatus = CTL_CODE(FILE_DEVICE_UNKNOWN, 0x800, METHOD_BUFFERED, FILE_READ_DATA);
}
#endif
