// Copyright (c) 2026 Reddysagar Chenchu.
#ifdef _WIN32
#include <windows.h>
#endif
#include "../Shared/Protocol.h"
#include "../Control/Options.h"
#include <iostream>
#include <stdexcept>
using namespace desktop_vdd;
void Require(bool condition, const char* message) { if (!condition) throw std::runtime_error(message); }
int main() {
    try {
        Status status{sizeof(Status), ProtocolVersion, AdapterReady | MonitorArrived | PathActive | SwapChainActive | PoweredOn,
            0, Width, Height, RefreshRate, 0, 1};
        Require(IsPipelineReady(status), "complete evidence should pass");
        for (auto flag : {AdapterReady, MonitorArrived, PathActive, SwapChainActive, PoweredOn}) {
            auto changed = status; changed.flags &= ~static_cast<std::uint32_t>(flag);
            Require(!IsPipelineReady(changed), "missing readiness stage must block");
        }
        auto changed = status; changed.completedFrames = 0;
        Require(!IsPipelineReady(changed), "enumeration without a frame must block");
        changed = status; changed.version = 99;
        Require(!IsPipelineReady(changed), "unknown protocol must block");
        changed = status; changed.size = 4;
        Require(!IsPipelineReady(changed), "truncated response must block");
        changed = status; changed.width = 800;
        Require(!IsPipelineReady(changed), "wrong mode must block");
        Require(ParseOptions({}).command == Command::Help, "default help");
        Require(ParseOptions({L"status"}).command == Command::Status, "read-only status");
        auto options = ParseOptions({L"hold", L"--seconds", L"1", L"--startup-timeout", L"120"});
        Require(options.seconds == 1 && options.startupSeconds == 120, "bounded values accepted");
        for (auto args : std::vector<std::vector<std::wstring>>{
            {L"install"}, {L"status", L"--seconds", L"1"}, {L"hold", L"--seconds"},
            {L"hold", L"--seconds", L"0"}, {L"hold", L"--seconds", L"3601"},
            {L"hold", L"--seconds", L"-1"}, {L"hold", L"--seconds", L"1x"},
            {L"hold", L"--seconds", L"9999999999999999999"}, {L"hold", L"--seconds", L""},
            {L"hold", L"--seconds", L"1", L"--seconds", L"2"}, {L"hold", L"--startup-timeout", L"121"}}) {
            bool rejected = false;
            try { ParseOptions(args); } catch (const std::wstring&) { rejected = true; }
            Require(rejected, "invalid CLI input must not reach native mutation");
        }
        std::cout << "VDD protocol and CLI contract tests passed.\n";
        return 0;
    } catch (const std::exception& error) { std::cerr << error.what() << '\n'; return 1; }
}
