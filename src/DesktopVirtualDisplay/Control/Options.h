// Copyright (c) 2026 Reddysagar Chenchu.
#pragma once
#include <string>
#include <vector>

namespace desktop_vdd {
enum class Command { Help, Status, Hold };
struct Options { Command command = Command::Help; unsigned seconds = 300; unsigned startupSeconds = 30; };
inline unsigned ParseSeconds(const std::wstring& text, unsigned maximum) {
    if (text.empty()) throw std::wstring(L"A duration is required.");
    unsigned value = 0;
    for (wchar_t ch : text) {
        if (ch < L'0' || ch > L'9' || value > maximum / 10) throw std::wstring(L"Invalid duration.");
        value = value * 10 + static_cast<unsigned>(ch - L'0');
        if (value > maximum) throw std::wstring(L"Duration exceeds limit.");
    }
    if (value == 0) throw std::wstring(L"Duration must be positive.");
    return value;
}
inline Options ParseOptions(const std::vector<std::wstring>& args) {
    Options result;
    if (args.empty() || (args.size() == 1 && (args[0] == L"--help" || args[0] == L"-h"))) return result;
    if (args.size() == 1 && args[0] == L"status") { result.command = Command::Status; return result; }
    if (args[0] != L"hold") throw std::wstring(L"Expected status or hold.");
    result.command = Command::Hold;
    bool duration = false, startup = false;
    for (std::size_t i = 1; i < args.size(); i += 2) {
        if (i + 1 == args.size()) throw std::wstring(L"Missing option value.");
        if (args[i] == L"--seconds" && !duration) { result.seconds = ParseSeconds(args[i + 1], 3600); duration = true; }
        else if (args[i] == L"--startup-timeout" && !startup) { result.startupSeconds = ParseSeconds(args[i + 1], 120); startup = true; }
        else throw std::wstring(L"Unknown or repeated option.");
    }
    return result;
}
}
