using System.Text.Json;
using DesktopAutomationDriver.Models.Playback;

namespace DesktopAutomationDriver.Services;

public interface IPlaybackService
{
    PlaybackResult Play(JsonElement payload);
}
