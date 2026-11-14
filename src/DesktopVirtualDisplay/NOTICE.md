# Source attribution

`Driver/Driver.cpp`, the driver INF/build configuration and native package imports
are adapted from Microsoft's Windows Driver Samples, IndirectDisplay example.
Original copyright: Microsoft Corporation. These derived files are distributed
under the Microsoft Public License included in `LICENSE-MS-PL.txt`.

Reference repository: https://github.com/microsoft/Windows-driver-samples
Pinned reference commit: `f4c1ff842b6a0001152f44cb61788b216a1c7627`
Reference paths: `video/IndirectDisplay/IddSampleDriver`, `Directory.Build.props`.

Changes by Reddysagar Chenchu include a descriptorless single-mode sink, bounded
mode output, termination-aware frame consumption, status telemetry, owned-device
lifetime utility, Server 2019 compatibility settings and acceptance fixtures.
Source attribution is retained independently from Git commit authorship.

The prototype is not endorsed or certified by Microsoft. It does not include or
redistribute a prebuilt third-party VDD. Microsoft SDK/WDK NuGet dependencies keep
their own license terms and are acquired during the build.
