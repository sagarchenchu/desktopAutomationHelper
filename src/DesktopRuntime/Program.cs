using DesktopExecutionRuntime;
using DesktopRuntime;

return DoctorCommand.Run(args, new WindowsDesktopEnvironmentProbe(), Console.Out, Console.Error);
