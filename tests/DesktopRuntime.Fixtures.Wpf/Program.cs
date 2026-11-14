using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace DesktopRuntime.Fixtures.Wpf;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        var window = new FixtureWindow();
        if (args.Length != 0)
        {
            if (args.Length != 2 || args[0] != "--display")
                throw new ArgumentException("Usage: DesktopRuntime.Fixtures.Wpf.exe [--display device-name]");
            var screen = System.Windows.Forms.Screen.AllScreens.SingleOrDefault(s => s.DeviceName.Equals(args[1], StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException("Requested display is not active.");
            window.SourceInitialized += (_, _) => {
                if (!SetWindowPos(new WindowInteropHelper(window).Handle, IntPtr.Zero,
                    screen.WorkingArea.Left + 30, screen.WorkingArea.Top + 30, 920, 640, 0x0004))
                    throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
            };
        }
        new Application().Run(window);
    }
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetWindowPos(IntPtr window, IntPtr after, int x, int y, int width, int height, uint flags);
}

internal sealed class FixtureWindow : Window
{
    private readonly RecordStore store = new();
    private readonly TextBox name = Identified(new TextBox { Width = 240 }, "NameInput");
    private readonly TextBox age = Identified(new TextBox { Width = 80, Text = "30" }, "AgeInput");
    private readonly ComboBox role = Identified(new ComboBox { Width = 150, ItemsSource = RecordStore.Roles, SelectedIndex = 0 }, "RoleInput");
    private readonly DataGrid records = Identified(new DataGrid { IsReadOnly = true, CanUserAddRows = false,
        CanUserDeleteRows = false, SelectionMode = DataGridSelectionMode.Single, AutoGenerateColumns = true,
        EnableRowVirtualization = false }, "RecordsGrid");
    private readonly TextBlock result = Identified(new TextBlock { Text = "Ready.", Margin = new Thickness(0, 10, 0, 0) }, "ResultText");
    private readonly TextBlock marker = Identified(new TextBlock { Height = 36, Width = 760 }, "FrameMarker");

    public FixtureWindow()
    {
        Title = "DesktopRuntime Fixture - WPF";
        Width = 920; Height = 640;
        AutomationProperties.SetAutomationId(this, "DesktopRuntimeFixture");
        var layout = new DockPanel { Margin = new Thickness(12) };
        AddTop(layout, marker);
        var inputs = new WrapPanel { Margin = new Thickness(0, 10, 0, 10), Height = 90 };
        foreach (var element in new FrameworkElement[] { new TextBlock { Text = "Name" }, name,
            new TextBlock { Text = "Age" }, age, new TextBlock { Text = "Role" }, role,
            Identified(new TextBox { Width = 170, Text = "Cannot edit", IsEnabled = false }, "DisabledInput") })
        { element.Margin = new Thickness(4); inputs.Children.Add(element); }
        AddTop(layout, inputs);
        var buttons = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 0, 0, 10) };
        buttons.Children.Add(Button("AddRecord", "Add", () => store.Add(name.Text, age.Text, role.Text)));
        buttons.Children.Add(Button("UpdateRecord", "Update", () => store.Update(SelectedId(), name.Text, age.Text, role.Text)));
        buttons.Children.Add(Button("DeleteRecord", "Delete", () => store.Delete(SelectedId())));
        AddTop(layout, buttons);
        DockPanel.SetDock(result, Dock.Bottom); layout.Children.Add(result);
        records.ItemsSource = store.Records; layout.Children.Add(records);
        Content = layout;
        var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
        int ticks = 0;
        timer.Tick += (_, _) => {
            marker.Text = $"FRAME {++ticks:D6} | {DateTime.UtcNow:O}";
            marker.Background = ticks % 2 == 0 ? Brushes.LightBlue : Brushes.LightGoldenrodYellow;
        };
        Closed += (_, _) => timer.Stop();
        timer.Start();
    }
    private static T Identified<T>(T element, string id) where T : DependencyObject
    { AutomationProperties.SetAutomationId(element, id); return element; }
    private static void AddTop(DockPanel panel, UIElement element)
    { DockPanel.SetDock(element, Dock.Top); panel.Children.Add(element); }
    private int SelectedId() => records.SelectedItem is CustomerRecord record
        ? record.Id : throw new ArgumentException("Select an existing row.");
    private Button Button(string id, string text, Func<string> action)
    {
        var button = Identified(new Button { Content = text, Width = 120, Margin = new Thickness(4) }, id);
        button.Click += (_, _) => { try { result.Text = action(); } catch (ArgumentException error) { result.Text = error.Message; } };
        return button;
    }
}
