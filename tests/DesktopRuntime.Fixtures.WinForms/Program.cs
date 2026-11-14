using DesktopRuntime.Fixtures;

namespace DesktopRuntime.Fixtures.WinForms;

internal static class Program
{
    [STAThread]
    private static void Main(string[] args)
    {
        ApplicationConfiguration.Initialize();
        using var form = new FixtureForm();
        if (args.Length != 0)
        {
            if (args.Length != 2 || args[0] != "--display")
                throw new ArgumentException("Usage: DesktopRuntime.Fixtures.WinForms.exe [--display device-name]");
            var screen = Screen.AllScreens.SingleOrDefault(s => s.DeviceName.Equals(args[1], StringComparison.OrdinalIgnoreCase))
                ?? throw new ArgumentException("Requested display is not active.");
            form.StartPosition = FormStartPosition.Manual;
            form.Location = new Point(screen.WorkingArea.Left + 30, screen.WorkingArea.Top + 30);
        }
        Application.Run(form);
    }
}

internal sealed class FixtureForm : Form
{
    private readonly RecordStore store = new();
    private readonly TextBox name = new() { Name = "NameInput", AccessibleName = "Name", Width = 260 };
    private readonly TextBox age = new() { Name = "AgeInput", AccessibleName = "Age", Text = "30", Width = 90 };
    private readonly ComboBox role = new() { Name = "RoleInput", AccessibleName = "Role", DropDownStyle = ComboBoxStyle.DropDownList, Width = 170 };
    private readonly DataGridView grid = new() { Name = "RecordsGrid", AccessibleName = "Records", Dock = DockStyle.Fill,
        ReadOnly = true, AllowUserToAddRows = false, AllowUserToDeleteRows = false, MultiSelect = false,
        SelectionMode = DataGridViewSelectionMode.FullRowSelect, AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill };
    private readonly Label result = new() { Name = "ResultText", AutoSize = true, Text = "Ready." };
    private readonly Label marker = new() { Name = "FrameMarker", AutoSize = false, Width = 760, Height = 36 };
    private readonly System.Windows.Forms.Timer timer = new() { Interval = 500 };
    private int ticks;

    public FixtureForm()
    {
        Name = "DesktopRuntimeFixture";
        Text = "DesktopRuntime Fixture - WinForms";
        ClientSize = new Size(920, 640);
        var layout = new TableLayoutPanel { Dock = DockStyle.Fill, RowCount = 5, ColumnCount = 1, Padding = new Padding(12) };
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 45));
        layout.Controls.Add(marker, 0, 0);
        var inputs = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = true };
        role.Items.AddRange(RecordStore.Roles);
        role.SelectedIndex = 0;
        inputs.Controls.AddRange([new Label { Text = "Name", AutoSize = true }, name,
            new Label { Text = "Age", AutoSize = true }, age,
            new Label { Text = "Role", AutoSize = true }, role,
            new TextBox { Name = "DisabledInput", AccessibleName = "Disabled field", Text = "Cannot edit", Enabled = false, Width = 180 }]);
        layout.Controls.Add(inputs, 0, 1);
        var buttons = new FlowLayoutPanel { Dock = DockStyle.Fill };
        buttons.Controls.Add(Button("AddRecord", "Add", () => store.Add(name.Text, age.Text, role.Text)));
        buttons.Controls.Add(Button("UpdateRecord", "Update", () => store.Update(SelectedId(), name.Text, age.Text, role.Text)));
        buttons.Controls.Add(Button("DeleteRecord", "Delete", () => store.Delete(SelectedId())));
        layout.Controls.Add(buttons, 0, 2);
        grid.DataSource = store.Records;
        layout.Controls.Add(grid, 0, 3);
        layout.Controls.Add(result, 0, 4);
        Controls.Add(layout);
        timer.Tick += (_, _) => {
            marker.Text = $"FRAME {++ticks:D6} | {DateTime.UtcNow:O} | {Screen.FromControl(this).DeviceName}";
            marker.BackColor = ticks % 2 == 0 ? Color.LightBlue : Color.LightGoldenrodYellow;
        };
        timer.Start();
    }
    private int SelectedId() => grid.CurrentRow?.DataBoundItem is CustomerRecord row
        ? row.Id : throw new ArgumentException("Select an existing row.");
    private Button Button(string id, string text, Func<string> action)
    {
        var button = new Button { Name = id, AccessibleName = text, Text = text, Width = 120 };
        button.Click += (_, _) => { try { result.Text = action(); } catch (ArgumentException error) { result.Text = error.Message; } };
        return button;
    }
    protected override void Dispose(bool disposing) { if (disposing) timer.Dispose(); base.Dispose(disposing); }
}
