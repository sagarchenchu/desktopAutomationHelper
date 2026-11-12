using System.ComponentModel;

namespace DesktopAutomationTestApp;

public sealed class MainForm : Form
{
    private readonly TextBox _eventLog = new();
    private readonly Label _statusLabel = new();
    private readonly TextBox _nameEdit = new();
    private readonly TextBox _notesEdit = new();
    private readonly ComboBox _countryCombo = new();
    private readonly ComboBox _typeAheadCombo = new();
    private readonly Label _selectedCountryLabel = new();
    private readonly Label _selectedTypeAheadLabel = new();
    private readonly Label _selectedRoleLabel = new();
    private readonly Label _selectedGridCellLabel = new();
    private readonly DataGridView _grid = new();
    private readonly BindingList<SampleRow> _rows = new();

    private ChildForm? _child;
    private SecondWindowForm? _second;

    public MainForm()
    {
        Name = UiIds.MainWindow;
        Text = "DA Test Harness";
        AccessibleName = "DA Test Harness";
        StartPosition = FormStartPosition.CenterScreen;
        Width = 980;
        Height = 760;
        MinimumSize = new Size(860, 640);

        BuildMenu();
        BuildLayout();
        SeedGrid();
        Log("Main window ready.");
    }

    private void BuildMenu()
    {
        var menu = new MenuStrip { Name = UiIds.MainMenu, AccessibleName = "Main Menu" };

        var file = new ToolStripMenuItem("File") { Name = UiIds.FileMenu, AccessibleName = "File" };
        var openChild = new ToolStripMenuItem("Open Child Window")
        {
            Name = UiIds.FileOpenChildMenuItem,
            AccessibleName = "Open Child Window"
        };
        openChild.Click += (_, _) => OpenChildWindow();
        var exit = new ToolStripMenuItem("Exit")
        {
            Name = UiIds.FileExitMenuItem,
            AccessibleName = "Exit"
        };
        exit.Click += (_, _) => Close();
        file.DropDownItems.Add(openChild);
        file.DropDownItems.Add(exit);

        var actions = new ToolStripMenuItem("Actions") { Name = UiIds.ActionsMenu, AccessibleName = "Actions" };
        var clearLog = new ToolStripMenuItem("Clear Log")
        {
            Name = UiIds.ActionsClearLogMenuItem,
            AccessibleName = "Clear Log"
        };
        clearLog.Click += (_, _) =>
        {
            _eventLog.Clear();
            Log("Log cleared.");
        };
        actions.DropDownItems.Add(clearLog);

        menu.Items.Add(file);
        menu.Items.Add(actions);
        MainMenuStrip = menu;
        Controls.Add(menu);
    }

    private void BuildLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 2,
            Padding = new Padding(8, 32, 8, 8)
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 70));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 30));

        var left = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5
        };
        for (var i = 0; i < 5; i++)
            left.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        left.Controls.Add(BuildButtonsGroup(), 0, 0);
        left.Controls.Add(BuildTextGroup(), 0, 1);
        left.Controls.Add(BuildComboGroup(), 0, 2);
        left.Controls.Add(BuildSelectionGroup(), 0, 3);
        left.Controls.Add(BuildWindowGroup(), 0, 4);

        var right = BuildGridGroup();
        right.Dock = DockStyle.Fill;

        _statusLabel.Name = UiIds.StatusLabel;
        _statusLabel.AccessibleName = "Status";
        _statusLabel.AutoSize = true;
        _statusLabel.Text = "Status: Ready";
        _statusLabel.Padding = new Padding(0, 4, 0, 4);

        _eventLog.Name = UiIds.EventLog;
        _eventLog.AccessibleName = "Event Log";
        _eventLog.Multiline = true;
        _eventLog.ScrollBars = ScrollBars.Vertical;
        _eventLog.ReadOnly = true;
        _eventLog.Dock = DockStyle.Fill;

        var logPanel = new Panel { Dock = DockStyle.Fill };
        logPanel.Controls.Add(_eventLog);
        logPanel.Controls.Add(_statusLabel);
        _statusLabel.Dock = DockStyle.Top;
        _eventLog.Dock = DockStyle.Fill;

        root.Controls.Add(left, 0, 0);
        root.Controls.Add(right, 1, 0);
        root.SetColumnSpan(logPanel, 2);
        root.Controls.Add(logPanel, 0, 1);

        Controls.Add(root);
    }

    private Control BuildButtonsGroup()
    {
        var box = Group("Buttons — click / double-click / right-click");
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };

        var click = Button(UiIds.ClickButton, "Click Me");
        click.Click += (_, _) =>
        {
            SetStatus("ClickButton clicked");
            Log("ClickButton: single click");
        };

        // Button suppresses StandardDoubleClick; use a control that raises MouseDoubleClick,
        // plus Click timing fallback for drivers that emit two rapid single clicks.
        var dbl = DoubleClickButton(UiIds.DoubleClickButton, "Double-Click Me");
        var lastDblClickUtc = DateTime.MinValue;
        var doubleClickHandled = false;
        dbl.MouseDoubleClick += (_, _) =>
        {
            doubleClickHandled = true;
            lastDblClickUtc = DateTime.MinValue;
            SetStatus("DoubleClickButton double-clicked");
            Log("DoubleClickButton: double click");
        };
        dbl.Click += (_, _) =>
        {
            if (doubleClickHandled)
            {
                doubleClickHandled = false;
                return;
            }

            var now = DateTime.UtcNow;
            if (lastDblClickUtc != DateTime.MinValue
                && (now - lastDblClickUtc).TotalMilliseconds <= SystemInformation.DoubleClickTime)
            {
                lastDblClickUtc = DateTime.MinValue;
                SetStatus("DoubleClickButton double-clicked");
                Log("DoubleClickButton: double click");
                return;
            }

            lastDblClickUtc = now;
            Log("DoubleClickButton: single click (use double-click for primary assert)");
        };

        var right = Button(UiIds.RightClickButton, "Right-Click Me");
        var menu = new ContextMenuStrip();
        menu.Items.Add(new ToolStripMenuItem("Copy")
        {
            Name = UiIds.RightClickCopyMenuItem,
            AccessibleName = "Copy"
        });
        menu.Items[^1].Click += (_, _) => Log("RightClickButton context: Copy");
        menu.Items.Add(new ToolStripMenuItem("Inspect")
        {
            Name = UiIds.RightClickInspectMenuItem,
            AccessibleName = "Inspect"
        });
        menu.Items[^1].Click += (_, _) => Log("RightClickButton context: Inspect");
        right.ContextMenuStrip = menu;
        right.MouseUp += (_, e) =>
        {
            if (e.Button == MouseButtons.Right)
            {
                SetStatus("RightClickButton right-clicked");
                Log("RightClickButton: right click");
            }
        };

        var disabled = Button(UiIds.DisabledButton, "Disabled");
        disabled.Enabled = false;

        flow.Controls.Add(click);
        flow.Controls.Add(dbl);
        flow.Controls.Add(right);
        flow.Controls.Add(disabled);
        box.Controls.Add(flow);
        return box;
    }

    private Control BuildTextGroup()
    {
        var box = Group("Text — type / clear");
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 2,
            AutoSize = true
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _nameEdit.Name = UiIds.NameEdit;
        _nameEdit.AccessibleName = "Name";
        _nameEdit.Width = 220;
        _nameEdit.TextChanged += (_, _) => Log($"NameEdit changed -> '{_nameEdit.Text}'");

        var clear = Button(UiIds.ClearNameButton, "Clear Name");
        clear.Click += (_, _) =>
        {
            _nameEdit.Clear();
            SetStatus("Name cleared");
            Log("ClearNameButton: cleared NameEdit");
        };

        _notesEdit.Name = UiIds.NotesEdit;
        _notesEdit.AccessibleName = "Notes";
        _notesEdit.Multiline = true;
        _notesEdit.Height = 48;
        _notesEdit.Dock = DockStyle.Fill;

        table.Controls.Add(new Label { Text = "Name:", AutoSize = true }, 0, 0);
        table.Controls.Add(_nameEdit, 1, 0);
        table.Controls.Add(clear, 2, 0);
        table.Controls.Add(new Label { Text = "Notes:", AutoSize = true }, 0, 1);
        table.Controls.Add(_notesEdit, 1, 1);
        table.SetColumnSpan(_notesEdit, 2);
        box.Controls.Add(table);
        return box;
    }

    private Control BuildComboGroup()
    {
        var box = Group("Dropdowns — select / type-ahead");
        var table = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true
        };
        table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _countryCombo.Name = UiIds.CountryCombo;
        _countryCombo.AccessibleName = "Country";
        _countryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
        _countryCombo.Items.AddRange(["United States", "Canada", "United Kingdom", "Germany", "India", "Japan"]);
        _countryCombo.SelectedIndex = 0;
        _countryCombo.SelectedIndexChanged += (_, _) =>
        {
            _selectedCountryLabel.Text = $"Selected country: {_countryCombo.SelectedItem}";
            Log($"CountryCombo selected -> '{_countryCombo.SelectedItem}'");
        };

        _typeAheadCombo.Name = UiIds.TypeAheadCombo;
        _typeAheadCombo.AccessibleName = "Type Ahead City";
        _typeAheadCombo.DropDownStyle = ComboBoxStyle.DropDown;
        _typeAheadCombo.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        _typeAheadCombo.AutoCompleteSource = AutoCompleteSource.ListItems;
        _typeAheadCombo.Items.AddRange([
            "Amsterdam", "Athens", "Atlanta", "Austin",
            "Bangalore", "Berlin", "Boston",
            "Cairo", "Chicago", "Copenhagen",
            "Dallas", "Delhi", "Denver",
            "London", "Los Angeles", "Madrid", "Mumbai", "New York", "Paris", "Seattle", "Tokyo"
        ]);
        _typeAheadCombo.TextChanged += (_, _) =>
        {
            _selectedTypeAheadLabel.Text = $"Type-ahead text: {_typeAheadCombo.Text}";
            Log($"TypeAheadCombo text -> '{_typeAheadCombo.Text}'");
        };
        _typeAheadCombo.SelectedIndexChanged += (_, _) =>
        {
            Log($"TypeAheadCombo selected -> '{_typeAheadCombo.SelectedItem}'");
        };

        _selectedCountryLabel.Name = UiIds.SelectedCountryLabel;
        _selectedCountryLabel.AccessibleName = "Selected Country";
        _selectedCountryLabel.AutoSize = true;
        _selectedCountryLabel.Text = "Selected country: United States";

        _selectedTypeAheadLabel.Name = UiIds.SelectedTypeAheadLabel;
        _selectedTypeAheadLabel.AccessibleName = "Selected Type Ahead";
        _selectedTypeAheadLabel.AutoSize = true;
        _selectedTypeAheadLabel.Text = "Type-ahead text: ";

        table.Controls.Add(new Label { Text = "Country:", AutoSize = true }, 0, 0);
        table.Controls.Add(_countryCombo, 1, 0);
        table.Controls.Add(new Label { Text = "City type-ahead:", AutoSize = true }, 0, 1);
        table.Controls.Add(_typeAheadCombo, 1, 1);
        table.Controls.Add(_selectedCountryLabel, 1, 2);
        table.Controls.Add(_selectedTypeAheadLabel, 1, 3);
        box.Controls.Add(table);
        return box;
    }

    private Control BuildSelectionGroup()
    {
        var box = Group("Selection — radio / checkbox");
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.LeftToRight };

        var admin = Radio(UiIds.RoleRadioAdmin, "Admin");
        var user = Radio(UiIds.RoleRadioUser, "User");
        var guest = Radio(UiIds.RoleRadioGuest, "Guest");
        user.Checked = true;

        void OnRole(object? sender, EventArgs e)
        {
            if (sender is RadioButton { Checked: true } rb)
            {
                _selectedRoleLabel.Text = $"Selected role: {rb.Text}";
                Log($"Role selected -> '{rb.Text}'");
            }
        }

        admin.CheckedChanged += OnRole;
        user.CheckedChanged += OnRole;
        guest.CheckedChanged += OnRole;

        var agree = new CheckBox
        {
            Name = UiIds.AgreeCheckBox,
            AccessibleName = "Agree",
            Text = "I agree",
            AutoSize = true
        };
        agree.CheckedChanged += (_, _) =>
        {
            Log($"AgreeCheckBox -> {(agree.Checked ? "checked" : "unchecked")}");
            SetStatus(agree.Checked ? "Agreed" : "Not agreed");
        };

        _selectedRoleLabel.Name = UiIds.SelectedRoleLabel;
        _selectedRoleLabel.AccessibleName = "Selected Role";
        _selectedRoleLabel.AutoSize = true;
        _selectedRoleLabel.Text = "Selected role: User";

        flow.Controls.Add(admin);
        flow.Controls.Add(user);
        flow.Controls.Add(guest);
        flow.Controls.Add(agree);
        flow.Controls.Add(_selectedRoleLabel);
        box.Controls.Add(flow);
        return box;
    }

    private Control BuildWindowGroup()
    {
        var box = Group("Windows — child / second / dialog / close");
        var flow = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true };

        var child = Button(UiIds.OpenChildWindowButton, "Open Child Window");
        child.Click += (_, _) => OpenChildWindow();

        var second = Button(UiIds.OpenSecondWindowButton, "Open Second Window");
        second.Click += (_, _) => OpenSecondWindow();

        var dialog = Button(UiIds.OpenDialogButton, "Open Dialog");
        dialog.Click += (_, _) =>
        {
            using var dlg = new SampleDialog();
            var result = dlg.ShowDialog(this);
            Log($"SampleDialog closed with {result}");
            SetStatus($"Dialog: {result}");
        };

        var close = Button(UiIds.CloseMainButton, "Close Main");
        close.Click += (_, _) => Close();

        flow.Controls.Add(child);
        flow.Controls.Add(second);
        flow.Controls.Add(dialog);
        flow.Controls.Add(close);
        box.Controls.Add(flow);
        return box;
    }

    private Control BuildGridGroup()
    {
        var box = Group("Grid / table headers");
        box.Dock = DockStyle.Fill;

        _grid.Name = UiIds.SampleGrid;
        _grid.AccessibleName = "Sample Grid";
        _grid.Dock = DockStyle.Fill;
        _grid.AllowUserToAddRows = false;
        _grid.AllowUserToDeleteRows = false;
        _grid.MultiSelect = false;
        _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        _grid.DataSource = _rows;
        _grid.CellClick += (_, e) =>
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
                return;
            var col = _grid.Columns[e.ColumnIndex].HeaderText;
            var val = _grid[e.ColumnIndex, e.RowIndex].Value?.ToString() ?? "";
            _selectedGridCellLabel.Text = $"Selected cell: [{e.RowIndex},{e.ColumnIndex}] {col}={val}";
            Log($"SampleGrid cell click -> row={e.RowIndex} col={e.ColumnIndex} {col}='{val}'");
        };

        _selectedGridCellLabel.Name = UiIds.SelectedGridCellLabel;
        _selectedGridCellLabel.AccessibleName = "Selected Grid Cell";
        _selectedGridCellLabel.Dock = DockStyle.Bottom;
        _selectedGridCellLabel.Height = 28;
        _selectedGridCellLabel.Text = "Selected cell: (none)";

        var refresh = Button(UiIds.RefreshGridButton, "Refresh Grid");
        refresh.Dock = DockStyle.Top;
        refresh.Click += (_, _) =>
        {
            SeedGrid();
            Log("SampleGrid refreshed");
            SetStatus("Grid refreshed");
        };

        box.Controls.Add(_grid);
        box.Controls.Add(_selectedGridCellLabel);
        box.Controls.Add(refresh);
        return box;
    }

    private void SeedGrid()
    {
        _rows.Clear();
        _rows.Add(new SampleRow { Id = "1", Name = "Alpha", Role = "Admin", Status = "Active" });
        _rows.Add(new SampleRow { Id = "2", Name = "Bravo", Role = "User", Status = "Active" });
        _rows.Add(new SampleRow { Id = "3", Name = "Charlie", Role = "Guest", Status = "Disabled" });
        _rows.Add(new SampleRow { Id = "4", Name = "Delta", Role = "User", Status = "Pending" });
        _rows.Add(new SampleRow { Id = "5", Name = "Echo", Role = "Admin", Status = "Active" });
    }

    private void OpenChildWindow()
    {
        if (_child is { IsDisposed: false })
        {
            _child.Activate();
            Log("ChildWindow already open — activated");
            return;
        }

        _child = new ChildForm();
        _child.FormClosed += (_, _) =>
        {
            Log("ChildWindow closed");
            _child = null;
        };
        _child.Show(this);
        Log("ChildWindow opened");
        SetStatus("Child window open");
    }

    private void OpenSecondWindow()
    {
        if (_second is { IsDisposed: false })
        {
            _second.Activate();
            Log("SecondWindow already open — activated");
            return;
        }

        _second = new SecondWindowForm();
        _second.FormClosed += (_, _) =>
        {
            Log("SecondWindow closed");
            _second = null;
        };
        _second.Show();
        Log("SecondWindow opened");
        SetStatus("Second window open");
    }

    private void SetStatus(string text)
    {
        _statusLabel.Text = $"Status: {text}";
    }

    private void Log(string message)
    {
        var line = $"{DateTime.Now:HH:mm:ss.fff}  {message}";
        if (_eventLog.TextLength > 0)
            _eventLog.AppendText(Environment.NewLine);
        _eventLog.AppendText(line);
    }

    private static GroupBox Group(string title) =>
        new()
        {
            Text = title,
            AutoSize = true,
            Dock = DockStyle.Top,
            Padding = new Padding(8),
            Margin = new Padding(0, 0, 0, 8)
        };

    private static Button Button(string id, string text) =>
        new()
        {
            Name = id,
            AccessibleName = text,
            Text = text,
            AutoSize = true,
            Margin = new Padding(4)
        };

    private static DoubleClickAwareButton DoubleClickButton(string id, string text) =>
        new()
        {
            Name = id,
            AccessibleName = text,
            Text = text,
            AutoSize = true,
            Margin = new Padding(4)
        };

    private static RadioButton Radio(string id, string text) =>
        new()
        {
            Name = id,
            AccessibleName = text,
            Text = text,
            AutoSize = true,
            Margin = new Padding(4)
        };

    private sealed class SampleRow
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Role { get; set; } = "";
        public string Status { get; set; } = "";
    }
}
