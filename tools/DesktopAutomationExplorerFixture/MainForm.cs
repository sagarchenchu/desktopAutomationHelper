namespace DesktopAutomationExplorerFixture;

internal sealed class MainForm : Form
{
    private readonly TextBox _eventLog;
    private readonly Label _status;
    private readonly ToolTip _toolTip = new();

    public MainForm()
    {
        Text = "DA Explorer Fixture";
        Name = UiIds.MainWindow;
        Width = 760;
        Height = 620;
        StartPosition = FormStartPosition.CenterScreen;

        var menu = new MenuStrip { Name = UiIds.MainMenu };
        var file = new ToolStripMenuItem("File") { Name = UiIds.FileMenu };
        var openDialog = new ToolStripMenuItem("Open Dialog") { Name = UiIds.OpenDialogMenuItem };
        openDialog.Click += (_, _) => OpenDialog();
        var openSecond = new ToolStripMenuItem("Open Second Window") { Name = UiIds.OpenSecondMenuItem };
        openSecond.Click += (_, _) => OpenSecond();
        file.DropDownItems.Add(openDialog);
        file.DropDownItems.Add(openSecond);
        menu.Items.Add(file);
        MainMenuStrip = menu;
        Controls.Add(menu);

        var tabs = new TabControl
        {
            Name = UiIds.TabControl,
            Left = 12,
            Top = 36,
            Width = 720,
            Height = 320
        };

        var home = new TabPage("Home") { Name = UiIds.TabHome };
        var details = new TabPage("Details") { Name = UiIds.TabDetails };

        var tree = new TreeView
        {
            Name = UiIds.Tree,
            Left = 8,
            Top = 8,
            Width = 220,
            Height = 240
        };
        var root = tree.Nodes.Add("Root");
        root.Name = UiIds.TreeRoot;
        var child = root.Nodes.Add("Child");
        child.Name = UiIds.TreeChild;
        child.Nodes.Add("Leaf");
        home.Controls.Add(tree);

        var list = new ListBox
        {
            Name = UiIds.List,
            Left = 240,
            Top = 8,
            Width = 200,
            Height = 240
        };
        list.Items.AddRange(["Alpha", "Beta", "Gamma"]);
        home.Controls.Add(list);

        var scroll = new Panel
        {
            Name = UiIds.ScrollPanel,
            Left = 460,
            Top = 8,
            Width = 230,
            Height = 240,
            AutoScroll = true,
            BorderStyle = BorderStyle.FixedSingle
        };
        for (var i = 0; i < 20; i++)
        {
            scroll.Controls.Add(new Label
            {
                Text = $"Scroll row {i + 1}",
                Top = i * 24,
                Left = 4,
                Width = 180
            });
        }

        home.Controls.Add(scroll);

        var hover = new Label
        {
            Name = UiIds.HoverTarget,
            Text = "Hover target",
            Left = 12,
            Top = 12,
            Width = 160,
            Height = 28,
            BorderStyle = BorderStyle.FixedSingle,
            TextAlign = ContentAlignment.MiddleCenter
        };
        _toolTip.SetToolTip(hover, "Safe hover tooltip");
        details.Controls.Add(hover);

        tabs.TabPages.Add(home);
        tabs.TabPages.Add(details);
        Controls.Add(tabs);

        var openDialogButton = MakeButton("Open Dialog", UiIds.OpenDialogButton, 12, 370, OpenDialog);
        var openSecondButton = MakeButton("Open Second", UiIds.OpenSecondButton, 140, 370, OpenSecond);
        // Clearly marked destructive controls — explorer must refuse these.
        var deleteButton = MakeButton("Delete", UiIds.DeleteButton, 280, 370, () => Log("DESTRUCTIVE: Delete"));
        var submitButton = MakeButton("Submit", UiIds.SubmitButton, 400, 370, () => Log("DESTRUCTIVE: Submit"));
        var saveButton = MakeButton("Save", UiIds.SaveButton, 520, 370, () => Log("DESTRUCTIVE: Save"));
        var purchaseButton = MakeButton("Purchase", UiIds.PurchaseButton, 640, 370, () => Log("DESTRUCTIVE: Purchase"));

        Controls.Add(openDialogButton);
        Controls.Add(openSecondButton);
        Controls.Add(deleteButton);
        Controls.Add(submitButton);
        Controls.Add(saveButton);
        Controls.Add(purchaseButton);

        _status = new Label
        {
            Name = UiIds.StatusLabel,
            Text = "Status: Ready",
            Left = 12,
            Top = 410,
            Width = 720,
            Height = 24
        };
        Controls.Add(_status);

        _eventLog = new TextBox
        {
            Name = UiIds.EventLog,
            Left = 12,
            Top = 440,
            Width = 720,
            Height = 120,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical
        };
        Controls.Add(_eventLog);
        Log("Fixture started");
    }

    private static Button MakeButton(string text, string automationId, int left, int top, Action onClick)
    {
        var button = new Button
        {
            Text = text,
            Name = automationId,
            Left = left,
            Top = top,
            Width = 110,
            Height = 28
        };
        button.Click += (_, _) => onClick();
        return button;
    }

    private void OpenDialog()
    {
        Log("Open dialog");
        using var dialog = new SampleDialog();
        dialog.ShowDialog(this);
        Log("Dialog closed");
    }

    private void OpenSecond()
    {
        Log("Open second window");
        var second = new SecondWindowForm();
        second.Show(this);
    }

    private void Log(string message)
    {
        var line = $"{DateTime.Now:HH:mm:ss.fff} {message}";
        _status.Text = $"Status: {message}";
        _eventLog.AppendText(line + Environment.NewLine);
    }
}
