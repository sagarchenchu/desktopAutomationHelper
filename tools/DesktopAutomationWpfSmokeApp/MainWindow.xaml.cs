using System.Windows;

namespace DesktopAutomationWpfSmokeApp;

public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void PrimaryButton_Click(object sender, RoutedEventArgs e)
    {
        StatusLabel.Text = "Status: Primary clicked";
    }

    private void OpenDialog_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new ChildDialog
        {
            Owner = this
        };
        // Modeless so observation can target the child while it remains open.
        dialog.Show();
        StatusLabel.Text = "Status: Dialog open";
    }
}
