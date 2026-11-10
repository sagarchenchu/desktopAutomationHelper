using System.Windows;

namespace DesktopAutomationWpfSmokeApp;

public partial class ChildDialog : Window
{
    public ChildDialog()
    {
        InitializeComponent();
    }

    private void Ok_Click(object sender, RoutedEventArgs e) => Close();

    private void Cancel_Click(object sender, RoutedEventArgs e) => Close();
}
