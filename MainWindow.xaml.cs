using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace BolOrderExporter;

public partial class MainWindow : Window
{
    private readonly BolApiClient _api = new();
    private readonly ObservableCollection<SavedCredential> _savedCredentials = new();
    private readonly ObservableCollection<CancelledOrder> _cancelledOrders = new();
    private string? _currentUser;
    private string? _currentPassword;
    private string? _loginToken;
    private byte[]? _xlsxBytes;
    private DateOnly? _completedDate;

    public MainWindow()
    {
        InitializeComponent();
        UserComboBox.ItemsSource = _savedCredentials;
        CancelledDataGrid.ItemsSource = _cancelledOrders;
        TargetDatePicker.SelectedDate = DateTime.Today.AddDays(-1);
        LoadCredentials();
        Closed += (_, _) => _api.Dispose();
    }

    private void LoadCredentials()
    {
        try
        {
            _savedCredentials.Clear();
            foreach (var credential in CredentialManager.LoadAll()) _savedCredentials.Add(credential);
            if (_savedCredentials.Count > 0) UserComboBox.SelectedIndex = 0;
        }
        catch (Exception ex)
        {
            LoginMessage.Text = "读取已保存账号失败：" + ex.Message;
        }
    }

    private void UserComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (UserComboBox.SelectedItem is SavedCredential credential)
        {
            UserComboBox.Text = credential.User;
            PasswordTextBox.Text = credential.Password;
        }
    }

    private async void LoginButton_Click(object sender, RoutedEventArgs e)
    {
        var user = UserComboBox.Text.Trim();
        var password = PasswordTextBox.Text;
        if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(password))
        {
            LoginMessage.Text = "请输入 User 和 Password。";
            return;
        }

        SetLoginBusy(true);
        LoginMessage.Foreground = System.Windows.Media.Brushes.SlateGray;
        LoginMessage.Text = "正在登录 bol.com API…";
        try
        {
            var token = await _api.LoginAsync(user, password);
            CredentialManager.Save(user, password);
            _currentUser = user;
            _currentPassword = password;
            _loginToken = token;
            LoginMessage.Text = string.Empty;
            ShowStage(DatePanel, "登录成功，请选择要处理的日期");
            LoadCredentials();
        }
        catch (Exception ex)
        {
            LoginMessage.Foreground = System.Windows.Media.Brushes.Firebrick;
            LoginMessage.Text = "登录失败：" + FriendlyMessage(ex);
        }
        finally
        {
            SetLoginBusy(false);
        }
    }

    private void DeleteCredentialButton_Click(object sender, RoutedEventArgs e)
    {
        var credential = UserComboBox.SelectedItem as SavedCredential;
        if (credential is null)
        {
            LoginMessage.Text = "请先从列表中选择要删除的账号。";
            return;
        }

        var answer = MessageBox.Show(
            $"确定删除 User“{credential.User}”及其对应 Password 吗？",
            "删除账号",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);
        if (answer != MessageBoxResult.Yes) return;

        try
        {
            CredentialManager.Delete(credential.User);
            _savedCredentials.Remove(credential);
            UserComboBox.SelectedItem = null;
            UserComboBox.Text = string.Empty;
            PasswordTextBox.Clear();
            LoginMessage.Text = "该 User 和 Password 已删除。";
        }
        catch (Exception ex)
        {
            LoginMessage.Text = "删除失败：" + ex.Message;
        }
    }

    private async void StartButton_Click(object sender, RoutedEventArgs e)
    {
        if (_currentUser is null || _currentPassword is null || _loginToken is null)
        {
            ShowStage(LoginPanel, "登录店铺 API 后开始处理订单");
            LoginMessage.Text = "登录信息已失效，请重新登录。";
            return;
        }
        if (TargetDatePicker.SelectedDate is not DateTime selected)
        {
            DateMessage.Text = "请选择日期。";
            return;
        }

        DateMessage.Text = string.Empty;
        _xlsxBytes = null;
        _completedDate = null;
        _cancelledOrders.Clear();
        CancelledCountText.Text = "（0）";
        SaveXlsxButton.IsEnabled = false;
        CopyCancelledButton.IsEnabled = false;
        RunErrorText.Text = string.Empty;
        RunProgressBar.IsIndeterminate = false;
        RunProgressBar.Value = 0;
        ProgressPercentText.Text = "0%";
        ShowStage(RunPanel, $"正在处理 {selected:yyyy-MM-dd} 的订单");

        var progress = new Progress<WorkflowProgress>(UpdateProgress);
        try
        {
            var runner = new WorkflowRunner(_api);
            var result = await runner.RunAsync(
                _currentUser,
                _currentPassword,
                _loginToken,
                DateOnly.FromDateTime(selected),
                progress);

            foreach (var order in result.CancelledOrders) _cancelledOrders.Add(order);
            CancelledCountText.Text = $"（{_cancelledOrders.Count}）";
            _xlsxBytes = result.XlsxBytes;
            _completedDate = DateOnly.FromDateTime(selected);
            SaveXlsxButton.IsEnabled = true;
            CopyCancelledButton.IsEnabled = _cancelledOrders.Count > 0;
        }
        catch (Exception ex)
        {
            RunProgressBar.IsIndeterminate = false;
            RunErrorText.Text = "运行失败：" + FriendlyMessage(ex);
            ProgressStatusText.Text = "任务没有完成，XLSX 文件不可用。";
            SaveXlsxButton.IsEnabled = false;
        }
    }

    private void UpdateProgress(WorkflowProgress update)
    {
        RunProgressBar.IsIndeterminate = update.IsIndeterminate;
        if (!update.IsIndeterminate)
        {
            RunProgressBar.Value = Math.Clamp(update.Percentage, 0, 100);
            ProgressPercentText.Text = $"{Math.Round(update.Percentage):0}%";
        }
        else
        {
            ProgressPercentText.Text = "处理中";
        }
        ProgressStatusText.Text = update.Status;
    }

    private void SaveXlsxButton_Click(object sender, RoutedEventArgs e)
    {
        if (_xlsxBytes is null || _completedDate is null) return;
        var dialog = new SaveFileDialog
        {
            Title = "保存 BOL 订单 XLSX",
            Filter = "Excel 工作簿 (*.xlsx)|*.xlsx",
            FileName = $"BOL订单_{_completedDate:yyyy-MM-dd}.xlsx",
            AddExtension = true,
            DefaultExt = ".xlsx"
        };
        if (dialog.ShowDialog(this) != true) return;

        try
        {
            File.WriteAllBytes(dialog.FileName, _xlsxBytes);
            MessageBox.Show(this, "XLSX 文件已保存。", "保存成功", MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, "保存失败：" + ex.Message, "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CopyCancelledButton_Click(object sender, RoutedEventArgs e)
    {
        if (_cancelledOrders.Count == 0) return;
        var text = new StringBuilder("orderId\torderPlacedDateTime\r\n");
        foreach (var order in _cancelledOrders)
            text.Append(order.OrderId).Append('\t').Append(order.OrderPlacedDateTime).Append("\r\n");
        Clipboard.SetText(text.ToString());
        ProgressStatusText.Text = "取消单订单号和下单时间已复制。";
    }

    private void BackToLoginButton_Click(object sender, RoutedEventArgs e) =>
        ShowStage(LoginPanel, "登录店铺 API 后开始处理订单");

    private void ChooseAnotherDateButton_Click(object sender, RoutedEventArgs e) =>
        ShowStage(DatePanel, "登录成功，请选择要处理的日期");

    private void SetLoginBusy(bool busy)
    {
        LoginButton.IsEnabled = !busy;
        DeleteCredentialButton.IsEnabled = !busy;
        UserComboBox.IsEnabled = !busy;
        PasswordTextBox.IsEnabled = !busy;
    }

    private void ShowStage(UIElement stage, string subtitle)
    {
        LoginPanel.Visibility = stage == LoginPanel ? Visibility.Visible : Visibility.Collapsed;
        DatePanel.Visibility = stage == DatePanel ? Visibility.Visible : Visibility.Collapsed;
        RunPanel.Visibility = stage == RunPanel ? Visibility.Visible : Visibility.Collapsed;
        HeaderSubtitle.Text = subtitle;
    }

    private static string FriendlyMessage(Exception ex)
    {
        var root = ex;
        while (root.InnerException is not null) root = root.InnerException;
        return root.Message;
    }
}
