using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Automation;

namespace window_tools
{
    /// <summary>
    /// Interaction logic for FindWindowDialog.xaml
    /// </summary>
    public partial class FindWindowDialog : Window
    {
        private List<WindowInfo> _foundWindows = new List<WindowInfo>();
        public WindowInfo? SelectedWindow { get; private set; }

        public FindWindowDialog()
        {
            InitializeComponent();
        }

        private void BtnFind_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _foundWindows.Clear();
                listResults.Items.Clear();

                // 获取桌面根元素
                AutomationElement desktop = AutomationElement.RootElement;
                
                // 获取所有窗口
                var windows = desktop.FindAll(TreeScope.Children, 
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));

                foreach (AutomationElement window in windows)
                {
                    try
                    {
                        var windowInfo = new WindowInfo
                        {
                            AutomationElement = window,
                            Title = window.Current.Name ?? "",
                            ClassName = window.Current.ClassName ?? "",
                            ProcessId = window.Current.ProcessId
                        };

                        // 根据查找条件过滤
                        bool matches = true;

                        if (chkMatchTitle.IsChecked == true && !string.IsNullOrEmpty(txtWindowTitle.Text))
                        {
                            matches = matches && windowInfo.Title.Contains(txtWindowTitle.Text, StringComparison.OrdinalIgnoreCase);
                        }

                        if (chkMatchClassName.IsChecked == true && !string.IsNullOrEmpty(txtClassName.Text))
                        {
                            matches = matches && windowInfo.ClassName.Contains(txtClassName.Text, StringComparison.OrdinalIgnoreCase);
                        }

                        if (chkMatchProcessId.IsChecked == true && !string.IsNullOrEmpty(txtProcessId.Text))
                        {
                            if (int.TryParse(txtProcessId.Text, out int processId))
                            {
                                matches = matches && windowInfo.ProcessId == processId;
                            }
                        }

                        if (matches)
                        {
                            _foundWindows.Add(windowInfo);
                            listResults.Items.Add($"{windowInfo.Title} ({windowInfo.ClassName}) - PID: {windowInfo.ProcessId}");
                        }
                    }
                    catch (Exception ex)
                    {
                        // 忽略无法访问的窗口
                        System.Diagnostics.Debug.WriteLine($"访问窗口时出错: {ex.Message}");
                    }
                }

                if (_foundWindows.Count == 0)
                {
                    MessageBox.Show("未找到匹配的窗口", "查找结果", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找窗口时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ListResults_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            btnSelect.IsEnabled = listResults.SelectedIndex >= 0;
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            if (listResults.SelectedIndex >= 0 && listResults.SelectedIndex < _foundWindows.Count)
            {
                SelectedWindow = _foundWindows[listResults.SelectedIndex];
                DialogResult = true;
                Close();
            }
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    public class WindowInfo
    {
        public AutomationElement AutomationElement { get; set; } = null!;
        public string Title { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public int ProcessId { get; set; }
    }
} 