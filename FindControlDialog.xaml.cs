using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Automation;
using UIAutomationCondition = System.Windows.Automation.Condition;

namespace window_tools
{
    /// <summary>
    /// Interaction logic for FindControlDialog.xaml
    /// </summary>
    public partial class FindControlDialog : Window
    {
        private List<ControlInfo> _foundControls = new List<ControlInfo>();
        public ControlInfo? SelectedControl { get; private set; }

        public FindControlDialog()
        {
            InitializeComponent();
            InitializeControlTypes();
        }

        private void InitializeControlTypes()
        {
            // 添加常用的控件类型
            var controlTypes = new[]
            {
                "Button", "Edit", "Text", "ComboBox", "List", "Tree", "Tab", "Menu", "MenuItem",
                "CheckBox", "RadioButton", "Slider", "ProgressBar", "ScrollBar", "ToolBar",
                "StatusBar", "ToolTip", "Group", "Pane", "Window", "Dialog", "Document"
            };

            foreach (var controlType in controlTypes)
            {
                cboControlType.Items.Add(controlType);
            }
        }

        private void BtnFind_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _foundControls.Clear();
                listResults.Items.Clear();

                // 获取桌面根元素
                AutomationElement desktop = AutomationElement.RootElement;
                
                // 构建查找条件
                var conditions = new List<UIAutomationCondition>();

                if (chkMatchName.IsChecked == true && !string.IsNullOrEmpty(txtControlName.Text))
                {
                    conditions.Add(new PropertyCondition(AutomationElement.NameProperty, txtControlName.Text));
                }

                if (chkMatchAutomationId.IsChecked == true && !string.IsNullOrEmpty(txtAutomationId.Text))
                {
                    conditions.Add(new PropertyCondition(AutomationElement.AutomationIdProperty, txtAutomationId.Text));
                }

                if (chkMatchControlType.IsChecked == true && !string.IsNullOrEmpty(cboControlType.Text))
                {
                    var controlType = GetControlTypeFromName(cboControlType.Text);
                    if (controlType != null)
                    {
                        conditions.Add(new PropertyCondition(AutomationElement.ControlTypeProperty, controlType));
                    }
                }

                if (chkMatchClassName.IsChecked == true && !string.IsNullOrEmpty(txtClassName.Text))
                {
                    conditions.Add(new PropertyCondition(AutomationElement.ClassNameProperty, txtClassName.Text));
                }

                // 如果没有条件，使用默认条件
                if (conditions.Count == 0)
                {
                    conditions.Add(UIAutomationCondition.TrueCondition);
                }

                // 组合条件
                UIAutomationCondition finalCondition = conditions.Count == 1 ? conditions[0] : new AndCondition(conditions.ToArray());

                // 确定搜索范围
                TreeScope searchScope = chkSearchInChildren.IsChecked == true ? 
                    TreeScope.Descendants : TreeScope.Children;

                // 搜索控件
                var elements = desktop.FindAll(searchScope, finalCondition);

                foreach (AutomationElement element in elements)
                {
                    try
                    {
                        var controlInfo = new ControlInfo
                        {
                            AutomationElement = element,
                            Name = element.Current.Name ?? "",
                            AutomationId = element.Current.AutomationId ?? "",
                            ControlType = element.Current.ControlType.ProgrammaticName.Replace("ControlType.", ""),
                            ClassName = element.Current.ClassName ?? "",
                            ProcessId = element.Current.ProcessId
                        };

                        _foundControls.Add(controlInfo);
                        listResults.Items.Add($"{controlInfo.Name} ({controlInfo.ControlType}) - ID: {controlInfo.AutomationId}");
                    }
                    catch (Exception ex)
                    {
                        // 忽略无法访问的控件
                        System.Diagnostics.Debug.WriteLine($"访问控件时出错: {ex.Message}");
                    }
                }

                if (_foundControls.Count == 0)
                {
                    MessageBox.Show("未找到匹配的控件", "查找结果", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"查找控件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private ControlType? GetControlTypeFromName(string controlTypeName)
        {
            return controlTypeName.ToLower() switch
            {
                "button" => ControlType.Button,
                "edit" => ControlType.Edit,
                "text" => ControlType.Text,
                "combobox" => ControlType.ComboBox,
                "list" => ControlType.List,
                "tree" => ControlType.Tree,
                "tab" => ControlType.Tab,
                "menu" => ControlType.Menu,
                "menuitem" => ControlType.MenuItem,
                "checkbox" => ControlType.CheckBox,
                "radiobutton" => ControlType.RadioButton,
                "slider" => ControlType.Slider,
                "progressbar" => ControlType.ProgressBar,
                "scrollbar" => ControlType.ScrollBar,
                "toolbar" => ControlType.ToolBar,
                "statusbar" => ControlType.StatusBar,
                "tooltip" => ControlType.ToolTip,
                "group" => ControlType.Group,
                "pane" => ControlType.Pane,
                "window" => ControlType.Window,
                "dialog" => ControlType.Window, // 对话框通常是Window类型
                "document" => ControlType.Document,
                _ => null
            };
        }

        private void ListResults_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            btnSelect.IsEnabled = listResults.SelectedIndex >= 0;
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            if (listResults.SelectedIndex >= 0 && listResults.SelectedIndex < _foundControls.Count)
            {
                SelectedControl = _foundControls[listResults.SelectedIndex];
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

    public class ControlInfo
    {
        public AutomationElement AutomationElement { get; set; } = null!;
        public string Name { get; set; } = string.Empty;
        public string AutomationId { get; set; } = string.Empty;
        public string ControlType { get; set; } = string.Empty;
        public string ClassName { get; set; } = string.Empty;
        public int ProcessId { get; set; }
    }
} 