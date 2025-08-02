using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Automation;
using System.Windows.Automation.Provider;
using System.Windows.Forms;
using System.Drawing.Imaging;
using UIAutomationCondition = System.Windows.Automation.Condition;

namespace window_tools
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Win32 API 声明
        [DllImport("user32.dll")]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, uint dwExtraInfo);

        [DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point pt);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out System.Drawing.Point lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const uint MOUSEEVENTF_LEFTUP = 0x0004;
        private const uint MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const uint MOUSEEVENTF_RIGHTUP = 0x0010;

        private const int WH_MOUSE_LL = 14;
        private const int WM_LBUTTONDOWN = 0x0201;

        private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT
        {
            public int x;
            public int y;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        private ObservableCollection<UIElementInfo> _uiElements;
        private UIElementInfo? _selectedElement;
        private bool _isHighlighting = false;
        private HighlightWindow? _currentHighlightWindow; // 添加当前高亮窗口的引用
        
        // 点击获取句柄相关字段
        private IntPtr _mouseHook = IntPtr.Zero;
        private LowLevelMouseProc _mouseProc;
        private bool _isPickingElement = false;

        public MainWindow()
        {
            InitializeComponent();
            _uiElements = new ObservableCollection<UIElementInfo>();
            _mouseProc = MouseHookProc; // 初始化鼠标钩子委托
            InitializeUI();
        }

        private void InitializeUI()
        {
            // 设置数据绑定
            treeViewControls.ItemsSource = _uiElements;
            
            // 初始化状态
            UpdateStatus("就绪");
            UpdateElementCount(0);
            
            // 加载桌面元素
            LoadDesktopElements();
        }

        private void LoadDesktopElements()
        {
            try
            {
                UpdateStatus("正在加载桌面元素...");
                
                // 获取桌面根元素
                AutomationElement desktop = AutomationElement.RootElement;
                
                // 清空现有数据
                _uiElements.Clear();
                
                // 获取所有顶级窗口
                var windows = desktop.FindAll(TreeScope.Children, 
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window));
                
                foreach (AutomationElement window in windows)
                {
                    try
                    {
                        var windowInfo = CreateUIElementInfo(window);
                        if (windowInfo != null)
                        {
                            _uiElements.Add(windowInfo);
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"加载窗口时出错: {ex.Message}");
                    }
                }
                
                UpdateElementCount(_uiElements.Count);
                UpdateStatus($"已加载 {_uiElements.Count} 个顶级窗口");
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载桌面元素时出错: {ex.Message}");
                MessageBox.Show($"加载桌面元素时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private UIElementInfo? CreateUIElementInfo(AutomationElement element)
        {
            try
            {
                if (element == null) return null;

                var elementInfo = new UIElementInfo
                {
                    AutomationElement = element,
                    DisplayName = GetElementName(element),
                    ControlType = GetControlTypeName(element),
                    Children = new ObservableCollection<UIElementInfo>()
                };

                // 获取子元素
                try
                {
                    var children = element.FindAll(TreeScope.Children, UIAutomationCondition.TrueCondition);
                    foreach (AutomationElement child in children)
                    {
                        var childInfo = CreateUIElementInfo(child);
                        if (childInfo != null)
                        {
                            elementInfo.Children.Add(childInfo);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"获取子元素时出错: {ex.Message}");
                }

                return elementInfo;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"创建UI元素信息时出错: {ex.Message}");
                return null;
            }
        }

        private string GetElementName(AutomationElement element)
        {
            try
            {
                if (element == null) return string.Empty;
                
                string? name = element.Current.Name;
                if (string.IsNullOrEmpty(name))
                {
                    name = element.Current.AutomationId;
                }
                if (string.IsNullOrEmpty(name))
                {
                    name = element.Current.ClassName;
                }
                return name ?? string.Empty; // 返回空字符串而不是"未命名"
            }
            catch
            {
                return string.Empty; // 异常时也返回空字符串
            }
        }

        private string GetControlTypeName(AutomationElement element)
        {
            try
            {
                if (element == null) return "Unknown";
                
                var controlType = element.Current.ControlType;
                if (controlType != null)
                {
                    return controlType.ProgrammaticName.Replace("ControlType.", "");
                }
                return "Unknown";
            }
            catch
            {
                return "Unknown";
            }
        }

        private void UpdateStatus(string message)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = message;
            });
        }

        private void UpdateElementCount(int count)
        {
            Dispatcher.Invoke(() =>
            {
                txtElementCount.Text = $"控件数量: {count}";
            });
        }

        // 事件处理方法
        private void BtnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadDesktopElements();
        }

        private void BtnHighlight_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                _isHighlighting = !_isHighlighting;
                btnHighlight.Content = _isHighlighting ? "取消高亮" : "高亮显示";
                
                if (_isHighlighting)
                {
                    HighlightElement(_selectedElement);
                }
                else
                {
                    RemoveHighlight();
                }
            }
        }

        private void BtnCapture_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                CaptureElement(_selectedElement);
            }
        }

        private void BtnFindWindow_Click(object sender, RoutedEventArgs e)
        {
            ShowFindWindowDialog();
        }

        private void BtnFindControl_Click(object sender, RoutedEventArgs e)
        {
            ShowFindControlDialog();
        }

        // 新增功能方法
        private void BtnFocus_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                SetElementFocus(_selectedElement);
            }
        }

        private void BtnContextMenu_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                ShowContextMenu(_selectedElement);
            }
        }

        private void BtnSetValue_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                ShowSetValueDialog(_selectedElement);
            }
        }

        private void BtnInvoke_Click(object sender, RoutedEventArgs e)
        {
            if (_selectedElement != null)
            {
                InvokeElement(_selectedElement);
            }
        }

        private void BtnPickElement_Click(object sender, RoutedEventArgs e)
        {
            if (_isPickingElement)
            {
                StopPickingElement();
            }
            else
            {
                StartPickingElement();
            }
        }



        private void TreeViewControls_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (e.NewValue is UIElementInfo elementInfo)
            {
                _selectedElement = elementInfo;
                LoadElementProperties(elementInfo);
            }
        }

        private void LoadElementProperties(UIElementInfo elementInfo)
        {
            try
            {
                var properties = new ObservableCollection<PropertyInfo>();
                var patterns = new ObservableCollection<PatternInfo>();

                if (elementInfo.AutomationElement != null)
                {
                    var element = elementInfo.AutomationElement;

                    // 基本属性
                    properties.Add(new PropertyInfo { Name = "名称", Value = element.Current.Name ?? "" });
                    properties.Add(new PropertyInfo { Name = "控件类型", Value = element.Current.ControlType.ProgrammaticName });
                    properties.Add(new PropertyInfo { Name = "本地化控件类型", Value = element.Current.LocalizedControlType });
                    properties.Add(new PropertyInfo { Name = "自动化ID", Value = element.Current.AutomationId ?? "" });
                    properties.Add(new PropertyInfo { Name = "类名", Value = element.Current.ClassName ?? "" });
                    properties.Add(new PropertyInfo { Name = "进程ID", Value = element.Current.ProcessId.ToString() });
                    properties.Add(new PropertyInfo { Name = "运行时ID", Value = string.Join(", ", element.GetRuntimeId()) });
                    properties.Add(new PropertyInfo { Name = "原生窗口句柄", Value = element.Current.NativeWindowHandle.ToString("X") });
                    properties.Add(new PropertyInfo { Name = "框架ID", Value = element.Current.FrameworkId ?? "" });
                    properties.Add(new PropertyInfo { Name = "是否启用", Value = element.Current.IsEnabled.ToString() });
                    properties.Add(new PropertyInfo { Name = "是否可见", Value = element.Current.IsOffscreen.ToString() });
                    properties.Add(new PropertyInfo { Name = "是否可聚焦", Value = element.Current.IsKeyboardFocusable.ToString() });
                    properties.Add(new PropertyInfo { Name = "是否已聚焦", Value = element.Current.HasKeyboardFocus.ToString() });
                    properties.Add(new PropertyInfo { Name = "访问键", Value = element.Current.AccessKey ?? "" });
                    properties.Add(new PropertyInfo { Name = "帮助文本", Value = element.Current.HelpText ?? "" });

                    // 边界矩形
                    var rect = element.Current.BoundingRectangle;
                    properties.Add(new PropertyInfo { Name = "边界矩形", Value = $"X:{rect.X}, Y:{rect.Y}, W:{rect.Width}, H:{rect.Height}" });

                    // 检查是否为密码字段
                    try
                    {
                        var valuePattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (valuePattern != null)
                        {
                            // 注意：IsPassword属性在某些版本中可能不可用
                            properties.Add(new PropertyInfo { Name = "值模式支持", Value = "是" });
                        }
                    }
                    catch { }

                    // 模式支持
                    CheckPatternSupport(element, patterns);

                    // 详细信息
                    var details = new StringBuilder();
                    details.AppendLine($"控件信息:");
                    details.AppendLine($"  名称: {element.Current.Name}");
                    details.AppendLine($"  类型: {element.Current.ControlType.ProgrammaticName}");
                    details.AppendLine($"  自动化ID: {element.Current.AutomationId}");
                    details.AppendLine($"  类名: {element.Current.ClassName}");
                    details.AppendLine($"  进程ID: {element.Current.ProcessId}");
                    details.AppendLine($"  窗口句柄: 0x{element.Current.NativeWindowHandle:X}");
                    details.AppendLine($"  框架: {element.Current.FrameworkId}");
                    details.AppendLine($"  边界: {rect}");
                    details.AppendLine($"  启用: {element.Current.IsEnabled}");
                    details.AppendLine($"  可见: {!element.Current.IsOffscreen}");
                    details.AppendLine($"  可聚焦: {element.Current.IsKeyboardFocusable}");
                    details.AppendLine($"  已聚焦: {element.Current.HasKeyboardFocus}");

                    txtDetails.Text = details.ToString();
                }

                dataGridProperties.ItemsSource = properties;
                dataGridPatterns.ItemsSource = patterns;
            }
            catch (Exception ex)
            {
                UpdateStatus($"加载属性时出错: {ex.Message}");
            }
        }

        private void CheckPatternSupport(AutomationElement element, ObservableCollection<PatternInfo> patterns)
        {
            var patternTypes = new[]
            {
                new { Name = "Invoke", Pattern = InvokePattern.Pattern },
                new { Name = "Value", Pattern = ValuePattern.Pattern },
                new { Name = "Text", Pattern = TextPattern.Pattern },
                new { Name = "Selection", Pattern = SelectionPattern.Pattern },
                new { Name = "SelectionItem", Pattern = SelectionItemPattern.Pattern },
                new { Name = "Scroll", Pattern = ScrollPattern.Pattern },
                new { Name = "ScrollItem", Pattern = ScrollItemPattern.Pattern },
                new { Name = "ExpandCollapse", Pattern = ExpandCollapsePattern.Pattern },
                new { Name = "Toggle", Pattern = TogglePattern.Pattern },
                new { Name = "Transform", Pattern = TransformPattern.Pattern },
                new { Name = "Window", Pattern = WindowPattern.Pattern },
                new { Name = "Dock", Pattern = DockPattern.Pattern },
                new { Name = "Grid", Pattern = GridPattern.Pattern },
                new { Name = "GridItem", Pattern = GridItemPattern.Pattern },
                new { Name = "Table", Pattern = TablePattern.Pattern },
                new { Name = "TableItem", Pattern = TableItemPattern.Pattern },
                new { Name = "RangeValue", Pattern = RangeValuePattern.Pattern },
                new { Name = "MultipleView", Pattern = MultipleViewPattern.Pattern },
                new { Name = "SynchronizedInput", Pattern = SynchronizedInputPattern.Pattern },
                new { Name = "VirtualizedItem", Pattern = VirtualizedItemPattern.Pattern },
                new { Name = "ItemContainer", Pattern = ItemContainerPattern.Pattern }
            };

            foreach (var patternType in patternTypes)
            {
                try
                {
                    bool isSupported = element.GetCurrentPattern(patternType.Pattern) != null;
                    patterns.Add(new PatternInfo { Name = patternType.Name, IsSupported = isSupported });
                }
                catch
                {
                    patterns.Add(new PatternInfo { Name = patternType.Name, IsSupported = false });
                }
            }
        }

        private void HighlightElement(UIElementInfo elementInfo)
        {
            try
            {
                // 先移除之前的高亮
                RemoveHighlight();

                var element = elementInfo.AutomationElement;
                var rect = element.Current.BoundingRectangle;
                
                // 检查边界矩形是否有效
                if (rect.Width <= 0 || rect.Height <= 0)
                {
                    UpdateStatus("无法高亮显示：控件边界无效");
                    return;
                }

                // 检查控件是否离屏
                if (element.Current.IsOffscreen)
                {
                    UpdateStatus("无法高亮显示：控件不在屏幕上");
                    return;
                }

                // 使用新的精确坐标方法创建高亮窗口
                _currentHighlightWindow = HighlightWindow.CreateFromAutomationRect(rect);

                if (_currentHighlightWindow != null)
                {
                    _currentHighlightWindow.Show();
                    UpdateStatus($"已高亮显示控件: {elementInfo.DisplayName}");
                }
                else
                {
                    UpdateStatus("无法创建高亮窗口");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"高亮显示时出错: {ex.Message}");
                Debug.WriteLine($"高亮显示异常: {ex}");
                // 不显示错误对话框，避免干扰用户体验
            }
        }

        private void RemoveHighlight()
        {
            try
            {
                if (_currentHighlightWindow != null)
                {
                    // 检查窗口是否仍然有效
                    if (_currentHighlightWindow.IsLoaded)
                    {
                        _currentHighlightWindow.Close();
                    }
                    _currentHighlightWindow = null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"移除高亮时出错: {ex.Message}");
                _currentHighlightWindow = null;
            }
        }

        private void CaptureElement(UIElementInfo elementInfo)
        {
            try
            {
                var rect = elementInfo.AutomationElement.Current.BoundingRectangle;
                if (rect.Width > 0 && rect.Height > 0)
                {
                    using (var bitmap = new Bitmap((int)rect.Width, (int)rect.Height))
                    using (var graphics = Graphics.FromImage(bitmap))
                    {
                        graphics.CopyFromScreen((int)rect.X, (int)rect.Y, 0, 0, bitmap.Size);
                        
                        var fileName = $"capture_{DateTime.Now:yyyyMMdd_HHmmss}.png";
                        bitmap.Save(fileName, ImageFormat.Png);
                        
                        UpdateStatus($"截图已保存: {fileName}");
                        MessageBox.Show($"截图已保存为: {fileName}", "截图完成", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"截图时出错: {ex.Message}");
            }
        }

        private void ShowFindWindowDialog()
        {
            var dialog = new FindWindowDialog();
            if (dialog.ShowDialog() == true)
            {
                // 实现窗口查找逻辑
                UpdateStatus("窗口查找功能待实现");
            }
        }

        private void ShowFindControlDialog()
        {
            var dialog = new FindControlDialog();
            if (dialog.ShowDialog() == true && dialog.SelectedControl != null)
            {
                try
                {
                    var selectedControl = dialog.SelectedControl;
                    UpdateStatus($"正在定位控件: {selectedControl.Name}");
                    
                    // 检查是否为自己的应用程序控件
                    if (IsOwnApplicationElement(selectedControl.AutomationElement))
                    {
                        UpdateStatus("不能检查自己程序的控件");
                        MessageBox.Show("不能检查自己程序的控件", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        return;
                    }
                    
                    // 首先尝试在现有树中查找
                    var foundElement = FindElementInTreeByAutomationElement(selectedControl.AutomationElement);
                    
                    if (foundElement != null)
                    {
                        // 直接选中找到的元素
                        SelectAndExpandToElement(foundElement);
                        UpdateStatus($"已定位到控件: {foundElement.DisplayName}");
                    }
                    else
                    {
                        // 如果找不到，刷新控件树并再次查找
                        UpdateStatus("正在刷新控件树...");
                        LoadDesktopElements();
                        
                        // 延迟查找，等待树加载完成
                        Dispatcher.BeginInvoke(new Action(() =>
                        {
                            var retryFoundElement = FindElementInTreeByAutomationElement(selectedControl.AutomationElement);
                            if (retryFoundElement != null)
                            {
                                SelectAndExpandToElement(retryFoundElement);
                                UpdateStatus($"已定位到控件: {retryFoundElement.DisplayName}");
                            }
                            else
                            {
                                // 如果仍然找不到，创建一个临时的元素信息来显示
                                CreateAndShowTemporaryElement(selectedControl.AutomationElement);
                                UpdateStatus($"已获取控件信息: {selectedControl.Name} (未在树中找到对应节点)");
                            }
                        }), System.Windows.Threading.DispatcherPriority.Background);
                    }
                }
                catch (Exception ex)
                {
                    UpdateStatus($"定位控件时出错: {ex.Message}");
                    MessageBox.Show($"定位控件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void SetElementFocus(UIElementInfo elementInfo)
        {
            try
            {
                var element = elementInfo.AutomationElement;
                
                // 检查控件是否仍然有效
                if (element == null)
                {
                    UpdateStatus("控件已失效，请重新选择");
                    MessageBox.Show("控件已失效，请重新选择", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 检查控件是否启用
                if (!element.Current.IsEnabled)
                {
                    UpdateStatus("控件已禁用，无法设置焦点");
                    MessageBox.Show("控件已禁用，无法设置焦点", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 检查是否支持键盘焦点
                if (element.Current.IsKeyboardFocusable)
                {
                    try
                    {
                        element.SetFocus();
                        UpdateStatus($"已设置焦点: {elementInfo.DisplayName}");
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许设置焦点");
                        MessageBox.Show("控件当前状态不允许设置焦点，请确保控件处于可交互状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    // 尝试使用Legacy Accessibility
                    try
                    {
                        var legacyPattern = element.GetCurrentPattern(AutomationPattern.LookupById(10018)) as object;
                        if (legacyPattern != null)
                        {
                            // 使用反射调用DoDefaultAction方法
                            var doDefaultActionMethod = legacyPattern.GetType().GetMethod("DoDefaultAction");
                            if (doDefaultActionMethod != null)
                            {
                                doDefaultActionMethod.Invoke(legacyPattern, null);
                                UpdateStatus($"已通过Legacy Accessibility设置焦点: {elementInfo.DisplayName}");
                            }
                        }
                        else
                        {
                            UpdateStatus("该控件不支持焦点设置");
                            MessageBox.Show("该控件不支持焦点设置", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许设置焦点");
                        MessageBox.Show("控件当前状态不允许设置焦点，请确保控件处于可交互状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch
                    {
                        UpdateStatus("该控件不支持焦点设置");
                        MessageBox.Show("该控件不支持焦点设置", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"设置焦点时出错: {ex.Message}");
                Debug.WriteLine($"设置焦点异常: {ex}");
                MessageBox.Show($"设置焦点失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ShowContextMenu(UIElementInfo elementInfo)
        {
            try
            {
                var element = elementInfo.AutomationElement;
                var rect = element.Current.BoundingRectangle;
                
                // 模拟右键点击显示上下文菜单
                var point = new System.Drawing.Point((int)(rect.X + rect.Width / 2), (int)(rect.Y + rect.Height / 2));
                
                // 使用Win32 API模拟右键点击
                SetCursorPos(point.X, point.Y);
                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                
                UpdateStatus($"已显示上下文菜单: {elementInfo.DisplayName}");
            }
            catch (Exception ex)
            {
                UpdateStatus($"显示上下文菜单时出错: {ex.Message}");
                MessageBox.Show($"显示上下文菜单失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void ShowSetValueDialog(UIElementInfo elementInfo)
        {
            try
            {
                var element = elementInfo.AutomationElement;
                
                // 检查控件是否仍然有效
                if (element == null)
                {
                    UpdateStatus("控件已失效，请重新选择");
                    MessageBox.Show("控件已失效，请重新选择", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 检查控件是否启用
                if (!element.Current.IsEnabled)
                {
                    UpdateStatus("控件已禁用，无法设置值");
                    MessageBox.Show("控件已禁用，无法设置值", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 检查是否支持值模式
                var valuePattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                if (valuePattern != null)
                {
                    try
                    {
                        // 检查值模式是否只读
                        if (valuePattern.Current.IsReadOnly)
                        {
                            UpdateStatus("控件为只读，无法设置值");
                            MessageBox.Show("控件为只读，无法设置值", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                            return;
                        }

                        string currentValue = valuePattern.Current.Value ?? "";
                        string newValue = Microsoft.VisualBasic.Interaction.InputBox(
                            $"为控件 '{elementInfo.DisplayName}' 设置新值:",
                            "设置值",
                            currentValue);
                        
                        if (!string.IsNullOrEmpty(newValue))
                        {
                            // 再次检查控件状态
                            if (element.Current.IsEnabled && !valuePattern.Current.IsReadOnly)
                            {
                                valuePattern.SetValue(newValue);
                                UpdateStatus($"已设置值: {elementInfo.DisplayName} = '{newValue}'");
                            }
                            else
                            {
                                UpdateStatus("控件状态已改变，无法设置值");
                                MessageBox.Show("控件状态已改变，无法设置值", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                            }
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许设置值");
                        MessageBox.Show("控件当前状态不允许设置值，请确保控件处于可编辑状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    // 尝试使用Legacy Accessibility
                    try
                    {
                        var legacyPattern = element.GetCurrentPattern(AutomationPattern.LookupById(10018)) as object;
                        if (legacyPattern != null)
                        {
                            // 使用反射获取当前值
                            var valueProperty = legacyPattern.GetType().GetProperty("Current");
                            if (valueProperty != null)
                            {
                                var current = valueProperty.GetValue(legacyPattern);
                                var valueProperty2 = current?.GetType().GetProperty("Value");
                                string currentValue = valueProperty2?.GetValue(current)?.ToString() ?? "";
                                
                                string newValue = Microsoft.VisualBasic.Interaction.InputBox(
                                    $"为控件 '{elementInfo.DisplayName}' 设置新值:",
                                    "设置值 (Legacy)",
                                    currentValue);
                                
                                if (!string.IsNullOrEmpty(newValue))
                                {
                                    // 使用反射调用SetValue方法
                                    var setValueMethod = legacyPattern.GetType().GetMethod("SetValue");
                                    if (setValueMethod != null)
                                    {
                                        setValueMethod.Invoke(legacyPattern, new object[] { newValue });
                                        UpdateStatus($"已通过Legacy Accessibility设置值: {elementInfo.DisplayName} = '{newValue}'");
                                    }
                                }
                            }
                        }
                        else
                        {
                            UpdateStatus("该控件不支持值设置");
                            MessageBox.Show("该控件不支持值设置", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许设置值");
                        MessageBox.Show("控件当前状态不允许设置值，请确保控件处于可编辑状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch
                    {
                        UpdateStatus("该控件不支持值设置");
                        MessageBox.Show("该控件不支持值设置", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"设置值时出错: {ex.Message}");
                Debug.WriteLine($"设置值异常: {ex}");
                MessageBox.Show($"设置值失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void InvokeElement(UIElementInfo elementInfo)
        {
            try
            {
                var element = elementInfo.AutomationElement;
                
                // 检查控件是否仍然有效
                if (element == null)
                {
                    UpdateStatus("控件已失效，请重新选择");
                    MessageBox.Show("控件已失效，请重新选择", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 检查控件是否启用
                if (!element.Current.IsEnabled)
                {
                    UpdateStatus("控件已禁用，无法调用");
                    MessageBox.Show("控件已禁用，无法调用", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                
                // 尝试使用Invoke模式
                var invokePattern = element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
                if (invokePattern != null)
                {
                    try
                    {
                        invokePattern.Invoke();
                        UpdateStatus($"已调用控件: {elementInfo.DisplayName}");
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许调用");
                        MessageBox.Show("控件当前状态不允许调用，请确保控件处于可交互状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
                else
                {
                    // 尝试使用Legacy Accessibility
                    try
                    {
                        var legacyPattern = element.GetCurrentPattern(AutomationPattern.LookupById(10018)) as object;
                        if (legacyPattern != null)
                        {
                            // 使用反射调用DoDefaultAction方法
                            var doDefaultActionMethod = legacyPattern.GetType().GetMethod("DoDefaultAction");
                            if (doDefaultActionMethod != null)
                            {
                                doDefaultActionMethod.Invoke(legacyPattern, null);
                                UpdateStatus($"已通过Legacy Accessibility调用控件: {elementInfo.DisplayName}");
                            }
                        }
                        else
                        {
                            UpdateStatus("该控件不支持调用操作");
                            MessageBox.Show("该控件不支持调用操作", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        UpdateStatus("控件当前状态不允许调用");
                        MessageBox.Show("控件当前状态不允许调用，请确保控件处于可交互状态", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                    catch
                    {
                        UpdateStatus("该控件不支持调用操作");
                        MessageBox.Show("该控件不支持调用操作", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"调用控件时出错: {ex.Message}");
                Debug.WriteLine($"调用控件异常: {ex}");
                MessageBox.Show($"调用控件失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // 复制和导出功能
        private void CopySelectedProperty_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridProperties.SelectedItem is PropertyInfo property)
            {
                try
                {
                    Clipboard.SetText($"{property.Name}: {property.Value}");
                    UpdateStatus($"已复制属性: {property.Name}");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"复制属性失败: {ex.Message}");
                }
            }
        }

        private void CopySelectedPropertyValue_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridProperties.SelectedItem is PropertyInfo property)
            {
                try
                {
                    Clipboard.SetText(property.Value);
                    UpdateStatus($"已复制属性值: {property.Value}");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"复制属性值失败: {ex.Message}");
                }
            }
        }

        private void CopyAllProperties_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var properties = dataGridProperties.ItemsSource as ObservableCollection<PropertyInfo>;
                if (properties != null && properties.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (var property in properties)
                    {
                        sb.AppendLine($"{property.Name}: {property.Value}");
                    }
                    Clipboard.SetText(sb.ToString());
                    UpdateStatus($"已复制 {properties.Count} 个属性");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制所有属性失败: {ex.Message}");
            }
        }

        private void CopyAllPropertyValues_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var properties = dataGridProperties.ItemsSource as ObservableCollection<PropertyInfo>;
                if (properties != null && properties.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (var property in properties)
                    {
                        sb.AppendLine(property.Value);
                    }
                    Clipboard.SetText(sb.ToString());
                    UpdateStatus($"已复制 {properties.Count} 个属性值");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制所有属性值失败: {ex.Message}");
            }
        }

        private void CopySelectedPattern_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridPatterns.SelectedItem is PatternInfo pattern)
            {
                try
                {
                    Clipboard.SetText($"{pattern.Name}: {(pattern.IsSupported ? "支持" : "不支持")}");
                    UpdateStatus($"已复制模式: {pattern.Name}");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"复制模式失败: {ex.Message}");
                }
            }
        }

        private void CopySelectedPatternStatus_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridPatterns.SelectedItem is PatternInfo pattern)
            {
                try
                {
                    Clipboard.SetText(pattern.IsSupported ? "支持" : "不支持");
                    UpdateStatus($"已复制模式状态: {(pattern.IsSupported ? "支持" : "不支持")}");
                }
                catch (Exception ex)
                {
                    UpdateStatus($"复制模式状态失败: {ex.Message}");
                }
            }
        }

        private void CopyAllPatterns_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var patterns = dataGridPatterns.ItemsSource as ObservableCollection<PatternInfo>;
                if (patterns != null && patterns.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (var pattern in patterns)
                    {
                        sb.AppendLine($"{pattern.Name}: {(pattern.IsSupported ? "支持" : "不支持")}");
                    }
                    Clipboard.SetText(sb.ToString());
                    UpdateStatus($"已复制 {patterns.Count} 个模式");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制所有模式失败: {ex.Message}");
            }
        }

        private void CopyAllPatternStatus_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var patterns = dataGridPatterns.ItemsSource as ObservableCollection<PatternInfo>;
                if (patterns != null && patterns.Count > 0)
                {
                    var sb = new StringBuilder();
                    foreach (var pattern in patterns)
                    {
                        sb.AppendLine(pattern.IsSupported ? "支持" : "不支持");
                    }
                    Clipboard.SetText(sb.ToString());
                    UpdateStatus($"已复制 {patterns.Count} 个模式状态");
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"复制所有模式状态失败: {ex.Message}");
            }
        }

        private void ExportPropertiesAsJson_Click(object sender, RoutedEventArgs e)
        {
            ExportToJson(dataGridProperties.ItemsSource as ObservableCollection<PropertyInfo> ?? new ObservableCollection<PropertyInfo>(), "properties");
        }

        private void ExportPropertiesAsTxt_Click(object sender, RoutedEventArgs e)
        {
            ExportToTxt(dataGridProperties.ItemsSource as ObservableCollection<PropertyInfo> ?? new ObservableCollection<PropertyInfo>(), "properties");
        }

        private void ExportPatternsAsJson_Click(object sender, RoutedEventArgs e)
        {
            ExportToJson(dataGridPatterns.ItemsSource as ObservableCollection<PatternInfo> ?? new ObservableCollection<PatternInfo>(), "patterns");
        }

        private void ExportPatternsAsTxt_Click(object sender, RoutedEventArgs e)
        {
            ExportToTxt(dataGridPatterns.ItemsSource as ObservableCollection<PatternInfo> ?? new ObservableCollection<PatternInfo>(), "patterns");
        }

        private void ExportToJson<T>(ObservableCollection<T> items, string type)
        {
            try
            {
                if (items == null || items.Count == 0)
                {
                    MessageBox.Show("没有数据可导出", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var json = System.Text.Json.JsonSerializer.Serialize(items, new System.Text.Json.JsonSerializerOptions 
                { 
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                });

                // 使用SaveFileDialog选择保存位置
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "保存JSON文件",
                    Filter = "JSON文件 (*.json)|*.json|所有文件 (*.*)|*.*",
                    DefaultExt = "json",
                    FileName = $"{type}_{DateTime.Now:yyyyMMdd_HHmmss}.json"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveFileDialog.FileName, json, Encoding.UTF8);
                    UpdateStatus($"已导出为JSON: {saveFileDialog.FileName}");
                    MessageBox.Show($"已导出为: {saveFileDialog.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"导出JSON失败: {ex.Message}");
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportToTxt<T>(ObservableCollection<T> items, string type)
        {
            try
            {
                if (items == null || items.Count == 0)
                {
                    MessageBox.Show("没有数据可导出", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var sb = new StringBuilder();
                sb.AppendLine($"UI Inspector 导出报告");
                sb.AppendLine($"导出时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"数据类型: {type}");
                sb.AppendLine(new string('-', 50));

                foreach (var item in items)
                {
                    if (item is PropertyInfo property)
                    {
                        sb.AppendLine($"{property.Name}: {property.Value}");
                    }
                    else if (item is PatternInfo pattern)
                    {
                        sb.AppendLine($"{pattern.Name}: {(pattern.IsSupported ? "支持" : "不支持")}");
                    }
                }

                // 使用SaveFileDialog选择保存位置
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "保存TXT文件",
                    Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                    DefaultExt = "txt",
                    FileName = $"{type}_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    File.WriteAllText(saveFileDialog.FileName, sb.ToString(), Encoding.UTF8);
                    UpdateStatus($"已导出为TXT: {saveFileDialog.FileName}");
                    MessageBox.Show($"已导出为: {saveFileDialog.FileName}", "导出成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"导出TXT失败: {ex.Message}");
                MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // 点击获取句柄功能实现
        private void StartPickingElement()
        {
            try
            {
                _isPickingElement = true;
                btnPickElement.Content = "取消获取";
                btnPickElement.Background = new SolidColorBrush(Colors.LightCoral);
                
                // 最小化窗口
                this.WindowState = WindowState.Minimized;
                
                // 安装鼠标钩子
                _mouseHook = SetWindowsHookEx(WH_MOUSE_LL, _mouseProc, GetModuleHandle("user32"), 0);
                
                if (_mouseHook == IntPtr.Zero)
                {
                    UpdateStatus("无法安装鼠标钩子");
                    StopPickingElement();
                    return;
                }
                
                UpdateStatus("窗口已最小化，请点击要检查的控件...（点击后窗口将保持最小化状态）");
            }
            catch (Exception ex)
            {
                UpdateStatus($"启动点击获取功能时出错: {ex.Message}");
                StopPickingElement();
            }
        }

        private void StopPickingElement()
        {
            try
            {
                _isPickingElement = false;
                btnPickElement.Content = "点击获取句柄";
                btnPickElement.Background = new SolidColorBrush(Colors.LightYellow);
                
                // 卸载鼠标钩子
                if (_mouseHook != IntPtr.Zero)
                {
                    UnhookWindowsHookEx(_mouseHook);
                    _mouseHook = IntPtr.Zero;
                }
                
                // 不恢复窗口状态，保持最小化
                // this.WindowState = WindowState.Maximized;
                // this.Activate();
                
                UpdateStatus("已停止点击获取功能，窗口保持最小化状态");
            }
            catch (Exception ex)
            {
                UpdateStatus($"停止点击获取功能时出错: {ex.Message}");
            }
        }

        private IntPtr MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                if (nCode >= 0 && _isPickingElement && wParam == (IntPtr)WM_LBUTTONDOWN)
                {
                    // 获取鼠标位置
                    var hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
                    var point = new System.Drawing.Point(hookStruct.pt.x, hookStruct.pt.y);
                    
                    // 在UI线程中处理点击
                    Dispatcher.BeginInvoke(new Action(() => HandleElementClick(point)));
                    
                    // 阻止点击事件传递
                    return (IntPtr)1;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"鼠标钩子处理异常: {ex.Message}");
            }
            
            return CallNextHookEx(_mouseHook, nCode, wParam, lParam);
        }

        private void HandleElementClick(System.Drawing.Point point)
        {
            try
            {
                // 停止点击获取模式
                StopPickingElement();
                
                // 首先尝试从点击位置获取具体的UI元素
                var element = AutomationElement.FromPoint(new System.Windows.Point(point.X, point.Y));
                if (element == null)
                {
                    UpdateStatus("无法获取点击位置的UI元素");
                    return;
                }
                
                // 检查是否点击了自己的应用程序
                if (IsOwnApplicationElement(element))
                {
                    UpdateStatus("不能检查自己程序的控件，请点击其他应用程序的控件");
                    MessageBox.Show("不能检查自己程序的控件，请点击其他应用程序的控件", "提示", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                // 获取点击元素的根窗口
                var rootWindow = GetRootWindow(element);
                if (rootWindow == null)
                {
                    UpdateStatus("无法获取根窗口");
                    return;
                }
                
                // 再次检查根窗口是否为自己的应用程序
                if (IsOwnApplicationElement(rootWindow))
                {
                    UpdateStatus("不能检查自己程序的窗口，请点击其他应用程序的窗口");
                    MessageBox.Show("不能检查自己程序的窗口，请点击其他应用程序的窗口", "提示", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                
                UpdateStatus($"正在查找控件: {GetElementName(element)}");
                
                // 询问用户是否要恢复窗口
                var result = MessageBox.Show(
                    $"已获取到控件: {GetElementName(element)}\n\n是否要恢复窗口以查看详细信息？", 
                    "获取成功", 
                    MessageBoxButton.YesNo, 
                    MessageBoxImage.Question);
                
                if (result == MessageBoxResult.Yes)
                {
                    // 恢复窗口
                    this.WindowState = WindowState.Maximized;
                    this.Activate();
                }
                
                // 首先尝试在现有树中查找
                var foundElement = FindElementInTreeByAutomationElement(element);
                
                if (foundElement != null)
                {
                    // 直接选中找到的元素
                    SelectAndExpandToElement(foundElement);
                    UpdateStatus($"已定位到控件: {foundElement.DisplayName}");
                }
                else
                {
                    // 如果找不到，可能需要展开更多节点或刷新树
                    UpdateStatus("正在刷新控件树...");
                    LoadDesktopElements();
                    
                    // 延迟查找，等待树加载完成
                    Dispatcher.BeginInvoke(new Action(() =>
                    {
                        var retryFoundElement = FindElementInTreeByAutomationElement(element);
                        if (retryFoundElement != null)
                        {
                            SelectAndExpandToElement(retryFoundElement);
                            UpdateStatus($"已定位到控件: {retryFoundElement.DisplayName}");
                        }
                        else
                        {
                            // 尝试创建一个临时的元素信息来显示
                            CreateAndShowTemporaryElement(element);
                        }
                    }), System.Windows.Threading.DispatcherPriority.Background);
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"处理点击事件时出错: {ex.Message}");
                Debug.WriteLine($"处理点击事件异常: {ex}");
            }
        }

        private UIElementInfo? FindElementInTreeByAutomationElement(AutomationElement targetElement)
        {
            try
            {
                if (targetElement == null) return null;
                
                // 获取目标元素的运行时ID
                var targetRuntimeId = targetElement.GetRuntimeId();
                if (targetRuntimeId == null) return null;
                
                // 在树中递归查找
                return FindElementInTreeRecursive(_uiElements, targetRuntimeId);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"在树中查找元素时出错: {ex.Message}");
                return null;
            }
        }

        private UIElementInfo? FindElementInTreeRecursive(ObservableCollection<UIElementInfo> elements, int[] targetRuntimeId)
        {
            foreach (var element in elements)
            {
                try
                {
                    if (element.AutomationElement != null)
                    {
                        var currentRuntimeId = element.AutomationElement.GetRuntimeId();
                        if (currentRuntimeId != null && RuntimeIdsEqual(currentRuntimeId, targetRuntimeId))
                        {
                            return element;
                        }
                        
                        // 递归查找子元素
                        var found = FindElementInTreeRecursive(element.Children, targetRuntimeId);
                        if (found != null)
                        {
                            return found;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"比较运行时ID时出错: {ex.Message}");
                }
            }
            
            return null;
        }

        private bool RuntimeIdsEqual(int[] id1, int[] id2)
        {
            if (id1 == null || id2 == null) return false;
            if (id1.Length != id2.Length) return false;
            
            for (int i = 0; i < id1.Length; i++)
            {
                if (id1[i] != id2[i]) return false;
            }
            
            return true;
        }

        private void SelectTreeViewItem(UIElementInfo elementInfo)
        {
            try
            {
                // 展开并滚动到目标节点
                ExpandAndScrollToItem(elementInfo);
                
                // 通过TreeViewItem来选中项
                var treeViewItem = GetTreeViewItemFromElement(elementInfo);
                if (treeViewItem != null)
                {
                    treeViewItem.IsSelected = true;
                    treeViewItem.BringIntoView();
                    treeViewItem.Focus();
                }
                
                // 更新选中的元素
                _selectedElement = elementInfo;
                LoadElementProperties(elementInfo);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"选中树节点时出错: {ex.Message}");
                UpdateStatus($"选中树节点时出错: {ex.Message}");
            }
        }

        private void ExpandAndScrollToItem(UIElementInfo elementInfo)
        {
            try
            {
                // 获取到目标节点的路径
                var path = GetPathToElement(elementInfo);
                if (path == null || path.Count == 0) return;

                // 从根节点开始逐级展开
                ItemsControl currentContainer = treeViewControls;
                TreeViewItem? lastTreeViewItem = null;

                for (int i = 0; i < path.Count; i++)
                {
                    var currentElement = path[i];
                    var treeViewItem = FindDirectTreeViewItem(currentContainer, currentElement);
                    
                    if (treeViewItem != null)
                    {
                        // 展开当前节点（除了最后一个节点）
                        if (i < path.Count - 1)
                        {
                            treeViewItem.IsExpanded = true;
                            // 确保子项容器已生成
                            treeViewItem.UpdateLayout();
                        }
                        
                        lastTreeViewItem = treeViewItem;
                        currentContainer = treeViewItem;
                    }
                    else
                    {
                        Debug.WriteLine($"无法找到路径中的节点: {currentElement.DisplayName}");
                        break;
                    }
                }

                // 选中并滚动到最终节点
                if (lastTreeViewItem != null)
                {
                    lastTreeViewItem.IsSelected = true;
                    lastTreeViewItem.BringIntoView();
                    lastTreeViewItem.Focus();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"展开并滚动到节点时出错: {ex.Message}");
            }
        }

        private List<UIElementInfo>? GetPathToElement(UIElementInfo targetElement)
        {
            try
            {
                // 在树中查找目标元素的路径
                var path = new List<UIElementInfo>();
                if (FindElementPath(_uiElements, targetElement, path))
                {
                    return path;
                }
                return null;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取元素路径时出错: {ex.Message}");
                return null;
            }
        }

        private bool FindElementPath(ObservableCollection<UIElementInfo> elements, UIElementInfo target, List<UIElementInfo> path)
        {
            foreach (var element in elements)
            {
                path.Add(element);
                
                if (element == target)
                {
                    return true;
                }
                
                if (FindElementPath(element.Children, target, path))
                {
                    return true;
                }
                
                path.RemoveAt(path.Count - 1);
            }
            
            return false;
        }

        private TreeViewItem? FindDirectTreeViewItem(ItemsControl container, UIElementInfo target)
        {
            try
            {
                // 确保容器已生成所有子项
                container.UpdateLayout();
                
                for (int i = 0; i < container.Items.Count; i++)
                {
                    var item = container.ItemContainerGenerator.ContainerFromIndex(i) as TreeViewItem;
                    if (item?.DataContext == target)
                    {
                        return item;
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"查找直接TreeViewItem时出错: {ex.Message}");
            }
            
            return null;
        }

        private void SelectAndExpandToElement(UIElementInfo elementInfo)
        {
            try
            {
                // 获取到目标元素的路径
                var path = GetPathToElement(elementInfo);
                if (path == null || path.Count == 0)
                {
                    UpdateStatus("无法找到元素路径");
                    return;
                }

                // 逐级展开父节点
                ExpandPathInTreeView(path);

                // 通过TreeViewItem来选中目标元素
                var targetTreeViewItem = GetTreeViewItemFromElement(elementInfo);
                if (targetTreeViewItem != null)
                {
                    targetTreeViewItem.IsSelected = true;
                }
                
                // 确保选中的项可见
                var treeViewItem = GetTreeViewItemFromElement(elementInfo);
                if (treeViewItem != null)
                {
                    treeViewItem.BringIntoView();
                    treeViewItem.Focus();
                }

                // 更新选中的元素
                _selectedElement = elementInfo;
                LoadElementProperties(elementInfo);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"选中并展开到元素时出错: {ex.Message}");
                UpdateStatus($"定位元素时出错: {ex.Message}");
            }
        }



        private void ExpandPathInTreeView(List<UIElementInfo> path)
        {
            try
            {
                // 从根节点开始逐级展开
                for (int i = 0; i < path.Count - 1; i++) // 不展开最后一个节点（目标节点）
                {
                    var element = path[i];
                    var treeViewItem = GetTreeViewItemFromElement(element);
                    if (treeViewItem != null)
                    {
                        treeViewItem.IsExpanded = true;
                        treeViewItem.UpdateLayout(); // 确保子项已生成
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"展开路径时出错: {ex.Message}");
            }
        }

        private TreeViewItem? GetTreeViewItemFromElement(UIElementInfo element)
        {
            try
            {
                return FindTreeViewItemRecursive(treeViewControls, element);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取TreeViewItem时出错: {ex.Message}");
                return null;
            }
        }

        private TreeViewItem? FindTreeViewItemRecursive(ItemsControl container, UIElementInfo target)
        {
            try
            {
                if (container == null) return null;

                for (int i = 0; i < container.Items.Count; i++)
                {
                    var containerItem = container.ItemContainerGenerator.ContainerFromIndex(i);
                    if (containerItem is TreeViewItem treeViewItem)
                    {
                        if (treeViewItem.DataContext == target)
                        {
                            return treeViewItem;
                        }

                        // 如果节点已展开，递归查找子节点
                        if (treeViewItem.IsExpanded)
                        {
                            var childItem = FindTreeViewItemRecursive(treeViewItem, target);
                            if (childItem != null)
                            {
                                return childItem;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"递归查找TreeViewItem时出错: {ex.Message}");
            }

            return null;
        }

        private bool IsOwnApplicationElement(AutomationElement element)
        {
            try
            {
                if (element == null) return false;
                
                // 获取当前进程ID和进程名
                var currentProcess = Process.GetCurrentProcess();
                var currentProcessId = currentProcess.Id;
                var currentProcessName = currentProcess.ProcessName;
                
                // 获取点击元素的进程ID
                var elementProcessId = element.Current.ProcessId;
                
                // 获取点击元素的进程名（用于更准确的判断）
                string elementProcessName = "";
                try
                {
                    var elementProcess = Process.GetProcessById(elementProcessId);
                    elementProcessName = elementProcess.ProcessName;
                }
                catch
                {
                    elementProcessName = "Unknown";
                }
                
                // 获取当前窗口句柄
                var currentWindowHandle = GetHandle();
                
                // 获取点击元素的窗口句柄
                var elementWindowHandle = element.Current.NativeWindowHandle;
                
                // 添加详细的调试信息
                Debug.WriteLine("=== 点击获取句柄调试信息 ===");
                Debug.WriteLine($"当前进程ID: {currentProcessId}");
                Debug.WriteLine($"当前进程名: {currentProcessName}");
                Debug.WriteLine($"当前窗口句柄: 0x{currentWindowHandle:X}");
                Debug.WriteLine($"点击元素进程ID: {elementProcessId}");
                Debug.WriteLine($"点击元素进程名: {elementProcessName}");
                Debug.WriteLine($"点击元素窗口句柄: 0x{elementWindowHandle:X}");
                Debug.WriteLine($"点击元素名称: '{GetElementName(element)}'");
                Debug.WriteLine($"点击元素类型: {GetControlTypeName(element)}");
                Debug.WriteLine($"点击元素类名: '{element.Current.ClassName}'");
                Debug.WriteLine($"点击元素AutomationId: '{element.Current.AutomationId}'");
                
                // 检查窗口句柄是否相同
                if (elementWindowHandle == currentWindowHandle)
                {
                    Debug.WriteLine(">>> 检测到点击了自己程序的窗口，拒绝处理");
                    UpdateStatus($"检测到自己程序窗口: 0x{elementWindowHandle:X}");
                    return true;
                }
                
                // 检查进程ID和进程名是否相同
                if (elementProcessId == currentProcessId && 
                    string.Equals(elementProcessName, currentProcessName, StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine(">>> 检测到点击了自己程序的控件，拒绝处理");
                    UpdateStatus($"检测到自己程序控件: {elementProcessName} (PID: {elementProcessId})");
                    return true;
                }
                
                Debug.WriteLine(">>> 检测到外部程序控件，允许处理");
                UpdateStatus($"检测到外部程序控件: {elementProcessName} (PID: {elementProcessId})");
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"检查是否为自己程序元素时出错: {ex.Message}");
                UpdateStatus($"检查元素归属时出错: {ex.Message}");
                // 出错时假设不是自己的元素，允许继续处理
                return false;
            }
        }

        private IntPtr GetHandle()
        {
            try
            {
                // 获取当前窗口的句柄
                var source = System.Windows.PresentationSource.FromVisual(this);
                if (source is System.Windows.Interop.HwndSource hwndSource)
                {
                    return hwndSource.Handle;
                }
                return IntPtr.Zero;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取窗口句柄时出错: {ex.Message}");
                return IntPtr.Zero;
            }
        }
        
        private AutomationElement? GetRootWindow(AutomationElement element)
        {
            try
            {
                var current = element;
                while (current != null)
                {
                    var parent = TreeWalker.ControlViewWalker.GetParent(current);
                    if (parent == null || parent == AutomationElement.RootElement)
                    {
                        return current;
                    }
                    current = parent;
                }
                return current;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"获取根窗口时出错: {ex.Message}");
                return null;
            }
        }

        private void CreateAndShowTemporaryElement(AutomationElement element)
        {
            try
            {
                // 创建一个临时的UIElementInfo来显示点击的元素信息
                var tempElement = new UIElementInfo
                {
                    AutomationElement = element,
                    DisplayName = GetElementName(element),
                    ControlType = GetControlTypeName(element),
                    Children = new ObservableCollection<UIElementInfo>()
                };

                // 直接显示属性，不添加到树中
                _selectedElement = tempElement;
                LoadElementProperties(tempElement);
                
                UpdateStatus($"已获取控件信息: {tempElement.DisplayName} (未在树中找到对应节点)");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"创建临时元素时出错: {ex.Message}");
                UpdateStatus($"创建临时元素时出错: {ex.Message}");
            }
        }

        // 窗口关闭时清理资源
        protected override void OnClosed(EventArgs e)
        {
            try
            {
                // 停止点击获取功能
                if (_isPickingElement)
                {
                    StopPickingElement();
                }
                
                // 移除高亮
                RemoveHighlight();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"窗口关闭清理时出错: {ex.Message}");
            }
            
            base.OnClosed(e);
        }
    }

    // 数据模型类
    public class UIElementInfo : INotifyPropertyChanged
    {
        private string _displayName = string.Empty;
        private string _controlType = string.Empty;
        private ObservableCollection<UIElementInfo> _children = new();

        public AutomationElement AutomationElement { get; set; } = null!;

        public string DisplayName
        {
            get => _displayName;
            set
            {
                _displayName = value;
                OnPropertyChanged(nameof(DisplayName));
            }
        }

        public string ControlType
        {
            get => _controlType;
            set
            {
                _controlType = value;
                OnPropertyChanged(nameof(ControlType));
            }
        }

        public ObservableCollection<UIElementInfo> Children
        {
            get => _children;
            set
            {
                _children = value;
                OnPropertyChanged(nameof(Children));
            }
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class PropertyInfo
    {
        public string Name { get; set; } = string.Empty;
        public string Value { get; set; } = string.Empty;
    }

    public class PatternInfo
    {
        public string Name { get; set; } = string.Empty;
        public bool IsSupported { get; set; }
    }
}