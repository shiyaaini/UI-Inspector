using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace window_tools
{
    /// <summary>
    /// 高亮窗口，用于显示选中元素的边界
    /// </summary>
    public class HighlightWindow : Window
    {
        private Rectangle _highlightRect;
        private DispatcherTimer _blinkTimer;
        private bool _isVisible = true;

        // Win32 API 声明
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        static extern bool ClientToScreen(IntPtr hWnd, ref POINT lpPoint);

        [DllImport("user32.dll")]
        static extern IntPtr GetDC(IntPtr hWnd);

        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);

        [DllImport("user32.dll")]
        static extern bool ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("user32.dll")]
        static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

        [DllImport("shcore.dll")]
        static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

        // DPI相关常量
        private const int LOGPIXELSX = 88;
        private const int LOGPIXELSY = 90;
        private const uint MONITOR_DEFAULTTONEAREST = 2;
        private const int MDT_EFFECTIVE_DPI = 0;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;     // X 坐标（左上角）
            public int Top;      // Y 坐标（左上角）
            public int Right;    // X 坐标（右下角）
            public int Bottom;   // Y 坐标（右下角）
            
            // 转换为易用的宽高属性
            public int Width => Right - Left;
            public int Height => Bottom - Top;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public HighlightWindow(System.Drawing.Rectangle bounds)
        {
            // 设置窗口属性
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            Background = Brushes.Transparent;
            Topmost = true;
            ShowInTaskbar = false;
            ResizeMode = ResizeMode.NoResize;
            SizeToContent = SizeToContent.Manual;

            // 获取系统DPI缩放因子
            var dpiScale = GetSystemDpiScale();
            
            // 将物理像素坐标转换为WPF逻辑坐标
            Left = bounds.X / dpiScale.X;
            Top = bounds.Y / dpiScale.Y;
            Width = bounds.Width / dpiScale.X;
            Height = bounds.Height / dpiScale.Y;

            // 创建高亮矩形
            _highlightRect = new Rectangle
            {
                Stroke = Brushes.Red,
                StrokeThickness = 2,
                Fill = Brushes.Transparent,
                Width = Width,
                Height = Height
            };

            Content = _highlightRect;

            // 创建闪烁定时器
            _blinkTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _blinkTimer.Tick += BlinkTimer_Tick;

            // 窗口加载完成后显示
            Loaded += (s, e) =>
            {
                try
                {
                    _blinkTimer.Start();
                    Activate();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"高亮窗口加载时出错: {ex.Message}");
                }
            };

            // 窗口关闭时清理资源
            Closed += (s, e) =>
            {
                try
                {
                    _blinkTimer?.Stop();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"高亮窗口关闭时出错: {ex.Message}");
                }
            };
        }

        /// <summary>
        /// 通过窗口句柄创建高亮窗口
        /// </summary>
        public static HighlightWindow? CreateFromHandle(IntPtr hwnd)
        {
            try
            {
                if (GetWindowRect(hwnd, out RECT rect))
                {
                    return new HighlightWindow(new System.Drawing.Rectangle(
                        rect.Left, rect.Top, rect.Width, rect.Height));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"通过句柄创建高亮窗口失败: {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// 通过客户端区域创建高亮窗口
        /// </summary>
        public static HighlightWindow? CreateFromClientRect(IntPtr hwnd)
        {
            try
            {
                if (GetClientRect(hwnd, out RECT clientRect))
                {
                    // 将客户端坐标转换为屏幕坐标
                    POINT topLeft = new POINT { X = 0, Y = 0 };
                    POINT bottomRight = new POINT { X = clientRect.Right, Y = clientRect.Bottom };
                    
                    if (ClientToScreen(hwnd, ref topLeft) && ClientToScreen(hwnd, ref bottomRight))
                    {
                        return new HighlightWindow(new System.Drawing.Rectangle(
                            topLeft.X, topLeft.Y, 
                            bottomRight.X - topLeft.X, 
                            bottomRight.Y - topLeft.Y));
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"通过客户端区域创建高亮窗口失败: {ex.Message}");
            }
            return null;
        }

        /// <summary>
        /// 获取系统DPI缩放因子
        /// </summary>
        private DpiScale GetDpiScale()
        {
            try
            {
                // 获取主显示器的DPI
                IntPtr hdc = GetDC(IntPtr.Zero);
                if (hdc != IntPtr.Zero)
                {
                    int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                    ReleaseDC(IntPtr.Zero, hdc);

                    // 标准DPI是96
                    double scaleX = dpiX / 96.0;
                    double scaleY = dpiY / 96.0;

                    return new DpiScale(scaleX, scaleY);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取DPI缩放因子失败: {ex.Message}");
            }

            // 默认返回无缩放
            return new DpiScale(1.0, 1.0);
        }

        /// <summary>
        /// 获取系统DPI缩放因子（简化版本）
        /// </summary>
        private System.Drawing.PointF GetSystemDpiScale()
        {
            try
            {
                // 获取主显示器的DPI
                IntPtr hdc = GetDC(IntPtr.Zero);
                if (hdc != IntPtr.Zero)
                {
                    int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                    ReleaseDC(IntPtr.Zero, hdc);

                    // 标准DPI是96
                    float scaleX = dpiX / 96.0f;
                    float scaleY = dpiY / 96.0f;

                    return new System.Drawing.PointF(scaleX, scaleY);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取系统DPI缩放因子失败: {ex.Message}");
            }

            // 默认返回无缩放
            return new System.Drawing.PointF(1.0f, 1.0f);
        }

        /// <summary>
        /// 创建精确的高亮窗口，使用UI Automation坐标
        /// </summary>
        public static HighlightWindow CreateFromAutomationRect(System.Windows.Rect automationRect)
        {
            // UI Automation在不同DPI设置下的行为：
            // - 在96 DPI (100%缩放)下，返回物理像素坐标
            // - 在高DPI下，返回的坐标需要根据系统DPI感知设置来处理
            
            // 获取当前进程的DPI感知模式
            var dpiAwareness = GetProcessDpiAwareness();
            
            System.Drawing.Rectangle physicalRect;
            
            if (dpiAwareness == ProcessDpiAwareness.DpiUnaware)
            {
                // DPI不感知模式：UI Automation返回的是缩放后的坐标，需要转换为物理坐标
                var systemDpi = GetStaticSystemDpiScale();
                physicalRect = new System.Drawing.Rectangle(
                    (int)Math.Round(automationRect.X * systemDpi.X), 
                    (int)Math.Round(automationRect.Y * systemDpi.Y), 
                    (int)Math.Round(automationRect.Width * systemDpi.X), 
                    (int)Math.Round(automationRect.Height * systemDpi.Y));
            }
            else
            {
                // DPI感知模式：UI Automation返回的已经是物理像素坐标
                physicalRect = new System.Drawing.Rectangle(
                    (int)Math.Round(automationRect.X), 
                    (int)Math.Round(automationRect.Y), 
                    (int)Math.Round(automationRect.Width), 
                    (int)Math.Round(automationRect.Height));
            }

            return new HighlightWindow(physicalRect);
        }

        /// <summary>
        /// DPI感知模式枚举
        /// </summary>
        private enum ProcessDpiAwareness
        {
            DpiUnaware = 0,
            SystemDpiAware = 1,
            PerMonitorDpiAware = 2
        }

        /// <summary>
        /// 获取当前进程的DPI感知模式
        /// </summary>
        private static ProcessDpiAwareness GetProcessDpiAwareness()
        {
            try
            {
                // 简化处理：假设应用程序是DPI感知的
                // 在实际应用中，可以通过GetProcessDpiAwareness API获取
                return ProcessDpiAwareness.SystemDpiAware;
            }
            catch
            {
                return ProcessDpiAwareness.DpiUnaware;
            }
        }

        /// <summary>
        /// 获取系统DPI缩放因子（静态版本）
        /// </summary>
        private static System.Drawing.PointF GetStaticSystemDpiScale()
        {
            try
            {
                // 获取主显示器的DPI
                IntPtr hdc = GetDC(IntPtr.Zero);
                if (hdc != IntPtr.Zero)
                {
                    int dpiX = GetDeviceCaps(hdc, LOGPIXELSX);
                    int dpiY = GetDeviceCaps(hdc, LOGPIXELSY);
                    ReleaseDC(IntPtr.Zero, hdc);

                    // 标准DPI是96
                    float scaleX = dpiX / 96.0f;
                    float scaleY = dpiY / 96.0f;

                    return new System.Drawing.PointF(scaleX, scaleY);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"获取系统DPI缩放因子失败: {ex.Message}");
            }

            // 默认返回无缩放
            return new System.Drawing.PointF(1.0f, 1.0f);
        }

        private void BlinkTimer_Tick(object? sender, EventArgs e)
        {
            _isVisible = !_isVisible;
            _highlightRect.Opacity = _isVisible ? 1.0 : 0.3;
        }

        protected override void OnClosed(EventArgs e)
        {
            _blinkTimer?.Stop();
            base.OnClosed(e);
        }
    }
} 