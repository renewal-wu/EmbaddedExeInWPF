using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Interop;

namespace HostDemo
{
    public partial class HostControl : UserControl, IDisposable
    {
        [System.Runtime.InteropServices.StructLayoutAttribute(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct HWND__
        {
            /// int
            public int unused;
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        private static extern long GetWindowThreadProcessId(long hWnd, long lpdwProcessId);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern long SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongA", SetLastError = true)]
        private static extern long GetWindowLong(IntPtr hwnd, int nIndex);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongA", SetLastError = true)]
        public static extern int SetWindowLongA([System.Runtime.InteropServices.InAttribute()] System.IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern long SetWindowPos(IntPtr hwnd, long hWndInsertAfter, long x, long y, long cx, long cy, long wFlags);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool MoveWindow(IntPtr hwnd, int x, int y, int cx, int cy, bool repaint);

        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hwnd, WindowEnumProc func, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("Shcore.dll")]
        private static extern int SetProcessDpiAwareness(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private const int SWP_NOOWNERZORDER = 0x200;
        private const int SWP_NOREDRAW = 0x8;
        private const int SWP_NOZORDER = 0x4;
        private const int SWP_SHOWWINDOW = 0x0040;
        private const int WS_EX_MDICHILD = 0x40;
        private const int SWP_FRAMECHANGED = 0x20;
        private const int SWP_NOACTIVATE = 0x10;
        private const int SWP_ASYNCWINDOWPOS = 0x4000;
        private const int SWP_NOMOVE = 0x2;
        private const int SWP_NOSIZE = 0x1;
        private const int GWL_STYLE = (-16);
        private const int WS_VISIBLE = 0x10000000;
        private const int WS_CHILD = 0x40000000;
        private const int WM_ACTIVATE = 0x0006;
        private readonly IntPtr WA_ACTIVE = new IntPtr(1);
        private readonly IntPtr WA_INACTIVE = new IntPtr(0);

        public event EventHandler HostProcessStarted;

        private void OnHostProcessStarted()
        {
            HostProcessStarted?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Track if the application has been created
        /// </summary>
        public bool IsCreated { get; private set; } = false;

        public string ExeName { get; set; }

        /// <summary>
        /// Track if the control is disposed
        /// </summary>
        private bool isDisposed = false;

        /// <summary>
        /// Handle to the application Window
        /// </summary>
        private IntPtr appWin;

        private Process childp;

        public HostControl()
        {
            InitializeComponent();
            this.Loaded += HostControl_Loaded;
            this.Unloaded += HostControl_Unloaded;
            this.SizeChanged += HostControl_SizeChanged;
        }

        ~HostControl()
        {
            this.Dispose();
        }

        private void HostControl_Loaded(object sender, RoutedEventArgs e)
        {
            if (IsCreated)
            {
                return;
            }

            if (string.IsNullOrEmpty(ExeName))
            {
                return;
            }

            Application.Current.Exit -= Current_Exit;
            Application.Current.Exit += Current_Exit;

            appWin = IntPtr.Zero;

            try
            {
                var procInfo = new ProcessStartInfo(this.ExeName);
                procInfo.WorkingDirectory = System.IO.Path.GetDirectoryName(this.ExeName);

                childp = Process.Start(procInfo);

                IsCreated = true;

                childp.WaitForInputIdle();

                this.appWin = childp.MainWindowHandle;
                var helper = new WindowInteropHelper(Window.GetWindow(this));

                SetParent(appWin, helper.Handle);
                SetWindowLongA(appWin, GWL_STYLE, WS_VISIBLE);

                if (childp != null && childp.HasExited == false)
                {
                    OnHostProcessStarted();
                }

                UpdateSize();
            }
            catch (Exception ex)
            {
                Debug.Print(ex.Message + "Error");

                // 出錯了，把自己隱藏起來
                this.Visibility = Visibility.Collapsed;
            }
        }

        private void HostControl_Unloaded(object sender, RoutedEventArgs e)
        {
            Application.Current.Exit -= Current_Exit;
            this.Dispose();
        }

        private void Current_Exit(object sender, ExitEventArgs e)
        {
            this.Dispose();
        }

        private void HostControl_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            UpdateSize();
        }

        private void UpdateSize()
        {
            if (this.appWin != IntPtr.Zero)
            {
                PresentationSource source = PresentationSource.FromVisual(this);

                var scaleX = 1D;
                var scaleY = 1D;
                if (source != null)
                {
                    scaleX = source.CompositionTarget.TransformToDevice.M11;
                    scaleY = source.CompositionTarget.TransformToDevice.M22;
                }

                var width = (int)(this.ActualWidth * scaleX);
                var height = (int)(this.ActualHeight * scaleY);

                MoveWindow(appWin, 0, 0, width, height, true);
            }
        }

        protected void Dispose(bool disposing)
        {
            if (!isDisposed)
            {
                if (disposing)
                {
                    if (IsCreated && childp != null && !childp.HasExited)
                    {
                        childp.Kill();
                    }

                    if (appWin != IntPtr.Zero)
                    {
                        appWin = IntPtr.Zero;
                    }
                }

                isDisposed = true;
            }
        }

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }
    }
}