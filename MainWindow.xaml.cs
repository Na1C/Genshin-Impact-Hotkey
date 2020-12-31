using CSCore.CoreAudioAPI;
using Ownskit.Utils;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;

namespace Genshin_WPF
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        [DllImport("user32.dll")]
        public static extern bool RegisterHotKey(IntPtr hWnd, int id, int fsModifiers, int vlc);
        [DllImport("user32.dll")]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll")]
        static extern int GetSystemMetrics(SystemMetric smIndex);
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);



        [StructLayout(LayoutKind.Sequential)]
        struct INPUT
        {
            public SendInputEventType type;
            public MouseKeybdhardwareInputUnion mkhi;
        }
        [StructLayout(LayoutKind.Explicit)]
        struct MouseKeybdhardwareInputUnion
        {
            [FieldOffset(0)]
            public MouseInputData mi;

            [FieldOffset(0)]
            public KEYBDINPUT ki;

            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }
        [StructLayout(LayoutKind.Sequential)]
        struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
        [StructLayout(LayoutKind.Sequential)]
        struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }
        struct MouseInputData
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public MouseEventFlags dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
        [Flags]
        enum MouseEventFlags : uint
        {
            MOUSEEVENTF_MOVE = 0x0001,
            MOUSEEVENTF_LEFTDOWN = 0x0002,
            MOUSEEVENTF_LEFTUP = 0x0004,
            MOUSEEVENTF_RIGHTDOWN = 0x0008,
            MOUSEEVENTF_RIGHTUP = 0x0010,
            MOUSEEVENTF_MIDDLEDOWN = 0x0020,
            MOUSEEVENTF_MIDDLEUP = 0x0040,
            MOUSEEVENTF_XDOWN = 0x0080,
            MOUSEEVENTF_XUP = 0x0100,
            MOUSEEVENTF_WHEEL = 0x0800,
            MOUSEEVENTF_VIRTUALDESK = 0x4000,
            MOUSEEVENTF_ABSOLUTE = 0x8000
        }
        enum SendInputEventType : int
        {
            InputMouse,
            InputKeyboard,
            InputHardware
        }

        int CalculateAbsoluteCoordinateX(int x)
        {
            return (x * 65536) / GetSystemMetrics(SystemMetric.SM_CXSCREEN);
        }

        int CalculateAbsoluteCoordinateY(int y)
        {
            return (y * 65536) / GetSystemMetrics(SystemMetric.SM_CYSCREEN);
        }

        public static void SendKeyBoradKey(ushort key)
        {
            INPUT input_down = new INPUT();
            input_down.type = SendInputEventType.InputKeyboard;
            input_down.mkhi.ki.dwFlags = 0;
            input_down.mkhi.ki.wVk = key;
            SendInput(1, ref input_down, Marshal.SizeOf(input_down));//keydown     

            //INPUT input_up = new INPUT();
            //input_up.type = SendInputEventType.InputKeyboard;
            //input_up.mkhi.ki.wVk = key;
            //input_up.mkhi.ki.dwFlags = 1;
            //SendInput(1, ref input_up, Marshal.SizeOf(input_up));//keyup      

        }
        public void ClickLeftMouseButton(int x, int y)
        {
            INPUT mouseInput = new INPUT();
            mouseInput.type = SendInputEventType.InputMouse;
            mouseInput.mkhi.mi.dx = CalculateAbsoluteCoordinateX(x);
            mouseInput.mkhi.mi.dy = CalculateAbsoluteCoordinateY(y);
            mouseInput.mkhi.mi.mouseData = 0;


            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_MOVE | MouseEventFlags.MOUSEEVENTF_ABSOLUTE;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTDOWN;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));

            mouseInput.mkhi.mi.dwFlags = MouseEventFlags.MOUSEEVENTF_LEFTUP;
            SendInput(1, ref mouseInput, Marshal.SizeOf(new INPUT()));
        }

        enum SystemMetric
        {
            SM_CXSCREEN = 0,
            SM_CYSCREEN = 1,
        }
        private const int HOTKEY_ID = 9000;

        //Modifiers:
        private const int MOD_NONE = 0x0000; //[NONE]
        private const int MOD_ALT = 0x0001; //ALT
        private const int MOD_CONTROL = 0x0002; //CTRL
        private const int MOD_SHIFT = 0x0004; //SHIFT
        private const int MOD_WIN = 0x0008; //WINDOWS
        private const int VK_CAPITAL = 0x14;
        private const int VK_F1 = 0x70;
        private const int VK_F2 = 0x71;
        private const int VK_N1 = 0x61;
        private const int VK_N2 = 0x62;
        private const int VK_N3 = 0x63;
        private const int VK_A = 0x41;
        private const int VK_ESC = 0x1B;
        private const int VK_HOME = 0x24;
        private const int VK_END = 0x23;
        KeyboardListener KListener = new KeyboardListener();

        private IntPtr _windowHandle;
        private HwndSource _source;
        public MainWindow()
        {
            InitializeComponent();
            //KListener.KeyDown += KListener_KeyDown;

        }
        private void MainW_Loaded(object sender, RoutedEventArgs e)
        {
            _windowHandle = new WindowInteropHelper(this).Handle;
            _source = HwndSource.FromHwnd(new WindowInteropHelper(this).Handle);
            _source.AddHook(HwndHook);

            RegisterHotKey(_windowHandle, 0, MOD_NONE, VK_HOME);
            RegisterHotKey(_windowHandle, 1, MOD_NONE, VK_END);
        }

        private IntPtr HwndHook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == (int)0x312) //핫키가 눌러지면 312 정수 메세지
            {
                if (wParam == (IntPtr)0x0) // 그 키의 ID가 0이면
                {
                    Debug.WriteLine("Home press");
                    Task.Run(() => ChangeVolume(0.80f, 0.00f));
                }
                if (wParam == (IntPtr)0x1) // 그 키의 ID가 1이면
                {
                    Debug.WriteLine("END press");
                    SendKeyBoradKey(VK_ESC);
                    Thread.Sleep(1000);
                    ClickLeftMouseButton(350, 410);

                }
            }

            return IntPtr.Zero;
        }

        protected override void OnClosed(EventArgs e)
        {
            _source.RemoveHook(HwndHook);
            UnregisterHotKey(_windowHandle, HOTKEY_ID);
            base.OnClosed(e);
        }


        private void button_Click(object sender, RoutedEventArgs e)
        {
            Task.Run(() => ChangeVolume());
        }

        private void ChangeVolume(float fOn = 0.10f, float fMute = 0.03f)
        {
            using (var sessionManager = GetDefaultAudioSessionManager2(DataFlow.Render))
            {
                using (var sessionEnumerator = sessionManager.GetSessionEnumerator())
                {
                    foreach (var session in sessionEnumerator)
                    {
                        using (var simpleVolume = session.QueryInterface<SimpleAudioVolume>())
                        using (var sessionControl = session.QueryInterface<AudioSessionControl2>())
                        {
                            if (sessionControl.Process.ProcessName == "GenshinImpact")
                            {
                                if (simpleVolume.MasterVolume >= 0.05f)
                                {
                                    simpleVolume.MasterVolume = fMute;
                                    this.Dispatcher.Invoke(() =>
                                    {
                                        lblMute.Content = "음소거 (ON상태) Home";
                                        MainW.Background = Brushes.Green;
                                    });
                                    //lblMute.Content = "음소거 (ON상태) F1";
                                }
                                else
                                {
                                    simpleVolume.MasterVolume = fOn;
                                    this.Dispatcher.Invoke(() =>
                                    {
                                        lblMute.Content = "음소거 (OFF상태) Home";
                                        MainW.Background = Brushes.Tomato;
                                    });
                                    //lblParty.Content = "음소거 (OFF상태) F1";
                                }

                            }
                        }
                    }
                }
            }
        }

        private static AudioSessionManager2 GetDefaultAudioSessionManager2(DataFlow dataFlow)
        {
            using (var enumerator = new MMDeviceEnumerator())
            {
                using (var device = enumerator.GetDefaultAudioEndpoint(dataFlow, Role.Multimedia))
                {
                    Debug.WriteLine("DefaultDevice: " + device.FriendlyName);
                    var sessionManager = AudioSessionManager2.FromMMDevice(device);
                    return sessionManager;
                }
            }
        }


    }
}
