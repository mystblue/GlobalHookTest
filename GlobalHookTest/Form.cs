using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace GlobalHookTest
{
    class MyForm : Form
    {
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct POINT
        {
            public int x;
            public int y;
        }
        
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public int mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, int dwThreadId);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, int dwThreadId);

        [System.Runtime.InteropServices.UnmanagedFunctionPointer(System.Runtime.InteropServices.CallingConvention.Cdecl)]
        delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, ref KBDLLHOOKSTRUCT lParam);

        [System.Runtime.InteropServices.UnmanagedFunctionPointer(System.Runtime.InteropServices.CallingConvention.Cdecl)]
        delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, ref KBDLLHOOKSTRUCT lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern IntPtr CallNextHookEx(IntPtr hHook, int nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool UnhookWindowsHookEx(IntPtr hHook);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern int GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public const int VK_CONTROL = 0x11;
        public const int WH_KEYBOARD_LL = 13;
        public const int WH_MOUSE_LL = 14;
        public const int WM_KEYDOWN = 0x0100;
        public const int WM_KEYUP = 0x0101;
        public const int WM_SYSKEYDOWN = 0x0104;
        public const int WM_SYSKEYUP = 0x0105;

        const int WM_LBUTTONDOWN = 0x201;
        const int WM_LBUTTONUP = 0x202;
        const int WM_MBUTTONDOWN = 0x207;
        const int WM_MBUTTONUP = 0x208;
        const int WM_RBUTTONDOWN = 0x204;
        const int WM_RBUTTONUP = 0x205;
        const int WM_MOUSEMOVE = 0x200;
        const int WM_MOUSEWHEEL = 0x20A;

        private static IntPtr s_hook;
        private static IntPtr mHook;

        public sealed class KeybordCaptureEventArgs : EventArgs
        {
            private int m_keyCode;
            private int m_scanCode;
            private int m_flags;
            private int m_time;
            private bool m_cancel;

            internal KeybordCaptureEventArgs(KBDLLHOOKSTRUCT keyData)
            {
                this.m_keyCode = keyData.vkCode;
                this.m_scanCode = keyData.scanCode;
                this.m_flags = keyData.flags;
                this.m_time = keyData.time;
                this.m_cancel = false;
            }

            public int KeyCode { get { return this.m_keyCode; } }
            public int ScanCode { get { return this.m_scanCode; } }
            public int Flags { get { return this.m_flags; } }
            public int Time { get { return this.m_time; } }
            public bool Cancel
            {
                set { this.m_cancel = value; }
                get { return this.m_cancel; }
            }
        }

        public sealed class MouseCaptureEventArgs : System.EventArgs
        {
            private bool m_cancel;
            private int m_nativeWParam;
            private MSLLHOOKSTRUCT m_nativeLParam;
            private bool m_isValueUpdate;

            private MouseButtons m_button;
            private int m_x;
            private int m_y;
            private int m_delta;
            private int m_time;

            internal MouseCaptureEventArgs(int wParam, MSLLHOOKSTRUCT lParam)
            {
                this.m_cancel = false;
                this.m_isValueUpdate = false;
                this.m_nativeWParam = wParam;
                this.m_nativeLParam = lParam;

                this.m_button = MouseButtons.None;

                switch (wParam)
                {
                    case WM_LBUTTONDOWN:
                    case WM_LBUTTONUP:
                        this.m_button = MouseButtons.Left;
                        break;
                    case WM_MBUTTONDOWN:
                    case WM_MBUTTONUP:
                        this.m_button = MouseButtons.Middle;
                        break;
                    case WM_RBUTTONDOWN:
                    case WM_RBUTTONUP:
                        this.m_button = MouseButtons.Right;
                        break;

                    default:
                        break;
                }
                this.m_x = lParam.pt.x;
                this.m_y = lParam.pt.y;
                this.m_delta = (wParam == WM_MOUSEWHEEL)
                    ? (int)((short)(lParam.mouseData >> 16)) : 0;
                this.m_time = lParam.time;
            }

            public MouseButtons Button { get { return this.m_button; } }
            public int X { get { return this.m_x; } }
            public int Y { get { return this.m_y; } }
            public int Delta { get { return this.m_delta; } }
            public int Time { get { return this.m_time; } }

            public bool Cancel
            {
                set { this.m_cancel = value; }
                get { return this.m_cancel; }
            }
            public int NativeWParam
            {
                set
                {
                    this.m_nativeWParam = value;
                    this.m_isValueUpdate = true;
                }
                get { return this.m_nativeWParam; }
            }
            public MSLLHOOKSTRUCT NativeLParam
            {
                set
                {
                    this.m_nativeLParam = value;
                    this.m_isValueUpdate = true;
                }
                get { return this.m_nativeLParam; }
            }
            internal bool IsValueUpdate { get { return this.m_isValueUpdate; } }
        }

        public MyForm()
        {
            this.Text = "Microsoft Office Test";

            // 位置
            this.StartPosition = FormStartPosition.Manual;
            this.Location = new System.Drawing.Point(100, 20);

            // サイズ
            this.Size = new System.Drawing.Size(1024, 700);

            WebBrowser webBrowser = new WebBrowser();
            webBrowser.Size = new System.Drawing.Size(1024 - 16, 700 - 14 - 26);
            this.Controls.Add(webBrowser);

            // アンカーの設定
            webBrowser.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Bottom | AnchorStyles.Right;

            // レジストリの変更
            Microsoft.Win32.RegistryKey regkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\Shell\AttachmentExecute\{0002DF01-0000-0000-C000-000000000046}");
            // 8 12 6 も同様に行なっておく
            regkey.SetValue("Excel.Sheet.8", new byte[] { }, Microsoft.Win32.RegistryValueKind.Binary);

            //webBrowser.Navigate("http://inos-whitelist.appspot.com/Book1.xls");
            webBrowser.Navigate(@"C:\Users\314.FINOS\Documents\docs\Book1.xls");
            //webBrowser.Navigate("http://www.google.co.jp");


            //this.FormClosing += new FormClosingEventHandler(MyForm_FormClosing);
            this.Activated += new EventHandler(MyForm_Activated);
            this.Deactivate += new EventHandler(MyForm_Deactivate);
        }

        void MyForm_Deactivate(object sender, EventArgs e)
        {
            UnhookWindowsHookEx(s_hook);
            UnhookWindowsHookEx(mHook);
        }

        void MyForm_Activated(object sender, EventArgs e)
        {
            s_hook = SetWindowsHookEx(WH_KEYBOARD_LL, HookProc,
                Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0]), 0);

            mHook = SetWindowsHookEx(WH_MOUSE_LL, MouseHookProc,
                Marshal.GetHINSTANCE(System.Reflection.Assembly.GetExecutingAssembly().GetModules()[0]), 0);
        }

        void MyForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnhookWindowsHookEx(s_hook);
        }

        static IntPtr HookProc(int nCode, IntPtr wParam, ref KBDLLHOOKSTRUCT lParam)
        {
            bool cancel = false;
            const int HC_ACTION = 0;
            if (nCode == HC_ACTION)
            {
                KeybordCaptureEventArgs ev = new KeybordCaptureEventArgs(lParam);
                Keys keyCode = (Keys)lParam.vkCode;
                switch (wParam.ToInt32())
                {
                    case WM_KEYDOWN:
                        //if (GetAsyncKeyState(VK_CONTROL) != 0)
                        //    cancel = true;
                        if (keyCode == Keys.P && GetAsyncKeyState(VK_CONTROL) != 0)
                            cancel = true;
                        break;

                    case WM_KEYUP:
                        //if (GetAsyncKeyState(VK_CONTROL) != 0)
                        //    cancel = true;
                        if (keyCode == Keys.P && GetAsyncKeyState(VK_CONTROL) != 0)
                            cancel = true;
                        break;

                    //case WM_SYSKEYDOWN:
                    //    CallEvent(SysKeyDown, ev);
                    //    break;

                    //case WM_SYSKEYUP:
                    //    CallEvent(SysKeyUp, ev);
                    //    break;
                }
                //cancel = ev.Cancel;
            }
            return cancel ? (IntPtr)1 : CallNextHookEx(s_hook, nCode, wParam, ref lParam);
        }

        static IntPtr MouseHookProc(int nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam)
        {
            bool cancel = false;
            const int HC_ACTION = 0;
            if (nCode == HC_ACTION)
            {
                MouseCaptureEventArgs ev = new MouseCaptureEventArgs((int)wParam, lParam);
                switch (wParam.ToInt32())
                {
                    //case WM_LBUTTONDOWN:
                    //case WM_MBUTTONDOWN:
                    case WM_RBUTTONDOWN:
                        cancel = true;
                        break;

                    //case WM_LBUTTONUP:
                    //case WM_MBUTTONUP:
                    case WM_RBUTTONUP:
                        cancel = true;
                        break;

                    //case WM_MOUSEMOVE:
                    //    CallEvent(MouseMove, ev);
                    //    break;
                    //case WM_MOUSEWHEEL:
                    //    CallEvent(MouseWheel, ev);
                    //    break;
                }
                //if (ev.IsValueUpdate)
                //{
                //    wParam = (IntPtr)ev.NativeWParam;
                //    lParam = ev.NativeLParam;
                //}
                //cancel = ev.Cancel;
            }
            return cancel ? (IntPtr)1 : CallNextHookEx(s_hook, nCode, wParam, ref lParam);
        }

    }
}
