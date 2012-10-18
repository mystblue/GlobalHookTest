using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace GlobalHookTest
{
    class CheckActiveWindow
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int lpdwProcessId);

        public static bool getForegroundProcessName(string name)
        {
            // アクティブなウィンドウハンドルの取得
            IntPtr hWnd = GetForegroundWindow();
            if (hWnd != null && hWnd != IntPtr.Zero)
            {
                int id;
                // ウィンドウハンドルからプロセスIDを取得
                GetWindowThreadProcessId(hWnd, out id);
                Process process = Process.GetProcessById(id);
                if (process != null)
                    return true;
                //    if (name.Equals(process.ProcessName))
                //        return true;
            }
            return false;
        }
    }
}
