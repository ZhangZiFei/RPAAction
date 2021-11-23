using Microsoft.Office.Interop.Excel;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RPAAction.Excel_CSO
{
    public static class ApplicationExpand
    {
        /// <summary>
        /// 杀死Excel进程
        /// </summary>
        /// <param name="app"></param>
        public static void Kill(this _Application app)
        {
            if (app == null)
            {
                return;
            }
            else
            {
                GetWindowThreadProcessId(new IntPtr(app.Hwnd), out uint processId);
                app.Quit();
                Process p = Process.GetProcessById((int)processId);
                p.ExitOrKill();
            }
        }

        /// <summary>
        /// 检测<see cref="_Application"/>实例是否可用,如果不可用则清理
        /// </summary>
        /// <returns>可用返回true,不可用返回false</returns>
        public static bool Check(this _Application app)
        {
            if (app != null)
            {
                try
                {
                    app.Visible = app.Visible;
                    return true;
                }
                catch (COMException)
                {
                    try { app.Kill(); } catch (Exception) { }
                }
            }
            return false;
        }

        /// <summary>
        /// 显示Excel应用程序
        /// </summary>
        public static void Show(this _Application app)
        {
            app.Visible = true;
        }

        /// <summary>
        /// 隐藏Excel应用程序
        /// </summary>
        public static void Hide(this _Application app)
        {
            app.Visible = false;
        }

        /// <summary>
        /// 处理<see cref="_Application"/>适应自动化
        /// </summary>
        public static _Application ChangeForRPA(this _Application app)
        {
            //禁止Excel进程的各种弹窗
            app.DisplayAlerts = false;
            //取消用户控制模式
            app.UserControl = false;
            //显示Excel窗口
            app.Show();
            return app;
        }

        /// <summary>
        /// 适应用户
        /// </summary>
        public static _Application ChangeForUser(this _Application app)
        {
            //启用Excel进程的各种弹窗
            app.DisplayAlerts = true;
            //开启用户控制模式
            app.UserControl = true;
            //显示Excel窗口
            app.Show();
            return app;
        }

        #region user32.dll oleacc.dll
        [DllImport("user32.dll")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
             IntPtr hwnd,
             uint id,
             ref Guid iid,
             [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject
        );
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
        #endregion
    }
}
