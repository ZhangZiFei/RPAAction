using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RPAAction.Access_CSO
{
    public abstract class AccessAction : Base.RPAAction
    {
        public static _Application ChangeAppForUser(_Application app)
        {
            app.UserControl = true;
            //app.Visible = true;
            return app;
        }

        public static _Application ChangeAppForRPA(_Application app)
        {
            app.UserControl = false;
            return app;
        }

        public static _Application GetApplication(string accessPath)
        {
            if (CheckString(accessPath))
            {
                accessPath = System.IO.Path.GetFullPath(accessPath);
                if (!apps.ContainsKey(accessPath))
                {
                    _Application app = new Application
                    {
                        Visible = true
                    };
                    app.OpenCurrentDatabase(accessPath);
                    apps.Add(accessPath, app);
                }
                return ChangeAppForRPA(apps[accessPath]);
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// 清理缓存的Acess进程
        /// </summary>
        /// <param name="accessPath"></param>
        public static void ClearUp(string accessPath = null)
        {
            if (CheckString(accessPath))
            {
                KillAccess(apps[accessPath]);
                apps.Remove(accessPath);
            }
            else
            {
                foreach (var i in apps)
                {
                    KillAccess(i.Value);
                }
                apps.Clear();
            }
        }

        public AccessAction(string accessPath)
        {
            this.accessPath = accessPath;
        }

        protected _Application app = null;

        protected Database db = null;

        protected readonly string accessPath;

        protected override void Action()
        {
            app = GetApplication(accessPath);
            db = app.CurrentDb();
        }

        protected override void AfterRun()
        {
            base.AfterRun();

            if (app != null)
            {
                ChangeAppForUser(app);
            }
        }

        private static void KillAccess(_Application app)
        {
            GetWindowThreadProcessId(new IntPtr(app.hWndAccessApp()), out uint pid);
            app.Quit();
            Process.GetProcessById((int)pid).Kill();
        }

        private static readonly Dictionary<string, _Application> apps = new Dictionary<string, _Application>();

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    }
}