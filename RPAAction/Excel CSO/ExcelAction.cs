using Microsoft.Office.Interop.Excel;
using RPAAction.Base;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace RPAAction.Excel_CSO
{
    public abstract class ExcelAction : Base.Action
    {
        /// <summary>
        /// 杀死Excel进程
        /// </summary>
        /// <param name="app"></param>
        public static void KillApp(_Application app)
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
                if (p.WaitForExit(100))
                {
                    return;
                }
                p.Kill();
                p.WaitForExit(10000);
            }
        }

        /// <summary>
        /// 检测<see cref="_Application"/>实例是否可用,如果不可用则清理
        /// </summary>
        /// <returns>可用返回true,不可用返回false</returns>
        public static bool CheckApp(_Application app)
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
                    try { KillApp(app); } catch (Exception) { }
                }
            }
            return false;
        }

        /// <summary>
        /// 设置Excel进程的显示状态
        /// </summary>
        public static void ShowApp(_Application app, bool b = true)
        {
            app.Visible = b;
        }

        /// <summary>
        /// 处理<see cref="_Application"/>以适应自动化操作
        /// </summary>
        public static _Application ChangeAppForRPA(_Application app)
        {
            //禁止Excel进程的各种弹窗
            app.DisplayAlerts = false;
            //取消用户控制模式
            app.UserControl = false;
            return app;
        }

        /// <summary>
        /// 处理<see cref="_Application"/>以适应用户操作
        /// </summary>
        public static _Application ChangeAppForUser(_Application app)
        {
            //禁止Excel进程的各种弹窗
            app.DisplayAlerts = true;
            //取消用户控制模式
            app.UserControl = true;
            ShowApp(app);
            return app;
        }

        /// <summary>
        /// 为RPA程式创建新的Excel进程,会改变<see cref="app"/>的指向
        /// </summary>
        /// <returns></returns>
        public static _Application CreateAppForRPA()
        {
            app = new Application();
            ChangeAppForRPA(app);
            ShowApp(app);
            return app;
        }

        /// <summary>
        /// 连接并且返回可用的<see cref="_Application"/>,如果连接失败返回null
        /// </summary>
        public static _Application AttachApp()
        {
            if (CheckApp(app))
            {
                return app;
            }
            else
            {
                do
                {
                    try
                    {
                        //连接Excel进程
                        app = (_Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        app = null;
                        break;
                    }
                } while (!CheckApp(app));
            }
            return app == null ? null : ChangeAppForRPA(app);
        }

        /// <summary>
        /// 链接或者打开<see cref="_Application"/>
        /// </summary>
        /// <returns></returns>
        public static _Application AttachOrOpenApp()
        {
            AttachApp();
            return app ?? CreateAppForRPA();
        }

        /// <summary>
        /// 处理<see cref="_Workbook"/>以适应自动化操作
        /// </summary>
        public static void ChangeWorkbookForRPA(_Workbook wb)
        {
            wb.CheckCompatibility = false;//控制兼容性检查器运行自动保存工作簿时。 为可读/写属性。
            wb.UpdateLinks = XlUpdateLinks.xlUpdateLinksNever;//禁止更新链接
        }

        /// <summary>
        /// 连接工作簿,如果失败则返回null,可能会改变<see cref="app"/>的指向
        /// </summary>
        public static _Workbook AttachWorkbook(string wbPath)
        {
            _Workbook wb = null;
            wbPath = Path.GetFullPath(wbPath);

            //检测路径和Excel进程
            if (File.Exists(wbPath) && AttachApp() != null)
            {
                wb = IAttachWorkbook1(wbPath);//方案一
                if (wb == null)
                {
                    wb = IAttachWorkbook2(wbPath);//方案二
                }
            }

            if (wb != null)
            {
                ChangeWorkbookForRPA(wb);
            }

            return wb;
        }

        /// <summary>
        /// 打开工作簿
        /// </summary>
        public static _Workbook OpenWorkbook(string wbPath, bool readOnly = false, string pwd = null, string delimiter = null, string writePwd = null)
        {
            AttachOrOpenApp();
            wbPath = Path.GetFullPath(wbPath);
            _Workbook wb = app.Workbooks.Open(
                wbPath,
                XlUpdateLinks.xlUpdateLinksNever,
                readOnly,
                CheckString(delimiter) ? delimiter : Type.Missing,
                CheckString(pwd) ? pwd : Type.Missing,
                CheckString(writePwd) ? writePwd : Type.Missing,
                true,//则不让 Microsoft Excel 显示只读的建议消息
                Type.Missing,
                CheckString(delimiter) ? 6 : Type.Missing,
                false,//则加载项将以隐藏方式打开
                false//当文件不能以可读写模式打开时,不会请求通知，并且任何打开不可用文件的尝试都将失败。
            );
            ChangeWorkbookForRPA(wb);
            return wb;
        }

        /// <summary>
        /// 连接或者打开新的Excel,可能会改变<see cref="app"/>的指向
        /// </summary>
        /// <param name="wbPath"></param>
        /// <param name="readOnly"></param>
        /// <param name="pwd"></param>
        /// <param name="delimiter"></param>
        /// <param name="writePwd"></param>
        /// <returns></returns>
        public static _Workbook AttachOrOpenWorkbook(string wbPath, bool readOnly = false, string pwd = null, string delimiter = null, string writePwd = null)
        {
            _Workbook wb = AttachWorkbook(wbPath);
            return wb ?? OpenWorkbook(wbPath, readOnly, pwd, delimiter, writePwd);
        }

        /// <summary>
        /// 目前支持xlsx,xls,csv,html,txt,xml,dif除此之外默认txt
        /// </summary>
        /// <param name="wbPath"></param>
        /// <returns></returns>
        public static XlFileFormat GetXlFileFormatByWbPath(string wbPath)
        {
            wbPath = Path.GetFullPath(wbPath);
            string ext = Path.GetExtension(wbPath).ToLower();

            switch (ext)
            {
                case ".xlsx":
                    return XlFileFormat.xlWorkbookDefault;
                case ".xls":
                    return XlFileFormat.xlWorkbookNormal;
                case ".xlsxm":
                    return XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                case ".csv":
                    return XlFileFormat.xlCSV;
                case ".html":
                    return XlFileFormat.xlHtml;
                case ".txt":
                    return XlFileFormat.xlUnicodeText;
                case ".xml":
                    return XlFileFormat.xlXMLSpreadsheet;
                case ".dif":
                    return XlFileFormat.xlDIF;
                default:
                    return XlFileFormat.xlUnicodeText;
            }
        }

        /// <param name="wbPath">工作簿路径, 如果为空视为获取活动工作簿</param>
        /// <param name="wsName">工作表名称, 如果为空视为获取活动工作表</param>
        public ExcelAction(string wbPath = null, string wsName = null, string range = null)
            : base()
        {
            this.wbPath = wbPath;
            wbFileName = CheckString(wbPath) ? null : Path.GetFileName(wbPath);
            this.wsName = wsName;
            this.range = range;
        }

        //---------- protected ----------

        /// <summary>
        /// 工作簿路径
        /// </summary>
        protected string wbPath = null;

        /// <summary>
        /// 工作簿文件名(带后缀)
        /// </summary>
        protected string wbFileName = null;

        /// <summary>
        /// 工作表名称
        /// </summary>
        protected string wsName = null;

        /// <summary>
        /// 单元格名称
        /// </summary>
        protected string range = null;

        /// <summary>
        /// Excel应用,在<see cref="ExcelAction"/>中,任何对不是当前<see cref="app"/>或其子属性的操作都将指向新的<see cref="_Application"/>,
        /// </summary>
        protected static _Application app = null;

        /// <summary>
        /// 工作簿
        /// </summary>
        protected _Workbook wb = null;

        /// <summary>
        /// 工作表
        /// </summary>
        protected _Worksheet ws = null;


        /// <summary>
        /// 单元格
        /// </summary>
        protected Range R = null;

        /// <summary>
        /// <see cref="app"/>是否由当前的Action打开
        /// </summary>
        protected bool isOpenApp = false;

        /// <summary>
        /// <see cref="wb"/>是否由当前Action打开
        /// </summary>
        protected bool isOpenWorkbook = false;

        /// <summary>
        /// 自动连接或者打开Excel,自动获取<see cref="app"/>,<see cref="wb"/>和<see cref="ws"/>
        /// </summary>
        protected override void action()
        {
            GetWorkbook();
            GetSheet();
            GetR();
        }

        /// <summary>
        /// 自动设置<see cref="wb"/>
        /// </summary>
        protected void GetWorkbook()
        {
            isOpenApp = AttachApp() == null;
            if (CheckString(wbPath))
            {
                wb = AttachWorkbook(wbPath);
                if (wb == null)
                {
                    wb = OpenWorkbook(wbPath);
                    isOpenWorkbook = true;
                }
            }
            else
            {
                AttachOrOpenApp();
                if (app.Workbooks.Count > 0)
                {
                    wb = app.ActiveWorkbook;
                    wbPath = wb.FullName;
                    wbFileName = CheckString(wbPath) ? null : Path.GetFileName(wbPath);
                }
                else
                {
                    throw new ActionException("找不到活动工作簿");
                }
            }
            wb.Activate();
        }

        /// <summary>
        /// 自动设置<see cref="ws"/>
        /// </summary>
        protected void GetSheet()
        {
            if (CheckString(wsName))
            {
                try
                {
                    ws = wb.Worksheets[wsName];
                }
                catch (COMException)
                {
                    throw new ActionException(string.Format("在工作簿({0})中没有找到工作表({1})", wbPath, wsName));
                }
            }
            else
            {
                ws = wb.ActiveSheet;
                wsName = ws.Name;
            }
            ws.Activate();
        }

        /// <summary>
        /// 自动设置<see cref="R"/>
        /// </summary>
        protected void GetR()
        {
            if (CheckString(range))
            {
                switch (range)
                {
                    case "used":
                        R = ws.UsedRange;
                        break;
                    default:
                        R = app.Range[range];
                        break;
                }
            }
            else
            {
                dynamic r = app.Selection;
                if (r is Range)
                {
                    R = r;
                }
            }
        }

        //---------- private ----------

        /// <summary>
        /// Workbook连接方案一
        /// </summary>
        /// <returns></returns>
        private static _Workbook IAttachWorkbook1(string wbPath)
        {
            _Workbook wb = null;
            string wbFileName = Path.GetFileName(wbPath);
            try
            {
                wb = app.Workbooks[wbFileName];
            }
            catch (Exception) { }
            if (wb != null)
            {
                if (wb.FullName == wbPath)
                {
                    return wb;
                }
                else
                {
                    wb = null;
                }
            }
            return wb;
        }

        /// <summary>
        /// Workbook连接方案二
        /// </summary>
        /// <param name="wbPath"></param>
        /// <returns></returns>
        private static _Workbook IAttachWorkbook2(string wbPath)
        {
            _Workbook _wb;
            dynamic wb = null;
            uint OBJID_NATIVEOM = Convert.ToUInt32("FFFFFFF0", 16);
            Guid IID_DISPATCH = new Guid("00020400-0000-0000-C000-000000000046");
            IntPtr XLhwnd = IntPtr.Zero;
            do
            {
                //---------------
                XLhwnd = FindWindowEx(IntPtr.Zero, XLhwnd, "XLMAIN", null);
                if (IntPtr.Zero.Equals(XLhwnd))
                {
                    throw new Exception(string.Format("沒有找到已經打開的工作簿({0})", wbPath));
                }
                IntPtr XLDESKhwnd = FindWindowEx(XLhwnd, IntPtr.Zero, "XLDESK", null);
                IntPtr WBhwnd = FindWindowEx(XLDESKhwnd, IntPtr.Zero, "EXCEL7", null);
                AccessibleObjectFromWindow(WBhwnd, OBJID_NATIVEOM, ref IID_DISPATCH, ref wb);
                //----------------
                _wb = (Workbook)wb.ActiveCell.Parent.Parent;
                if (_wb != null)
                {
                    if (_wb.FullName != wbPath)
                    {
                        continue;
                    }
                    else
                    {
                        break;
                    }
                }
            } while (true);
            return wb;
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
