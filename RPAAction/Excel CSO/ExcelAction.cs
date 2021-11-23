using Microsoft.Office.Interop.Excel;
using RPAAction.Base;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace RPAAction.Excel_CSO
{
    public abstract class ExcelAction : Base.RPAAction
    {
        /// <summary>
        /// 为RPA程式创建新的Excel进程,会改变<see cref="App"/>
        /// </summary>
        /// <returns></returns>
        public static _Application CreateAppForRPA()
        {
            return new Application().ChangeForRPA();
        }

        /// <summary>
        /// 连接并且返回可用的<see cref="_Application"/>,如果连接失败返回null,可能会改变<see cref="App"/>
        /// </summary>
        public static _Application AttachApp()
        {
            if (App.Check())
            {
                return App;
            }
            else
            {
                do
                {
                    try
                    {
                        //连接Excel进程
                        App = (_Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        App = null;
                        break;
                    }
                } while (!App.Check());
                return App?.ChangeForRPA();
            }
        }

        /// <summary>
        /// 链接或者打开<see cref="_Application"/>,可能会改变<see cref="App"/>
        /// </summary>
        /// <returns></returns>
        public static _Application AttachOrOpenApp()
        {
            AttachApp();
            return App ?? CreateAppForRPA();
        }

        /// <summary>
        /// 连接工作簿,如果失败则返回null,可能会改变<see cref="App"/>
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

            return wb?.ChangeForRPA();
        }

        /// <summary>
        /// 打开工作簿,可能会改变<see cref="App"/>
        /// </summary>
        public static _Workbook OpenWorkbook(string wbPath, bool readOnly = false, string pwd = null, string delimiter = null, string writePwd = null)
        {
            AttachOrOpenApp();
            wbPath = Path.GetFullPath(wbPath);
            _Workbook wb = App.Workbooks.Open(
                wbPath,
                XlUpdateLinks.xlUpdateLinksNever,
                readOnly,
                delimiter.CheckNoVoid() ? delimiter : Type.Missing,
                pwd.CheckNoVoid() ? pwd : Type.Missing,
                writePwd.CheckNoVoid() ? writePwd : Type.Missing,
                true,//则不让 Microsoft Excel 显示只读的建议消息
                Type.Missing,
                delimiter.CheckNoVoid() ? 6 : Type.Missing,
                false,//则加载项将以隐藏方式打开
                false//当文件不能以可读写模式打开时,不会请求通知，并且任何打开不可用文件的尝试都将失败。
            );
            return wb.ChangeForRPA();
        }

        /// <summary>
        /// 连接或者打开新的Excel,可能会改变<see cref="App"/>
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
        /// 目前支持xlsx,xlsb,xls,csv,html,txt,xml,dif除此之外默认txt
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
                case ".xml":
                    return XlFileFormat.xlXMLSpreadsheet;
                case ".dif":
                    return XlFileFormat.xlDIF;
                case ".xlsb":   //Excel 二进制工作簿
                    return XlFileFormat.xlExcel12;
                case ".txt":
                default:
                    return XlFileFormat.xlUnicodeText;
            }
        }

        public static Range GetRange(_Worksheet ws, string range)
        {
            Range R = null;
            if (range.CheckNoVoid())
            {
                switch (range)
                {
                    case "used":
                        R = ws.UsedRange;
                        break;
                    default:
                        R = App.Range[range];
                        break;
                }
            }
            else
            {
                dynamic r = App.Selection;
                if (r is Range)
                {
                    R = r;
                }
            }
            return R;
        }

        /// <param name="wbPath">工作簿路径, 如果为空视为获取活动工作簿</param>
        /// <param name="wsName">工作表名称, 如果为空视为获取活动工作表</param>
        public ExcelAction(string wbPath = null, string wsName = null, string range = null)
            : base()
        {
            WbPath = wbPath;
            WsName = wsName;
            Range = range;
        }

        //---------- protected ----------

        /// <summary>
        /// 工作簿路径
        /// </summary>
        protected string WbPath
        {
            get
            {
                return wbPath;
            }
            set
            {
                wbPath = Path.GetFullPath(value);
                wbFileName = Path.GetFileName(value);
            }
        }

        /// <summary>
        /// 工作簿文件名(带后缀)
        /// </summary>
        protected string WbFileName => wbFileName;

        /// <summary>
        /// 工作表名称
        /// </summary>
        protected string WsName = null;

        /// <summary>
        /// 单元格名称
        /// </summary>
        protected string Range = null;

        /// <summary>
        /// Excel应用,在<see cref="ExcelAction"/>中,任何对不是当前<see cref="App"/>或其子属性的操作都将指向新的<see cref="_Application"/>,
        /// </summary>
        protected static _Application App = null;

        /// <summary>
        /// 工作簿
        /// </summary>
        protected _Workbook Wb = null;

        /// <summary>
        /// 工作表
        /// </summary>
        protected _Worksheet Ws = null;


        /// <summary>
        /// 单元格
        /// </summary>
        protected Range R = null;

        /// <summary>
        /// 是否需要主動創建工作簿
        /// </summary>
        protected bool CreateWorkbook = false;

        /// <summary>
        /// 是否需要主動創建工作表
        /// </summary>
        protected bool CreateWorksheet = false;

        /// <summary>
        /// <see cref="App"/>是否由当前的Action打开
        /// </summary>
        protected bool isOpenApp = false;

        /// <summary>
        /// <see cref="Wb"/>是否由当前Action打开
        /// </summary>
        protected bool isOpenWorkbook = false;

        /// <summary>
        /// <see cref="Wb"/>是否由当前Action创建
        /// </summary>
        protected bool isCreateWorkbook = false;

        /// <summary>
        /// <see cref="Ws"/>是否由当前Action创建
        /// </summary>
        protected bool isCreateWorksheet = false;

        /// <summary>
        /// 自动连接或者打开Excel,自动获取<see cref="App"/>,<see cref="Wb"/>和<see cref="Ws"/>
        /// </summary>
        protected override void Action()
        {
            SetWorkbook();
            SetSheet();
            SetR();
        }

        protected override void AfterRun()
        {
            base.AfterRun();

            //如果Excel进程有效且处于显示状态则切换为用户模式
            if (App.Check() && App.Visible == true)
            {
                App.ChangeForUser();
            }
        }

        /// <summary>
        /// 自动设置<see cref="Wb"/>
        /// </summary>
        protected void SetWorkbook()
        {
            isOpenApp = AttachApp() != null;
            if (WbPath.CheckNoVoid())
            {
                if (File.Exists(WbPath))
                {
                    Wb = AttachWorkbook(WbPath);
                    if (Wb == null)
                    {
                        Wb = OpenWorkbook(WbPath);
                        isOpenWorkbook = true;
                    }
                }
                else if (CreateWorkbook)
                {
                    Wb = new Process_CreateWorkbook(WbPath).Wb;
                    isOpenWorkbook = true;
                    isCreateWorkbook = true;
                }
                else
                {
                    throw new ActionException($"文件({WbPath})不存在");
                }
            }
            else
            {
                AttachOrOpenApp();
                if (App.Workbooks.Count > 0)
                {
                    Wb = App.ActiveWorkbook;
                    WbPath = Wb.FullName;
                }
                else
                {
                    throw new ActionException("找不到活动工作簿");
                }
            }
            Wb.Activate();
        }

        /// <summary>
        /// 自动设置<see cref="Ws"/>
        /// </summary>
        protected void SetSheet()
        {
            if (isCreateWorkbook)
            {
                Ws = Wb.Worksheets[1];
                if (WsName.CheckNoVoid())
                {
                    Ws.Name = WsName;
                }
                else
                {
                    WsName = Ws.Name;
                }
                //删除其他的工作表
                while (Wb.Worksheets.Count > 1)
                {
                    _Worksheet worksheet = Wb.Worksheets[2];
                    if (!worksheet.Name.Equals(WsName))
                    {
                        worksheet.Delete();
                    }
                }
            }
            else if (WsName.CheckNoVoid())
            {
                try
                {
                    Ws = Wb.Worksheets[WsName];
                }
                catch (COMException)
                {
                    if (CreateWorksheet)
                    {
                        Ws = new Workbook_CreateWorksheet(WbPath, WsName).Ws;
                        isCreateWorksheet = true;
                    }
                    else
                    {
                        throw new ActionException($"在工作簿({WbPath})中没有找到工作表({WsName})");
                    }
                }
            }
            else
            {
                Ws = Wb.ActiveSheet;
                WsName = Ws.Name;
            }
            Ws.Activate();
        }

        /// <summary>
        /// 自动设置<see cref="R"/>
        /// </summary>
        protected void SetR()
        {
            R = GetRange(Ws, Range);
        }

        //---------- private ----------

        /// <summary>
        /// 工作簿路径
        /// </summary>
        private string wbPath = null;

        /// <summary>
        /// 工作簿文件名(带后缀)
        /// </summary>
        private string wbFileName = null;

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
                wb = App.Workbooks[wbFileName];
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
            _Workbook _wb = null;
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
                    return null;
                }
                IntPtr XLDESKhwnd = FindWindowEx(XLhwnd, IntPtr.Zero, "XLDESK", null);
                IntPtr WBhwnd = FindWindowEx(XLDESKhwnd, IntPtr.Zero, "EXCEL7", null);
                AccessibleObjectFromWindow(WBhwnd, OBJID_NATIVEOM, ref IID_DISPATCH, ref wb);
                //----------------
                try
                {
                    _wb = (Workbook)wb.ActiveCell.Parent.Parent;
                }
                catch (Exception) { }
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
