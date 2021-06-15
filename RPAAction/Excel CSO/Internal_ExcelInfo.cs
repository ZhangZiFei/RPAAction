using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 内部-ExcelInfo
    /// 自动打开Excel并且获取想相关COM对象和相关信息
    /// </summary>
    public class Internal_ExcelInfo : ExcelAction
    {
        /// <param name="wbPath">工作簿路径, 如果为空视为获取活动工作簿</param>
        /// <param name="wsName">工作表名称, 如果为空视为获取活动工作表</param>
        public Internal_ExcelInfo(string wbPath = null, string wsName = null, string range = null)
            : base(wbPath, wsName, range)
        {
        }

        public _Application App => ExcelAction.app;

        public _Workbook Wb => base.wb;

        public _Worksheet Ws => base.ws;

        public new Range R => base.R;

        /// <summary>
        /// 工作簿路径
        /// </summary>
        public string WbPath => base.wbPath;

        /// <summary>
        /// 工作簿文件名(带后缀)
        /// </summary>
        public string WbFileName => base.wbFileName;

        /// <summary>
        /// 工作表名称
        /// </summary>
        public string WsName => base.wsName;

        /// <summary>
        /// 单元格名称
        /// </summary>
        public new string Range => base.range;

        /// <summary>
        /// <see cref="ExcelAction.app"/>是否由当前的Action打开
        /// </summary>
        public bool IsOpenApp => base.isOpenApp;

        /// <summary>
        /// <see cref="Wb"/>是否由当前Action打开
        /// </summary>
        public bool IsOpenWorkbook => base.isOpenWorkbook;
    }
}
