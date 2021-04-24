using Microsoft.Office.Interop.Excel;

namespace RPAAction.Oledb_CSO
{
    /// <summary>
    /// Excel数据导入Oledb
    /// </summary>
    class ExcelImportToOledb : OledbAction
    {
        /// <param name="excelPath">Excel文件路径</param>
        /// <param name="acessPath">Access文件路径</param>
        /// <param name="sheetName">Excel的工作表名称,默认"",""视为活动工作簿</param>
        /// <param name="tableName">Access的数据表,默认"",""视为同sheetName</param>
        public ExcelImportToOledb(string excelPath, string acessPath, string sheetName = "", string tableName = "")
            : base("", acessPath, "")
        {
            this.acessPath = acessPath;
            this.excelPath = excelPath;
            this.sheetName = sheetName;
            this.tableName = tableName;
        }

        protected override void action()
        {
            base.action();
            //获取工作表
            _Worksheet ws = getSheet();
            ws.Range["A1"].Select();

            var a = new GetAll("", acessPath, "SELECT * FROM a");
            a.Run();
        }

        private readonly string acessPath;
        private string excelPath;
        private string sheetName;
        private string tableName;
        private _Workbook wb;

        private _Worksheet getSheet()
        {
            if (wb == null)
            {
                wb = AttachOrOpenExcelWorkbook(excelPath);
            }
            _Worksheet ws;
            if (sheetName == null || object.Equals("", sheetName))
            {
                ws = wb.ActiveSheet;
                //矫正sheetName
                sheetName = ws.Name;
            }
            else
            {
                ws = wb.Worksheets[sheetName];
            }
            //矫正tableName
            if (tableName == null || object.Equals("", tableName))
            {
                tableName = sheetName;
            }
            return ws;
        }
    }
}
