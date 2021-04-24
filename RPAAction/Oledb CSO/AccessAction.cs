using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace RPAAction.Oledb_CSO
{
    abstract class OledbAction : Base.Action
    {
        /// <summary>
        /// 释放指定的连接,如果参数 connStr 为 "" 或者 <see cref="null"/> 则视为释放所有连接
        /// </summary>
        /// <param name="connStr">连接字符串</param>
        static public void CloseConn(string connStr = null)
        {
            if (connStr == null || object.Equals("", connStr))
            {
                foreach (var item in connMap)
                {
                    item.Value.Close();
                    item.Value.Dispose();
                }
                connMap = new Dictionary<string, OleDbConnection>();
            }
            else
            {
                OleDbConnection conn = connMap[connStr];
                conn.Close();
                conn.Dispose();
                connMap.Remove(connStr);
            }
        }

        static public _Application AttachOrOpenExcel()
        {
            _Application app;
            try
            {
                //连接Excel进程
                app = (_Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                app = new Application();
                app.Visible = true;
                app.UserControl = true;
            }
            return app;
        }

        /// <summary>
        /// 打开或者连接Excel工作簿
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <returns></returns>
        static public _Workbook AttachOrOpenExcelWorkbook(string excelPath)
        {
            _Application app = AttachOrOpenExcel();
            _Workbook wb = null;
            //试图连接
            if (excelPath == null || object.Equals("", excelPath))
            {
                wb = app.ActiveWorkbook;
            }
            else
            {
                excelPath = Path.GetFullPath(excelPath);
                string wbName = Path.GetFileName(excelPath);
                wb = (_Workbook)app.Workbooks[wbName];
                if (!object.Equals(wb.FullName, excelPath))
                {
                    throw new Exception(string.Format(@"Excel无法打开两个名称相同的工作簿(试图打开--{0};已经打开--{1})", excelPath, wb.FullName));
                }
            }
            //打开
            if (wb == null)
            {
                wb = app.Workbooks.Open(excelPath);
            }
            return wb;
        }

        /// <param name="connStr">Oledb的连接字符串</param>
        /// <param name="filePath">可以用Oledb连接的文件路径(目前仅支持accdb格式的Access文件),如果此参数不为 "" 或者 null 则会忽略 connStr 参数</param>
        /// <param name="sql">sql语句</param>
        public OledbAction(string connStr, string filePath, string sql)
        {
            if (filePath == null || object.Equals("", filePath))
            {
                this.connStr = connStr;
            }
            else
            {
                this.connStr = filePathToConnStr(filePath);
            }
            this.sql = sql;
        }

        ~OledbAction()
        {
            if (comm != null)
            {
                comm.Dispose();
            }
        }

        protected OleDbConnection conn;
        protected OleDbCommand comm;

        protected override void action()
        {
            //获取 conn
            if (!connMap.TryGetValue(connStr, out conn))
            {
                conn = new OleDbConnection(connStr);
                conn.Open();
                connMap.Add(connStr, conn);
            }
            //获取comm
            if (sql == null || object.Equals("", sql))
            {
                comm = conn.CreateCommand();
            }
            else
            {
                comm = new OleDbCommand(sql, conn);
            }
        }

        /// <summary>
        /// 数据库连接
        /// </summary>
        static private Dictionary<string, OleDbConnection> connMap = new Dictionary<string, OleDbConnection>();

        private readonly string connStr;
        private readonly string sql;

        /// <summary>
        /// 将支持Oledb连接的文件路径转为连接字符串(目前仅支持accdb格式的Access文件)
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>转换好的连接字符串</returns>
        private string filePathToConnStr(string filePath)
        {
            Regex check_accdb = new Regex(@"\.accdb$", RegexOptions.IgnoreCase);

            if (check_accdb.IsMatch(filePath))
            {
                return string.Format(@"Provider=Microsoft.ACE.OLEDB.16.0;Data Source={0};", filePath);
            }
            else
            {
                throw new Exception(string.Format(@"不支持将该文件({0})转换为Oledb的连接字符串。", filePath));
            }
        }
    }
}
