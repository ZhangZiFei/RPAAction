using System.IO;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 进程-打开工作簿
    /// 如果工作簿已经打开则会关闭工作簿(且不会保存)并重新打开
    /// </summary>
    public class Process_OpenWorkbook : ExcelAction
    {
        /// <param name="wbPath">工作簿路径</param>
        /// <param name="readOnly">是否以只读模式打开,默认false</param>
        /// <param name="pwd">密码</param>
        /// <param name="delimiter">分隔符号,请在打开文本文件时放置此参数,默认为Tab符号</param>
        /// <param name="writeResPassword">写权限密码,有的工作簿是受到保护,需要密码才能回去写入权限</param>
        public Process_OpenWorkbook(string wbPath, bool readOnly = false, string pwd = null, string delimiter = null, string writePwd = null)
        {
            this.wbPath = Path.GetFullPath(wbPath);
            this.readOnly = readOnly;
            this.pwd = pwd;
            this.delimiter = delimiter;
            this.writePwd = writePwd;
            Run();
        }

        protected override void Action()
        {
            wb = AttachWorkbook(wbPath);
            if (wb != null)
            {
                AttachOrOpenApp();
                bool Visible = app.Visible;
                wb.Close(false);
                app.Visible = Visible;
            }
            wb = OpenWorkbook(wbPath, readOnly, pwd, delimiter, writePwd);
        }

        private readonly bool readOnly;
        private readonly string pwd;
        private readonly string delimiter;
        private readonly string writePwd;
    }
}
