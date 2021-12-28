using RPAAction.Data_CSO;
using SAP.Middleware.Connector;
using System;

namespace RPAAction.CSRFC
{

    /// <summary>
    /// https://www.w3xue.com/exp/article/202012/67010.html
    /// </summary>
    public static class CSRFC
    {
        /// <summary>
        /// 配置RFC
        /// </summary>
        /// <param name="host">主机地址，如果注主机地址为空则忽略功能名称外的所有参数，忽略的参数会采用上一次设置的参数</param>
        /// <param name="user">用戶名</param>
        /// <param name="pwd">密碼</param>
        /// <param name="client">客戶端</param>
        /// <param name="language">語言,默認EN</param>
        /// <param name="timeout">超時時間(單位:秒),默認600</param>
        /// <param name="name">名稱,默認"CNQ"</param>
        /// <param name="funcName">RFC函數名稱</param>
        public static void SetRFCConf(string host, string user, string pwd, string client, string language, string timeout, string name, string funcName)
        {
            if (!string.IsNullOrEmpty(host))
            {
                //更新SAP连接
                RfcConfigParameters parms = new RfcConfigParameters
                {
                    { RfcConfigParameters.AppServerHost, host }, //SAP主机IP
                    { RfcConfigParameters.SystemNumber, "0" },  //SAP实例
                    { RfcConfigParameters.User, user },  //用户名
                    { RfcConfigParameters.Password, pwd },  //密码
                    { RfcConfigParameters.Client, client },  // Client 
                    { RfcConfigParameters.Language, string.IsNullOrEmpty(language) ? "EN" : language },  //登陆语言
                    { RfcConfigParameters.PoolSize, "10" },
                    { RfcConfigParameters.IdleTimeout, string.IsNullOrEmpty(timeout) ? "600" : timeout },
                    { RfcConfigParameters.Name, string.IsNullOrEmpty(name) ? "CNQ" : name }
                };
                rfcDestination = RfcDestinationManager.GetDestination(parms);
            }
            FuncName = funcName;

            SetFunc();
        }

        /// <summary>
        /// 配置RFC
        /// </summary>
        private static void SetFunc()
        {
            if (rfcDestination == null)
            {
                throw new Exception("未调用操作\"配置RFC/SetRFCConf\"");
            }
            if (!isSetRFCConf)
            {
                func = rfcDestination.Repository.CreateFunction(FuncName);
                isSetRFCConf = true;
            }
        }

        /// <summary>
        /// 设置RFC参数
        /// </summary>
        /// <param name="name">参数名称</param>
        /// <param name="value">参数值</param>
        public static void SetValue(string name, string value)
        {
            SetFunc();
            func.SetValue(name, value);
        }

        /// <summary>
        /// 执行RFC
        /// </summary>
        public static void Invoke()
        {
            SetFunc();
            func.Invoke(rfcDestination);
            isSetRFCConf = false;
        }

        /// <summary>
        /// 获取返回值
        /// </summary>
        /// <param name="name">返回值名称</param>
        /// <returns></returns>
        public static string GetValue(string name)
        {
            return func.GetValue(name).ToString();
        }

        /// <summary>
        /// 获取表导入TXT
        /// </summary>
        /// <param name="rfctable">表名称</param>
        /// <param name="path">文件路径</param>
        /// <param name="delimiter">分隔符,如果为空则判断文件后缀,csv为",",其余默认"\t"</param>
        public static void GetTableToTXT(string rfctable, string path, string delimiter = "")
        {
            RPADataExport.ImportDispose(
                new IRfcTableRPADataReader(func.GetTable(rfctable)),
                new TXTDataExport(path, delimiter)
            );
        }

        /// <summary>
        /// 获取表导入Excel
        /// </summary>
        /// <param name="rfctable"></param>
        /// <param name="ExcelPath"></param>
        /// <param name="Sheet"></param>
        public static void GetTableToExcel(string rfctable, string ExcelPath, string Sheet)
        {
            RPADataExport.ImportDispose(
                new IRfcTableRPADataReader(func.GetTable(rfctable)),
                new ExcelDataExport(ExcelPath, Sheet)
            );
        }

        /// <summary>
        /// 获取表导入SqlServer
        /// </summary>
        /// <param name="rfctable"></param>
        /// <param name="DataSource"></param>
        /// <param name="DataBase"></param>
        /// <param name="user"></param>
        /// <param name="pwd"></param>
        /// <param name="table"></param>
        /// <param name="appand">是否附加数据,默认true,否则清空表</param>
        /// <param name="timeout">超时时间(秒),默认1800(半小时),小于0时使用默认值</param>
        public static void GetTableToSqlServer(string rfctable, string DataSource, string DataBase, string user, string pwd, string table, bool appand = true, decimal timeout = 0)
        {
            if (timeout < 0)
            {
                timeout = 1800;
            }
            RPADataExport.ImportDispose(
                new IRfcTableRPADataReader(func.GetTable(rfctable)),
                new SQLServerDataExport(DataSource, DataBase, user, pwd, table, appand, (int)timeout)//超时时间半小时
            );
        }

        private static RfcDestination rfcDestination = null;
        private static IRfcFunction func;
        private static bool isSetRFCConf = false;
        private static string FuncName;
    }
}
