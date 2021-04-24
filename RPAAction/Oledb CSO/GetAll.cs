using System.Data;
using System.Data.OleDb;

namespace RPAAction.Oledb_CSO
{
    /// <summary>
    /// 执行sql并获取结果和行数
    /// </summary>
    class GetAll : OledbAction
    {
        /// <summary>
        /// 结果
        /// </summary>
        public DataTable Table => table;

        /// <summary>
        /// 结果行数
        /// </summary>
        public decimal Count => count;

        /// <param name="connStr">Oledb的连接字符串</param>
        /// <param name="filePath">可以用Oledb连接的文件路径,如果此参数不为 "" 或者 null 则会忽略 connStr 参数</param>
        /// <param name="sql">sql语句</param>
        public GetAll(string connStr, string filePath, string sql)
            : base(connStr, filePath, sql)
        {
        }

        protected override void action()
        {
            base.action();
            table = new DataTable();

            using (OleDbDataAdapter a = new OleDbDataAdapter(comm))
            {
                count = a.Fill(table);
            }
        }
        private DataTable table;
        private decimal count;
    }
}
