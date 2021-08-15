namespace RPAAction.Oledb_CSO
{
    /// <summary>
    /// 执行sql并获取影响行数
    /// </summary>
    public class GetAc : OledbAction
    {
        /// <summary>
        /// 影响行数
        /// </summary>
        public decimal Count => count;

        /// <param name="connStr">Oledb的连接字符串</param>
        /// <param name="filePath">可以用Oledb连接的文件路径,如果此参数不为 "" 或者 null 则会忽略 connStr 参数</param>
        /// <param name="sql">sql语句</param>
        public GetAc(string connStr, string filePath, string sql)
            : base(connStr, filePath, sql)
        {
        }

        protected override void Action()
        {
            base.Action();
            count = comm.ExecuteNonQuery();
        }

        private decimal count;
    }
}
