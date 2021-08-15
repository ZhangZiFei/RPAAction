namespace RPAAction.Access_CSO
{
    public class Export_Excel : Sql_GetAc
    {
        public Export_Excel(string accessPath, string excelPath, string sheet, string sql)
            : base(
                  accessPath,
                  $"select * into [Excel 12.0 Xml;database={excelPath}].[{sheet}] from ({sql})"
            )
        {
        }
    }
}