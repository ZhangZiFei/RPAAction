namespace RPAAction.Access_CSO
{
    public class Import_Excel : Sql_GetAc
    {
        public Import_Excel(string accessPath, string excelPath, string sheet, string table)
            : base(
                  accessPath,
                  $"select * into {table} from [Excel 8.0;HDR=YES;DATABASE={excelPath};IMEX=1].[{sheet}$];"
            )
        {
        }
    }
}