using RPAAction.Data_CSO;
using RPAAction.Excel_CSO;
using RPAAP;
using System;
using System.Collections.Generic;

namespace RPAAction
{
    class RPAAction : ResponseClientStd
    {
        protected override ResponseData RunAction(RequestData requestData)
        {
            if (requestData.ObjectName.Equals("Excel CSO"))
            {
                //高阶-运行Excel函数 HighLevel_RunFunction
                if (requestData.Action.Equals("HighLevel_RunFunction"))
                {
                    HighLevel_RunFunction highLevel_RunFunction = new HighLevel_RunFunction(
                        (string)requestData.InputParams["wbPath"].Value,
                        (string)requestData.InputParams["wsName"].Value,
                        (string)requestData.InputParams["VBACode"].Value,
                        (string)requestData.InputParams["FunctionName"].Value,
                        (string)requestData.InputParams["param1"].Value,
                        (string)requestData.InputParams["param2"].Value,
                        (string)requestData.InputParams["param3"].Value,
                        (string)requestData.InputParams["param4"].Value,
                        (string)requestData.InputParams["param5"].Value,
                        (string)requestData.InputParams["param6"].Value,
                        (string)requestData.InputParams["param7"].Value,
                        (string)requestData.InputParams["param8"].Value,
                        (string)requestData.InputParams["param9"].Value,
                        (string)requestData.InputParams["param10"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>()
                    {
                        { "Ret", new Param(highLevel_RunFunction.Ret)}
                    });
                }
                //单元格-写入集合 Range_WriteToDataTable
                else if (requestData.Action.Equals("Range_WriteToDataTable"))
                {
                    RPADataExport.ImportDispose(
                        new DataTableDataReader((System.Data.DataTable)requestData.InputParams["Table"].Value),
                        new ExcelDataExport(
                            requestData.InputParams["ExcelPath"].Value as string,
                            requestData.InputParams["Sheet"].Value as string,
                            requestData.InputParams["range"].Value as string,
                            (bool)requestData.InputParams["withTitle"].Value
                        )
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //单元格-读取集合 Range_WriteFromDataTable
                else if (requestData.Action.Equals("Range_WriteToDataTable"))
                {
                    var table = new System.Data.DataTable();
                    RPADataExport.ImportDispose(
                        new ExcelDataReader(
                            (string)requestData.InputParams["ExcelPath"].Value,
                            (string)requestData.InputParams["Sheet"].Value,
                            (string)requestData.InputParams["range"].Value
                        ),
                        new DataTableDataExport(table)
                    );
                    return new ResponseData(new Dictionary<string, Param>()
                    {
                        { "Table", new Param(table)}
                    });
                }
            }
            else if(requestData.ObjectName.Equals("CSRFC"))
            {
                //配置RFC SetRFCConf
                if (requestData.Action.Equals("SetRFCConf"))
                {
                    CSRFC.CSRFC.SetRFCConf(
                        (string)requestData.InputParams["host"].Value,
                        (string)requestData.InputParams["user"].Value,
                        (string)requestData.InputParams["pwd"].Value,
                        (string)requestData.InputParams["client"].Value,
                        (string)requestData.InputParams["language"].Value,
                        (string)requestData.InputParams["timeout"].Value,
                        (string)requestData.InputParams["name"].Value,
                        (string)requestData.InputParams["funcName"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //设置RFC参数
                else if (requestData.Action.Equals("SetValue"))
                {
                    CSRFC.CSRFC.SetValue(
                        (string)requestData.InputParams["name"].Value,
                        (string)requestData.InputParams["value"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //执行RFC
                else if (requestData.Action.Equals("Invoke"))
                {
                    CSRFC.CSRFC.Invoke();
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //获取返回值
                else if (requestData.Action.Equals("GetValue"))
                {
                    string value =  CSRFC.CSRFC.GetValue(
                        (string)requestData.InputParams["name"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>()
                    {
                        { "value", new Param(value)}
                    });
                }
                //获取表导入TXT
                else if (requestData.Action.Equals("GetTableToTXT"))
                {
                    CSRFC.CSRFC.GetTableToTXT(
                        (string)requestData.InputParams["rfctable"].Value,
                        (string)requestData.InputParams["path"].Value,
                        (string)requestData.InputParams["delimiter"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //获取表导入Excel
                else if (requestData.Action.Equals("GetTableToExcel"))
                {
                    CSRFC.CSRFC.GetTableToExcel(
                        (string)requestData.InputParams["rfctable"].Value,
                        (string)requestData.InputParams["ExcelPath"].Value,
                        (string)requestData.InputParams["Sheet"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
                //获取表导入SqlServer
                else if (requestData.Action.Equals("GetTableToSqlServer"))
                {
                    CSRFC.CSRFC.GetTableToSqlServer(
                        (string)requestData.InputParams["rfctable"].Value,
                        (string)requestData.InputParams["DataSource"].Value,
                        (string)requestData.InputParams["DataBase"].Value,
                        (string)requestData.InputParams["user"].Value,
                        (string)requestData.InputParams["pwd"].Value,
                        (string)requestData.InputParams["table"].Value,
                        (bool)requestData.InputParams["appand"].Value,
                        (decimal)requestData.InputParams["timeout"].Value
                    );
                    return new ResponseData(new Dictionary<string, Param>());
                }
            }
            throw new Exception($"没有找到对象({requestData.ObjectName})或者操作({requestData.Action})");
        }
    }
}
