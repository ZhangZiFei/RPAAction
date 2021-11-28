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
            switch (requestData.ObjectName)
            {
                case "Excel CSO":
                    switch (requestData.Action)
                    {
                        #region 高阶-运行Excel函数 HighLevel_RunFunction
                        case "HighLevel_RunFunction":
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
                            return new ResponseData(new Dictionary<string, Param>() {
                                { "Ret", new Param(highLevel_RunFunction.Ret)}
                            });
                        #endregion
                        #region 单元格-写入集合 Range_WriteToDataTable
                        case "Range_WriteToDataTable":
                            RPADataImport.ImportDispose(
                                new DataTableDataReader((System.Data.DataTable)requestData.InputParams["Table"].Value),
                                new ExcelDataImport(
                                    requestData.InputParams["ExcelPath"].Value as string,
                                    requestData.InputParams["Sheet"].Value as string,
                                    requestData.InputParams["range"].Value as string,
                                    (bool)requestData.InputParams["withTitle"].Value
                                )
                            );
                            return new ResponseData(new Dictionary<string, Param>());
                        #endregion
                        #region 单元格-读取集合 Range_WriteFromDataTable
                        case "Range_ReadFromDataTable":
                            var table = new System.Data.DataTable();
                            RPADataImport.ImportDispose(
                                new ExcelDataReader(
                                    (string)requestData.InputParams["ExcelPath"].Value,
                                    (string)requestData.InputParams["Sheet"].Value,
                                    (string)requestData.InputParams["range"].Value
                                ),
                                new DataTableDataImport(table)
                            );
                            return new ResponseData(new Dictionary<string, Param>() {
                                { "Table", new Param(table)}
                            });
                            #endregion
                    }
                    break;
            }
            throw new Exception($"没有找到对象({requestData.ObjectName})或者操作({requestData.Action})");
        }
    }
}
