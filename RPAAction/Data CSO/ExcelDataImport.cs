using Microsoft.Office.Interop.Excel;
using RPAAction.Excel_CSO;
using System.Collections.Generic;
using System.Data.Common;

namespace RPAAction.Data_CSO
{
    public class ExcelDataImport : RPADataImport
    {
        /// <param name="ExcelPath"></param>
        /// <param name="Sheet"></param>
        /// <param name="range">如果忽略视为"A1"</param>
        /// <param name="MaxCashCount"></param>
        public ExcelDataImport(string ExcelPath = null, string Sheet = null, string range = "A1", bool withTitle = true, int MaxCashCount = 10000)
        {
            eInfo = new Internal_ExcelInfo(ExcelPath, Sheet, range)
            {
                CreateWorkbook = true,
                CreateWorksheet = true
            };
            eInfo.Run();
            this.withTitle = withTitle;
            this.MaxCashCount = MaxCashCount;
        }

        public override void Dispose()
        {
            PushCash();
            eInfo.wb.Save();
            if (eInfo.isOpenWorkbook)
            {
                eInfo.wb.Close();
                if (eInfo.isOpenApp)
                {
                    new Process_Close();
                }
            }
        }

        protected override void CreateTable(DbDataReader r)
        {
            FieldCount = r.FieldCount;
            range = eInfo.R.Resize[EachCashRow, FieldCount];
            cash = new object[EachCashRow, FieldCount];
            for (int i = 0; i < FieldCount; i++)
            {
                Fields.Add(r.GetName(i), i);
            }
            //标题
            if (withTitle)
            {
                foreach (var item in Fields)
                {
                    SetValue(item.Key, item.Key);
                }
                UpdataRow();
            }
        }

        protected override void SetValue(string field, object value)
        {
            cash[CashWriteRow, Fields[field]] = value;
        }

        protected override void UpdataRow()
        {
            if (++CashWriteRow >= EachCashRow)
            {
                PushCash();
            }
        }

        private readonly Dictionary<string, int> Fields = new Dictionary<string, int>();

        /// <summary>
        /// Excel信息
        /// </summary>
        private readonly Internal_ExcelInfo eInfo;

        /// <summary>
        /// 是否需要写入标题
        /// </summary>
        private readonly bool withTitle;

        /// <summary>
        /// 数据宽度
        /// </summary>
        private int FieldCount;

        #region 緩存

        /// <summary>
        /// 最大缓存数据量
        /// </summary>
        private readonly int MaxCashCount;

        /// <summary>
        /// 每次缓存行数
        /// </summary>
        private int EachCashRow => MaxCashCount / FieldCount;

        /// <summary>
        /// 缓存的写入行
        /// </summary>
        private int CashWriteRow = 0;

        /// <summary>
        /// 缓存行数
        /// </summary>
        private object[,] cash = null;

        private Range range = null;

        /// <summary>
        /// 推送缓存
        /// </summary>
        private void PushCash()
        {
            //写入缓存
            range.Value[10] = cash;
            //刷新缓存
            cash = new object[EachCashRow, FieldCount];
            range = range.Offset[EachCashRow];
            CashWriteRow = 0;
        }

        #endregion
    }
}