using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using RPAAction.Excel_CSO;

namespace RPAAction.Data_CSO
{
    public class ExcelDataReader : RPADataReader
    {
        public override bool IsClosed => isClosed;

        public override int FieldCount => _FieldCount;

        /// <param name="ExcelPath"></param>
        /// <param name="Sheet"></param>
        /// <param name="range">如果忽略视为已使用的单元格</param>
        public ExcelDataReader(string ExcelPath = null, string Sheet = null, string range = "used", int MaxCashCount = 10000)
            : base()
        {
            //准备Excel
            eInfo = new Internal_ExcelInfo(ExcelPath, Sheet, range);

            this.MaxCashCount = MaxCashCount;
            _FieldCount = eInfo.ws.UsedRange.Columns.Count;
            _RowCont = eInfo.ws.UsedRange.Rows.Count - 1;

            //获取标题
            if (HasTitle)
            {
                FieldValues = ((Range)eInfo.ws.UsedRange.Rows[1]).Value[10];
            }
        }

        public override void Close()
        {
            if (! isClosed)
            {
                if (eInfo.isOpenApp)
                {
                    new Process_Close();
                }
                else if (eInfo.isOpenWorkbook)
                {
                    eInfo.wb.Close(false);
                }
                isClosed = true;
            }
        }

        public override string GetName(int ordinal)
        {
            return FieldValues[1, ordinal + 1].ToString(); 
        }

        public override object GetValue(int ordinal)
        {
            return cache[readingRow - cacheRowBase, ordinal + 1];
        }

        public override bool Read()
        {
            ++readRow;
            if (CanRead)
            {
                if (needCashe)
                    ReadCashe();
                return true;
            }
            else
                return false;
        }

        private bool isClosed = false;

        private readonly Internal_ExcelInfo eInfo;

        /// <summary>
        /// 标题长度
        /// </summary>
        private readonly int _FieldCount;

        /// <summary>
        /// 数据行数,如果为负数,表示没有标题
        /// </summary>
        private readonly int _RowCont;

        /// <summary>
        /// 标题
        /// </summary>
        private readonly Object[,] FieldValues = null;

        /// <summary>
        /// 已经读取的行数,外部调用<see cref="Read"/>方法的次数
        /// </summary>
        private int readRow = -1;

        /// <summary>
        /// 正在读取的行数
        /// </summary>
        private int readingRow => readRow + 1;

        /// <summary>
        /// 数据是否存在标题
        /// </summary>
        /// <returns></returns>
        private bool HasTitle => _RowCont >= 0;

        /// <summary>
        /// 是否存在可以读取的数据
        /// </summary>
        /// <returns></returns>
        private bool CanRead => _RowCont - readingRow >= 0;

        #region 緩存

        /// <summary>
        /// 最大緩存读取数量
        /// </summary>
        private readonly int MaxCashCount;

        /// <summary>
        /// 已经缓存的行数
        /// </summary>
        private int cacheRow = 0;

        /// <summary>
        /// 已经缓存的行数的基數
        /// </summary>
        private int cacheRowBase = 0;

        /// <summary>
        /// 缓存数据
        /// </summary>
        private Object[,] cache;

        /// <summary>
        /// 每次缓存读取的行数
        /// </summary>
        private int EachReadRow => MaxCashCount / _FieldCount;

        /// <summary>
        /// 是否存在可以緩存的數據
        /// </summary>
        /// <returns></returns>
        private bool CanCashe => _RowCont - cacheRow > 0;

        /// <summary>
        /// 是否需要讀取緩存
        /// </summary>
        private bool needCashe => readingRow > cacheRow;

        /// <summary>
        /// 读取缓存
        /// </summary>
        private void ReadCashe()
        {
            if (CanCashe)
            {
                cacheRowBase = cacheRow;
                cacheRow = cacheRowBase + EachReadRow;
                if (cacheRow > _RowCont)
                    cacheRow = _RowCont;
                cache = ((Range)eInfo.ws.UsedRange.Rows[(cacheRowBase + 2) + ":" + (cacheRow + 2)]).Value[10];
            }
        }

        #endregion
    }
}
