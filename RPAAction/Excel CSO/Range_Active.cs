namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 单元格-聚焦
    /// </summary>
    public class Range_Active : ExcelAction
    {
        public Range_Active(string wbPath = null, string wsName = null, string range = null)
            : base(wbPath, wsName)
        {
            this.Range = range;
            Run();
        }

        protected override void Action()
        {
            base.Action();
            Wb.Activate();
            Ws.Select();
            if (Range != null && (!Range.Equals("")))
            {
                App.Range[Range].Select();
            }
        }
    }
}
