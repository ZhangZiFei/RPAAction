namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 工作表-聚焦
    /// </summary>
    public class Worksheet_Active : Range_Active
    {
        public Worksheet_Active(string wbPath = null, string wsName = null)
            : base(wbPath, wsName)
        {

        }
    }
}
