using Microsoft.Vbe.Interop;
using RPAAction.Base;
using System;

namespace RPAAction.Excel_CSO
{
    /// <summary>
    /// 高阶-运行Excel函数
    /// </summary>
    public class HighLevel_RunFunction : ExcelAction
    {
        public static void f()
        {
            new HighLevel_RunFunction();
        }

        /// <summary>
        /// vba函数返回值
        /// </summary>
        public string Ret = "";

        /// <param name="wbPath"></param>
        /// <param name="wsName"></param>
        /// <param name="VBACode"></param>
        /// <param name="MacroName">默认执行名称为"m"的宏</param>
        public HighLevel_RunFunction(string wbPath = null, string wsName = null, string VBACode = null, string MacroName = null,
            object param1 = null, object param2 = null, object param3 = null, object param4 = null, object param5 = null,
            object param6 = null, object param7 = null, object param8 = null, object param9 = null, object param10 = null
        )
            : base(wbPath, wsName)
        {
            this.VBACode = VBACode;
            this.MacroName = CheckString(MacroName) ? MacroName : "m";
            this.params1 = params1;
            this.params2 = params2;
            this.params3 = params3;
            this.params4 = params4;
            this.params5 = params5;
            this.params6 = params6;
            this.params7 = params7;
            this.params8 = params8;
            this.params9 = params9;
            this.params10 = params10;
            Run();
        }

        protected override void Action()
        {
            base.Action();
            //运行宏
            if (!CheckString(MacroName))
            {
                wb.Activate();
                try
                {
                    RunVBA();
                }
                //沒有信任存取VAB專案物件模型
                catch (System.Runtime.InteropServices.COMException come)
                {
                    //插入VBA代码
                    if (!CheckString(VBACode))
                    {
                        try
                        {
                            VBE vbe = app.VBE;
                            VBComponent vbComponent;
                            vbComponent = wb.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
                            vbComponent.CodeModule.AddFromString(VBACode);
                        }
                        catch (Exception e)
                        {
                            throw new ActionException("添加vb函數失敗\n" + e.ToString());
                        }
                        RunVBA();
                    }
                    else
                    {
                        throw come;
                    }
                }
            }
        }

        private void RunVBA()
        {
            Ret = app.Run($@"'{wbFileName}'!{MacroName}",
                params1 == null ? Type.Missing, params1,
                params2 == null ? Type.Missing, params2,
                params3 == null ? Type.Missing, params3,
                params4 == null ? Type.Missing, params4,
                params5 == null ? Type.Missing, params5,
                params6 == null ? Type.Missing, params6,
                params7 == null ? Type.Missing, params7,
                params8 == null ? Type.Missing, params8,
                params9 == null ? Type.Missing, params9,
                params10 == null ? Type.Missing, params10
            );
            if (Ret.StartsWith("Error"))
            {
                throw new Base.ActionException(Ret);
            }
        }

        private readonly string VBACode = null;
        private readonly string MacroName = null;

        private readonly object params1;
        private readonly object params2;
        private readonly object params3;
        private readonly object params4;
        private readonly object params5;
        private readonly object params6;
        private readonly object params7;
        private readonly object params8;
        private readonly object params9;
        private readonly object params10;
    }
}
