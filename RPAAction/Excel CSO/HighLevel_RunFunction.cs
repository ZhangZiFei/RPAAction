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
        /// <summary>
        /// vba函数返回值
        /// </summary>
        public string Ret = "";

        /// <param name="wbPath"></param>
        /// <param name="wsName"></param>
        /// <param name="VBACode"></param>
        /// <param name="MacroName">默认执行名称为"m"的宏</param>
        public HighLevel_RunFunction(string wbPath = null, string wsName = null, string VBACode = null, string FunctionName = null,
            string param1 = "", string param2 = "", string param3 = "", string param4 = "", string param5 = "",
            string param6 = "", string param7 = "", string param8 = "", string param9 = "", string param10 = ""
        )
            : base(wbPath, wsName)
        {
            this.VBACode = VBACode;
            this.FunctionName = FunctionName.CheckNoVoid() ? FunctionName : "f";
            this.param1 = param1.CheckNoVoid() ? param1 : Type.Missing;
            this.param2 = param2.CheckNoVoid() ? param2 : Type.Missing;
            this.param3 = param3.CheckNoVoid() ? param3 : Type.Missing;
            this.param4 = param4.CheckNoVoid() ? param4 : Type.Missing;
            this.param5 = param5.CheckNoVoid() ? param5 : Type.Missing;
            this.param6 = param6.CheckNoVoid() ? param6 : Type.Missing;
            this.param7 = param7.CheckNoVoid() ? param7 : Type.Missing;
            this.param8 = param8.CheckNoVoid() ? param8 : Type.Missing;
            this.param9 = param9.CheckNoVoid() ? param9 : Type.Missing;
            this.param10 = param10.CheckNoVoid() ? param10 : Type.Missing;
            Run();
        }

        protected override void Action()
        {
            base.Action();
            //运行宏
            if (FunctionName.CheckNoVoid())
            {
                Wb.Activate();
                try
                {
                    RunVBA();
                }
                //沒有信任存取VAB專案物件模型
                catch (System.Runtime.InteropServices.COMException come)
                {
                    //插入VBA代码
                    if (VBACode.CheckNoVoid())
                    {
                        try
                        {
                            VBE vbe = App.VBE;
                            VBComponent vbComponent;
                            vbComponent = Wb.VBProject.VBComponents.Add(vbext_ComponentType.vbext_ct_StdModule);
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
            Ret = App.Run($@"'{WbFileName}'!{FunctionName}",
                param1, param2, param3, param4, param5,
                param6, param7, param8, param9, param10
            );
            if (Ret.StartsWith("Error"))
            {
                throw new Base.ActionException(Ret);
            }
        }

        private readonly string VBACode = null;
        private readonly string FunctionName = null;

        private readonly object param1;
        private readonly object param2;
        private readonly object param3;
        private readonly object param4;
        private readonly object param5;
        private readonly object param6;
        private readonly object param7;
        private readonly object param8;
        private readonly object param9;
        private readonly object param10;
    }
}
