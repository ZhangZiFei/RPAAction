using System;

namespace RPAAction.Base
{
    public abstract class RPAAction
    {
        /// <summary>
        /// 该<see cref="Base.RPAAction"/>是否已经运行过
        /// </summary>
        public bool IsRan { get => isRan; }

        public string EType { get => e == null ? "" : e.GetType().ToString(); }
        public string EMes { get => e == null ? "" : e.ToString(); }

        /// <summary>
        /// 检查是不是有效字符串,如果是null或者空字符串则返回False
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool CheckString(string s)
        {
            return s != null && (!s.Equals(""));
        }

        public RPAAction Run()
        {
            if (!isRan)
            {
                try
                {
                    BeforeRun();
                    Action();
                }
#if !DEBUG
                catch (Exception e)
                {
                    this.e = e;
                }
#endif
                finally
                {
#if !DEBUG
                    try
                    {
#endif
                        AfterRun();
#if !DEBUG
                    }
                    catch (Exception) { }
#endif
                }
                isRan = true;
            }
            return this;
        }

        public RPAAction OutE(out string EType, out string EMes)
        {
            EType = this.EType;
            EMes = this.EMes;
            return this;
        }

        /// <summary>
        /// Action的实现内容,按照规范,类中所有的存在副作用的代码均需要在这里实现
        /// </summary>
        protected abstract void Action();

        protected virtual void BeforeRun()
        {

        }
        protected virtual void AfterRun()
        {

        }

        private bool isRan;
#if DEBUG
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE0044:添加只读修饰符", Justification = "<挂起>")]
#endif
        private Exception e = null;
    }
}