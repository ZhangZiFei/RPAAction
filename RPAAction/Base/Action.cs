using System;

namespace RPAAction.Base
{
    public abstract class Action
    {
        /// <summary>
        /// 该<see cref="Action"/>是否已经运行过
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

        public Action Run()
        {
            if (!isRan)
            {
#if DEBUG
                action();
#else
                try
                {
                    action();
                }
                catch (Exception e)
                {
                    this.e = e;
                }
#endif
                isRan = true;
            }
            return this;
        }

        /// <summary>
        /// Action的实现内容,按照规范,类中所有的存在副作用的代码均需要在这里实现
        /// </summary>
        abstract protected void action();

        private bool isRan;
        private Exception e = null;
    }
}