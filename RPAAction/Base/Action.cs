using System;

namespace RPAAction.Base
{
    public abstract class RPAAction
    {
        /// <summary>
        /// 该<see cref="Base.RPAAction"/>是否已经运行过
        /// </summary>
        public bool IsRan { get => isRan; }

        public RPAAction Run()
        {
            if (!isRan)
            {
                try
                {
                    BeforeRun();
                    Action();
                }
                finally
                {
                        AfterRun();
                }
                isRan = true;
            }
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
    }
}