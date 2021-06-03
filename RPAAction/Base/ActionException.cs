using System;

namespace RPAAction.Base
{
    public class ActionException : Exception
    {
        public ActionException()
            : base()
        {
        }

        public ActionException(string message)
            : base(message)
        {
        }

        public ActionException(string message, Exception innerException)
            : base(message, innerException)
        {
        }
    }
}
