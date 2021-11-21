using System;

namespace RPAAP
{
    public class RPAAPException : Exception
    {
        public RPAAPException()
        {
        }

        public RPAAPException(string message)
            : base(message)
        {
        }
    }
}
