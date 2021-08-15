using Microsoft.Office.Interop.Access;
using Microsoft.Office.Interop.Access.Dao;
using System.Collections.Generic;
namespace RPAAction.Access_CSO
{
    public class Internal_AccessInfo : AccessAction
    {
        public void Close()
        {
            if (!isClosed)
            {
                isClosed = true;
                base.AfterRun();
            }
        }

        public Internal_AccessInfo(string accessPath)
            : base(accessPath)
        {
            Run();
        }

        protected override void AfterRun()
        {

        }

        public _Application App => app;

        public Database Db => db;

        private bool isClosed = false;
    }
}
