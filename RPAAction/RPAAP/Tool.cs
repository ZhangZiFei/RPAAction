using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RPAAP
{
    static class Tool
    {
        public static string ExePath = @"E:\zifeiobject\RPAAP\RPAAP\bin\Debug\ResponseClientStdTest.exe";
        public static RequestClientStd R
        {
            get
            {
                if (r == null)
                {
                    if (!File.Exists(ExePath))
                    {
                        throw new Exception("没有找到ExePath(" + ExePath + ")文件");
                    }
                    r = new RequestClientStd(ExePath);
                }
                return r;
            }
            set
            {
                if (value == null)
                {
                    using (r) { };
                    r = value;
                }
            }
        }

        /// <summary>
        /// 版本号
        /// </summary>
        public static Version Version => version;
        public const string Version_s = "0.1.0.0";
        public static Encoding DefEncoding => defEncoding;

        private static RequestClientStd r;
        private static readonly Version version = new Version(Version_s);
        public static Encoding defEncoding = new UTF8Encoding(false);
    }
}
