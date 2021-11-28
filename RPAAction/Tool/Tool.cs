using System;
using System.Diagnostics;

namespace RPAAction
{
    static class Tool
    {
        /// <summary>
        /// 版本号
        /// </summary>
        public static Version Version => version;
        public const string Version_s = "0.1.1.2";
        private static readonly Version version = new Version(Version_s);

        /// <summary>
        /// 检查是不是有效字符串,如果是null或者空字符串则返回False
        /// </summary>
        /// <param name="s"></param>
        /// <returns></returns>
        public static bool CheckNoVoid(this string s)
        {
            return ! string.IsNullOrEmpty(s);
        }

        /// <summary>
        /// 等待进程退出,否则杀死进程
        /// </summary>
        /// <param name="p"></param>
        /// <param name="milliseconds">等待关联进程退出的时间(以毫秒为单位)。 0 值指定立即返回，而 -1 值则指定无限期等待。</param>
        public static void ExitOrKill(this Process p, int milliseconds = 100)
        {
            if (p.WaitForExit(milliseconds))
            {
                return;
            }
            p.Kill();
            p.WaitForExit(milliseconds);
        }
    }
}
