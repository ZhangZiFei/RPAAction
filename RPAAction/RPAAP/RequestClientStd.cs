using System.Diagnostics;
using Newtonsoft.Json;
using System.IO;
using System;

namespace RPAAP
{
    /// <summary>
    /// RPA Action 标准输入输出 请求端
    /// </summary>
    public class RequestClientStd : RequestClient
    {
        public RequestClientStd(string processPath)
        {
            Process = new Process()
            {
                StartInfo = new ProcessStartInfo(processPath)
                {
                    UseShellExecute = true,//不使用系統shell啟動
                    RedirectStandardInput = true,//接受調用持續的輸入
                    RedirectStandardOutput = true,//調用程序可獲取輸出
                    RedirectStandardError = true,//重定向標準錯誤輸出
                    CreateNoWindow = true//不使用window窗口打開
                }
            };
            Process.StartInfo.StandardOutputEncoding = Tool.DefEncoding;
            Process.StartInfo.StandardErrorEncoding = Tool.DefEncoding;
            Process.StartInfo.UseShellExecute = false;
            Process.Start();
            ProcessWriter = new StreamWriter(Process.StandardInput.BaseStream, Tool.DefEncoding, 4096);
        }

        ~RequestClientStd()
        {
            Dispose();
        }

        protected override ResponseData Request(RequestData requestData)
        {
            ProcessWriter.WriteLine(JsonConvert.SerializeObject(requestData));
            ProcessWriter.Flush();
            var s = Process.StandardOutput.ReadLine();
            if (s != null)
            {
                return JsonConvert.DeserializeObject<ResponseData>(s);
            }
            else
            {
                Process.StandardError.ReadToEnd();
                throw new Exception("Action意外结束");
            }
        }

        public override void Dispose()
        {
            if (Process != null)
            {
                ProcessWriter.WriteLine();
                ProcessWriter.Flush();
                Process.StandardOutput.ReadLine();

                ProcessWriter.Close();
                ProcessWriter.Dispose();
                ProcessWriter = null;

                Process.Close();
                Process.Dispose();
                Process = null;
            }
        }

        private Process Process;
        private StreamWriter ProcessWriter;
    }
}
