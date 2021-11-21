#if DEBUG
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;

/// <summary>
/// BluePrism 全局代码
/// </summary>
namespace ForBluePrism
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

    /// <summary>
    /// RPA参数
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    public class Param
    {
        /// <summary>
        /// 参数值
        /// </summary>
        public object Value
        {
            get
            {
                switch (type)
                {
                    case "Decimal":
                        return Decimal;
                    case "String":
                        return String;
                    case "DataTable":
                        return DataTable;
                    default:
                        return null;
                }
            }
        }

        /// <summary>
        /// RPA 参数类型
        /// </summary>
        public string Type => type;

        /// <param name="value">参数值</param>
        public Param(decimal value)
        {
            type = "Decimal";
            Decimal = value;
        }

        /// <param name="value">参数值</param>
        public Param(string value)
        {
            type = "String";
            String = value;
        }

        /// <param name="value">参数值</param>
        public Param(DataTable value)
        {
            type = "DataTable";
            DataTable = value;
        }

        [JsonConstructor]
        protected Param(string type, decimal Decimal, string String, DataTable DataTable)
        {
            this.type = type;
            this.Decimal = Decimal;
            this.String = String;
            this.DataTable = DataTable;
        }

        [JsonProperty]
        private readonly string type;

        [JsonProperty]
        private readonly decimal Decimal = 0;

        [JsonProperty]
        private readonly string String = "";

        [JsonProperty]
        private readonly DataTable DataTable = null;
    }

    /// <summary>
    /// 请求数据
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    public class RequestData
    {
        /// <summary>
        /// 版本号
        /// </summary>
        public Version Version => version;

        /// <summary>
        /// RPA操作对象名称
        /// </summary>
        public string ObjectName => objectName;

        /// <summary>
        /// RPA操作名称
        /// </summary>
        public string Action => action;

        /// <summary>
        /// 输入参数
        /// </summary>
        public Dictionary<string, Param> InputParams => inputParams;

        /// <param name="objectName">RPA操作对象名称</param>
        /// <param name="action">RPA操作名称</param>
        /// <param name="inputParams">输入参数</param>
        public RequestData(string objectName, string action, Dictionary<string, Param> inputParams)
        {
            this.objectName = objectName;
            this.action = action;
            this.inputParams = inputParams;
        }

        [JsonProperty]
        private readonly Version version = Tool.Version;

        [JsonProperty]
        private readonly string objectName;

        [JsonProperty]
        private readonly string action;

        [JsonProperty]
        private readonly Dictionary<string, Param> inputParams;
    }

    /// <summary>
    /// 响应数据
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    public class ResponseData
    {
        /// <summary>
        /// 输入参数
        /// </summary>
        public Dictionary<string, Param> OutputParams => outputParams;

        /// <summary>
        /// 版本号
        /// </summary>
        public Version Version => version;

        /// <summary>
        /// 异常信息
        /// </summary>
        public string Error => error;

        /// <param name="outputParams">输出参数</param>
        public ResponseData(Dictionary<string, Param> outputParams, string error = "")
        {
            this.outputParams = outputParams;
            this.error = error;
        }

        [JsonProperty]
        private readonly Version version = Tool.Version;

        [JsonProperty]
        private readonly Dictionary<string, Param> outputParams;

        [JsonProperty]
        private readonly string error;
    }


    /// <summary>
    /// RPA Action请求端
    /// </summary>
    public abstract class RequestClient : System.IDisposable
    {
        public RequestClient()
        {

        }

        /// <summary>
        /// 请求执行Action
        /// </summary>
        /// <param name="objectName">对象名称</param>
        /// <param name="action">操作名称</param>
        /// <param name="params_">参数</param>
        /// <returns>相应参数</returns>
        public Dictionary<string, Param> Request(string objectName, string action, Dictionary<string, Param> params_)
        {
            ResponseData res = Request(new RequestData(objectName, action, params_));
            if (res.Error.Equals(""))
            {
                return res.OutputParams;
            }
            else
            {
                throw new Exception(res.Error);
            }
        }

        /// <summary>
        /// 发起请求
        /// </summary>
        /// <param name="requestData">请求数据</param>
        /// <returns>响应数据</returns>
        protected abstract ResponseData Request(RequestData requestData);
        public abstract void Dispose();

    }

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
#endif