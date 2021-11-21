using System;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace RPAAP
{
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
}