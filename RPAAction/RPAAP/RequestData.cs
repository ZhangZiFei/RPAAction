using System;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace RPAAP
{
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
}
