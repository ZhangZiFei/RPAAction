using System;
using System.Collections.Generic;

namespace RPAAP
{
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
}
