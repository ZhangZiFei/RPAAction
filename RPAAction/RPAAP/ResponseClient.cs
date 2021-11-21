namespace RPAAP
{
    /// <summary>
    /// RPA Action响应端
    /// </summary>
    public abstract class ResponseClient
    {
        public ResponseClient()
        {
            BeforeCreate();

            RequestData r = Request();
            while (r != null)
            {
                ResponseData res;
                try
                {
                    res = RunAction(r);
                }
                catch (System.Exception e)
                {
                    res = new ResponseData(new System.Collections.Generic.Dictionary<string, Param>(), e.ToString());
                }
                Response(res);
                r = Request();
            }
        }

        /// <summary>
        /// 获取请求数据,如果返回<see cref="null"/>结束RPA对象
        /// </summary>
        /// <returns>请求数据</returns>
        protected abstract RequestData Request();

        /// <summary>
        /// 运行Action
        /// </summary>
        /// <param name="requestData">请求数据</param>
        /// <returns>响应数据</returns>
        protected abstract ResponseData RunAction(RequestData requestData);

        /// <summary>
        /// 响应请求
        /// </summary>
        /// <param name="requestData">响应数据</param>
        protected abstract void Response(ResponseData responseData);

        protected virtual void BeforeCreate() { }
    }
}
