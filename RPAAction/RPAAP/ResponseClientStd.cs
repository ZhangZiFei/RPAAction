using Newtonsoft.Json;
using System;

namespace RPAAP
{
    /// <summary>
    /// RPA Action 标准输入输出 响应端
    /// </summary>
    public abstract class ResponseClientStd : ResponseClient
    {
        public ResponseClientStd()
            : base()
        {

        }

        protected override RequestData Request()
        {
            string s =Console.ReadLine();

            if (s.Length == 0)
            {
                Console.WriteLine("EXIT");
                return null;
            }
            else
            {
                return JsonConvert.DeserializeObject<RequestData>(s);
            }
        }

        protected override void Response(ResponseData responseData)
        {
            Console.WriteLine(JsonConvert.SerializeObject(responseData));
        }

        protected override void BeforeCreate()
        {
            base.BeforeCreate();

            Console.InputEncoding = Tool.DefEncoding;
            Console.OutputEncoding = Tool.DefEncoding;
        }
    }
}
