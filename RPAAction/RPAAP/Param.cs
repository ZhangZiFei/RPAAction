using Newtonsoft.Json;
using System.Data;

namespace RPAAP
{
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
}
