using RPAAction.Base;
using System.Collections.Generic;
using System.Data.Common;
using System.IO;
using System.Text;

namespace RPAAction.Data_CSO
{
    public class TXTDataImport : RPADataImport
    {
        /// <param name="path">文件路径</param>
        /// <param name="delimiter">分隔符,如果为空则判断文件后缀,csv为",",其余默认"\t"</param>
        /// <param name="withTitle">是否写入标题</param>
        public TXTDataImport(string path, string delimiter = "", bool withTitle = true)
        {
            WithTitle = withTitle;
            Delimiter = delimiter;
            Path = System.IO.Path.GetFullPath(path);

            if (string.IsNullOrEmpty(Delimiter))
            {
                string ext = System.IO.Path.GetExtension(Path);
                if (ext.ToLower().Equals(".csv"))
                {
                    Delimiter = ",";
                }
                else
                {
                    Delimiter = "\t";
                }
            }
        }

        protected override void CreateTable(DbDataReader r)
        {
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(Path));
            writer = new StreamWriter(Path, false, new UTF8Encoding(false));

            int FieldCount = r.FieldCount;
            for (int i = 0; i < FieldCount; i++)
            {
                if (Fields.ContainsKey(r.GetName(i)))
                {
                    throw new ActionException($"出现相同的标题\"{r.GetName(i)}\"");
                }
                else
                {
                    Fields.Add(r.GetName(i), i);
                }
            }
            //标题
            if (WithTitle)
            {
                foreach (var item in Fields)
                {
                    SetValue(item.Key, item.Key);
                }
                UpdataRow();
            }
        }

        protected override void SetValue(string field, object value)
        {
            string s = value == null ? "" : value.ToString();
            if (s.IndexOf('\n') > 1 || s.IndexOf('\r') > 1 || s.IndexOf(Delimiter) > 1)
            {
                if (s.IndexOf('"') > 1)
                {
                    s = s.Replace("\"", "\"\"");
                }
                s = "\"" + s + "\"";
            }
            if (writeDelimiter)
            {
                s = Delimiter + s;
            }
            else
            {
                writeDelimiter = true;
            }
            if (writeLine)
            {
                writeLine = false;
                writer.WriteLine();
            }
            writer.Write(s);
        }

        protected override void UpdataRow()
        {
            writeDelimiter = false;
            writeLine = true;
        }
        protected override void Close()
        {
            writer.Flush();
            writer.Dispose();
        }

        private readonly bool WithTitle;
        /// <summary>
        /// 分隔符
        /// </summary>
        private readonly string Delimiter;
        private readonly string Path;
        private StreamWriter writer;
        /// <summary>
        /// 写入分隔符
        /// </summary>
        private bool writeDelimiter = false;
        private bool writeLine = false;
        private readonly Dictionary<string, int> Fields = new Dictionary<string, int>();
    }
}
