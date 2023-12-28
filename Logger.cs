using System;
using System.Collections.Generic;
using System.Text;

using System.IO;

namespace DrawDocument
{
    class Logger
    {
        public class TextFile
        {
            string path { get; set; }

            public TextFile(string path)
            {
                this.path = path;

                FileInfo fi = new FileInfo(path);
                var di = new DirectoryInfo(fi.DirectoryName);

                //если указанной в качестве аргумента директории не существует - создадим
                if (!di.Exists)
                    di.Create();

                //если файла не существует - создаём
                if (!fi.Exists)
                    File.Create(path).Close();
            }

            public void Add(string text)
            {
                StringBuilder sb = new StringBuilder();
                sb.Append(String.Format("{0} - ", DateTime.Now));
                sb.Append(text);
                sb.Append(Environment.NewLine);
                File.AppendAllText(path, sb.ToString(), Encoding.UTF8);
                sb.Clear();
            }
        }
    }
}
