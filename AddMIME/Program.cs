using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddMIME
{
    class Program
    {
        static void Main(string[] args)
        {
            Begin();
        }

        public static void AddMIMEType()
        {
            DirectoryEntry rootEntry = GetDirectoryEntry("IIS://localhost/w3svc/1/root");
            foreach (PropertyValueCollection pc in rootEntry.Properties)
            {
                Console.WriteLine(pc.PropertyName + ":" + pc.Value);
            }
            //rootEntry.Properties["MimeMap"].Add(
            IISOle.MimeMapClass _NewMime = new IISOle.MimeMapClass();          //新建MIME类型
            _NewMime.Extension = ".xap";
            _NewMime.MimeType = ".xap";
            rootEntry.Properties["MimeMap"].Add(_NewMime);  //添加MIME类型

            rootEntry.CommitChanges();//更改目录

        }
    }

}
