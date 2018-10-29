using Syncfusion.Drawing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Text;

namespace Pusintek.AspNetcore.DocIO
{
    public class DataFieldGroup
    {
        public DataFieldGroup()
        {
            Options = new Dictionary<string, Options>();
        }
        public string Key { get; set; }
        public Dictionary<string, Options> Options { get; set; }
        public IEnumerable Data { get; set; }

    }

    public class DataField 
    {
        public Dictionary<string, string> Data{ get; set; }
        public Dictionary<string, Options> Options { get; set; }       
    }

    public class Options
    {
        public Boolean FromUrl { get; set; }     
        public int Height { get; set; }
        public int Width { get; set; }
        public int PercentageResize { get; set; }
    }
}
