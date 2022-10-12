using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dev.Framework.Model
{
    internal class PlugInLoadConfig
    {
        public string Name { get; set; }

        public string Text { get; set; }

        public string AssemblyName { get; set; }

        public string ClassName { get; set; }

        public bool IsHidden { get; set; }

    }
}
