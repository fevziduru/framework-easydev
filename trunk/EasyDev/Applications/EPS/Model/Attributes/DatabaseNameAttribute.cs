using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Class)]
    public class DatabaseNameAttribute : Attribute
    {
        public string DatabaseName { get; set; }

        public DatabaseNameAttribute(string name)
        {
            this.DatabaseName = name;
        }
    }
}
