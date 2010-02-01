using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class DisplayNameAttribute : Attribute
    {
        public string Name
        {
            get;
            set;
        }

        public DisplayNameAttribute(string name)
        {
            Name = name;
        }
    }
}
