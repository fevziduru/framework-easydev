using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
    public class KeyPropertyAttribute : Attribute
    {
        public KeyPropertyAttribute() { }
    }
}
