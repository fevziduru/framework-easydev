using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS
{
    [AttributeUsage(AttributeTargets.Method)]
    public class EnsureLoginAttribute : Attribute
    {
        public EnsureLoginAttribute() { }
    }
}
