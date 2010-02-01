using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Portal.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class ListVisibleAttribute : Attribute
    {
        public bool Visible
        {
            get;
            set;
        }

        public ListVisibleAttribute()
        {
            Visible = true;
        }

        public ListVisibleAttribute(bool visible)
        {
            Visible = visible;
        }
    }
}
