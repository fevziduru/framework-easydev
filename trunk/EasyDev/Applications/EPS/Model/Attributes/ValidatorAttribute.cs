using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public abstract class ValidatorAttribute : Attribute
    {
        public virtual string ErrorMessage { get; set; }

        public virtual string ErrorMessageFormat { get; set; }

        public ValidatorAttribute() 
        {
            ErrorMessage = "Validation Error";
            ErrorMessageFormat = string.Empty;
        }

        public ValidatorAttribute(string msg)
        {
            ErrorMessage = msg;
        }
    }
}
