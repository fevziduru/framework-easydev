using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class RequiredValidatorAttribute : ValidatorAttribute, IValidator
    {
        public RequiredValidatorAttribute() { }

        public RequiredValidatorAttribute(string msg)
            : base(msg)
        { }

        #region IValidation 成员

        public string DoValidate(object value)
        {
            return (value == null || value.ToString().Length == 0) 
                ? ErrorMessage : string.Empty;
        }

        #endregion
    }
}
