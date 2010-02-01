using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.Portal.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class NumberRangeValidatorAttribute : ValidatorAttribute , IValidator
    {
        private double minNumber;
        private double maxNumber;

        public NumberRangeValidatorAttribute(string msg)
            : base(msg)
        { }

        public NumberRangeValidatorAttribute(string format, double min, double max)
        {
            ErrorMessage = string.Format(format, min, max);
        }

        #region IValidator 成员

        public string DoValidate(object value)
        {
            double tmp = Convert.ToDouble(value);
            if (tmp >= minNumber && tmp < maxNumber)
            {
                return string.Empty;
            }
            else
            {
                return ErrorMessage;
            }
        }

        #endregion
    }
}
