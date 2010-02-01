using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class UrlValidatorAttribute : ValidatorAttribute, IValidator
    {
        public UrlValidatorAttribute() { }

        public UrlValidatorAttribute(string msg)
            : base(msg)
        { }

        #region IValidation 成员

        /// <summary>
        /// 验证方法
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string DoValidate(object value)
        {
            if (Regex.IsMatch(value.ToString(), @"http://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)? ", RegexOptions.Singleline)) 
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
