using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class EmailValidatorAttribute : ValidatorAttribute, IValidator
    {
        public EmailValidatorAttribute() { }

        public EmailValidatorAttribute(string msg)
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
            if (Regex.IsMatch(value.ToString(), @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*", RegexOptions.Singleline)) 
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
