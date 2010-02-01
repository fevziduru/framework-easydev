using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace EasyDev.EPS.Attributes
{
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
    public class RegularExpressionValidatorAttribute : ValidatorAttribute, IValidator
    {
        public string RegularExpression { get; set; }

        public RegularExpressionValidatorAttribute(string msg)
            : base(msg)
        { }

        /// <summary>
        /// 正则表达式验证
        /// </summary>
        /// <param name="msg">错误信息</param>
        /// <param name="_regExp">正则表达式</param>
        public RegularExpressionValidatorAttribute(string msg, string _regExp)
            : base(msg)
        {
            RegularExpression = _regExp;
        }

        #region IValidation 成员

        /// <summary>
        /// 验证方法
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public string DoValidate(object value)
        {
            if (Regex.IsMatch(value.ToString(), RegularExpression, RegexOptions.Singleline))
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
