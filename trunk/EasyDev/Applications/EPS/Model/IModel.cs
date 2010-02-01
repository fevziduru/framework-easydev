using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace EasyDev.EPS
{
    public interface IModel
    {
        //object this[string propName] { get; set; }

        /// <summary>
        /// 取得模型对象的公共属性列表
        /// </summary>
        /// <returns></returns>
        IDictionary<string, PropertyInfo> GetProperties();

        /// <summary>
        /// 取得模型对象属性的值
        /// </summary>
        /// <param name="propName"></param>
        /// <returns></returns>
        object GetValue(string propName);

        /// <summary>
        /// 为模型对象的属性赋值
        /// </summary>
        /// <param name="propName"></param>
        /// <param name="value"></param>
        void SetValue(string propName, object value);

        /// <summary>
        /// 模型复制
        /// </summary>
        /// <param name="srcModel"></param>
        void Copy(IModel srcModel);

        /// <summary>
        /// 构造模型相关的SQL SELECT命令
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        string SelectCommand(string condition);

        /// <summary>
        /// 构造模型相关的SQL DELETE命令
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        string DeleteCommand(string condition);

        /// <summary>
        /// 构造模型相关的SQL UPDATE命令
        /// </summary>
        /// <param name="condition"></param>
        /// <returns></returns>
        string UpdateCommand(string condition);

        /// <summary>
        /// 构造模型相关的SQL INSERT命令
        /// </summary>
        /// <returns></returns>
        string InsertCommand();

        /// <summary>
        /// 构造模型相关的SQL SELECT命令
        /// </summary>
        /// <returns></returns>
        string SelectCommand();

        /// <summary>
        /// 构造模型相关的SQL DELETE命令
        /// </summary>
        /// <returns></returns>
        string DeleteCommand();

        /// <summary>
        /// 构造模型相关的SQL UPDATE命令
        /// </summary>
        /// <returns></returns>
        string UpdateCommand();

        /// <summary>
        /// 转换成对象数组
        /// </summary>
        /// <returns></returns>
        object[] ToParamObjects();

        /// <summary>
        /// 将模型转换成目录对象
        /// </summary>
        /// <returns></returns>
        IDictionary<string, object> ToParamDictionary();

        /// <summary>
        /// 取得模型对象键属性的值
        /// </summary>
        /// <returns></returns>
        object GetKeyPropertyValue();
    }
}
