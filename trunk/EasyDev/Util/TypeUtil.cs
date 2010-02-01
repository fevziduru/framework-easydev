using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace EasyDev.Util
{
        /// <summary>
        /// 
        /// </summary>
        public class TypeUtil
        {
                /// <summary>
                /// 
                /// </summary>
                /// <param name="t"></param>
                /// <returns></returns>
                public static IDictionary<string, PropertyInfo> GetProperties(Type t)
                {
                        return t.Properties();
                }

                /// <summary>
                /// 
                /// </summary>
                /// <typeparam name="T"></typeparam>
                /// <returns></returns>
                public static IDictionary<string, PropertyInfo> GetProperties<T>()
                {
                        return typeof(T).Properties();
                }

                /// <summary>
                /// 
                /// </summary>
                /// <param name="tObject"></param>
                /// <param name="propName"></param>
                /// <returns></returns>
                public static object GetValueFromTypeInstance(Type tObject, string propName)
                {
                        return tObject.ValueOfProperty(propName);
                }

                /// <summary>
                /// 
                /// </summary>
                /// <typeparam name="TObject"></typeparam>
                /// <param name="propName"></param>
                /// <returns></returns>
                public static object GetValueFromTypeInstance<TObject>(string propName)
                {
                        return typeof(TObject).ValueOfProperty(propName);
                }

                /// <summary>
                /// 判断类型t中是否有attr这个自定义属性
                /// </summary>
                /// <param name="t"></param>
                /// <param name="attr"></param>
                /// <returns></returns>
                public static bool IsTypeContainsAttribute(Type t, Type attr)
                {
                        return t.HasAttribute(attr);
                }

                /// <summary>
                /// 判断类型TObject中是否有TAttribute这个自定义属性
                /// </summary>
                /// <typeparam name="TObject"></typeparam>
                /// <typeparam name="TAttribute"></typeparam>
                /// <returns></returns>
                public static bool IsTypeContainsAttribute<TObject, TAttribute>()
                    where TAttribute : Attribute
                {
                        return typeof(TObject).HasAttribute(typeof(TAttribute));
                }

                /// <summary>
                /// 判断类型tobject中的名为propName的属性是否有类型为tattr的属性
                /// 返回空则不存在
                /// </summary>
                /// <param name="tobject"></param>
                /// <param name="propName"></param>
                /// <param name="tattr"></param>
                /// <returns></returns>
                public static object IsPropertyContainsAttribute(Type tobject, string propName, Type tattr)
                {
                        return tobject.HasAttribute(propName, tattr);
                }

                public static object IsPropertyContainsAttribute(Type type, PropertyInfo property, Type tattr)
                {
                        return property.HasAttribute(type, tattr);
                }

                /// <summary>
                /// 判断类型TObject中的名为propName的属性是否有类型为TAttribute的属性
                /// 返回空则不存在
                /// </summary>
                /// <typeparam name="TObject"></typeparam>
                /// <typeparam name="TAttribute"></typeparam>
                /// <param name="propName"></param>
                /// <returns></returns>
                public static object IsPropertyContainsAttribute<TObject, TAttribute>(string propName)
                    where TAttribute : Attribute
                {
                        return typeof(TObject).HasAttribute(propName, typeof(TAttribute));
                }

                /// <summary>
                /// 从类型tObject的名为propName的成员中找到类型为tAttr的自定义属性，并取得这个自定义属性中名为fieldName的公共成员的值
                /// </summary>
                /// <param name="tObject"></param>
                /// <param name="propName"></param>
                /// <param name="tAttr"></param>
                /// <param name="fieldName"></param>
                /// <returns></returns>
                public static object GetValueFromAttributeOfProperty(Type tObject, string propName, Type tAttr, string fieldName)
                {
                        return tObject.ValueOfAttribute(propName, tAttr, fieldName);
                }

                /// <summary>
                /// 从类型TObject的名为propName的成员中找到类型为TAttribute的自定义属性，并取得这个自定义属性中名为fieldName的公共成员的值
                /// </summary>
                /// <typeparam name="TObject"></typeparam>
                /// <typeparam name="TAttribute"></typeparam>
                /// <param name="propName"></param>
                /// <param name="fieldName"></param>
                /// <returns></returns>
                public static object GetValueFromAttributeOfProperty<TObject, TAttribute>(string propName, string fieldName)
                    where TAttribute : Attribute
                {
                        return typeof(TObject).ValueOfAttribute(propName, typeof(TAttribute), fieldName);
                }
        }
}
