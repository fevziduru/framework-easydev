using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Data;
using System.Collections;
using EasyDev.Util;
using System.Web;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS
{
    public abstract class AbstractModel : IModel
    {
        public AbstractModel()
        {
            Initialize();
        }

        //public virtual object this[string propName]
        //{
        //    get 
        //    {                
        //        return GetProperties()[propName].GetValue(this, null);
        //    }
        //    set
        //    {
        //        GetProperties()[propName].SetValue(this, value, null);
        //    }
        //}

        public virtual object GetValue(string propName)
        {
            return GetProperties()[propName].GetValue(this, null);
        }

        public virtual void SetValue(string propName, object value)
        {
            GetProperties()[propName].SetValue(this, value, null);
        }

        public virtual IDictionary<string, PropertyInfo> GetProperties()
        {
            IDictionary<string, PropertyInfo> properties =
                HttpRuntime.Cache[this.GetType().Name + "_Properties"] as IDictionary<string, PropertyInfo>;

            if (properties == null)
            {
                properties = TypeUtil.GetProperties(this.GetType());
                HttpRuntime.Cache.Insert(this.GetType().Name + "_Properties", properties);
            }
            return properties;
        }
        
        protected virtual void Initialize()
        {
            ParameterSymbol = ":";

            object[] attributes = this.GetType().GetCustomAttributes(true);
            for (int i = 0; i < attributes.Length; i++)
            {
                if (typeof(DatabaseNameAttribute).Equals(attributes[i].GetType()))
                {
                    DatabaseNameAttribute attr = attributes[i] as DatabaseNameAttribute;

                    if (attr != null)
                    {
                        ParameterSymbol =
                            attr.DatabaseName.Equals("Oracle",
                            StringComparison.CurrentCultureIgnoreCase) ? ":" : "@";

                        break;
                    }
                }
            }
        }

        protected virtual string ParameterSymbol
        {
            get;
            set;
        }

        /// <summary>
        /// 复制模型对象，参数中的对象为源对象，将参数对象中的数据复制到当前对象中
        /// </summary>
        /// <typeparam name="T">复制类型</typeparam>
        /// <param name="sourceModel">复制源</param>
        public virtual void Copy<T>(T sourceModel) where T : IModel
        {
            try
            {
                PropertyInfo[] destType = this.GetType().GetProperties();
                PropertyInfo[] properties = sourceModel.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    destType[i].SetValue(this, properties[i].GetValue(sourceModel, null), null);
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 复制模型对象，参数中的对象为源对象，将参数对象中的数据复制到当前对象中
        /// </summary>
        /// <param name="srcModel">复制源</param>
        public virtual void Copy(IModel srcModel)
        {
            if (this.GetType().Equals(srcModel.GetType()))
            {
                PropertyInfo[] destType = this.GetType().GetProperties();
                PropertyInfo[] properties = srcModel.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    destType[i].SetValue(this, properties[i].GetValue(srcModel, null), null);
                }
            }
            else
            {
                throw new ArgumentException("type_is_not_fit");
            }
        }

        /// <summary>
        /// 模型比较
        /// </summary>
        /// <typeparam name="TDest"></typeparam>
        /// <param name="obj"></param>
        /// <returns></returns>
        public virtual bool isModelEqualTo<TDest>(object obj) where TDest : IModel, new()
        {
            if (this.GetType().Equals(typeof(TDest)) == false)
            {
                return false;
            }

            PropertyInfo[] srcProperties = this.GetType().GetProperties();
            PropertyInfo[] destProperties = typeof(TDest).GetProperties();

            if (srcProperties.Length != srcProperties.Length)
            {
                return false;
            }

            bool isEqual = true;

            for (int i = 0; i < destProperties.Length; i++)
            {
                isEqual = destProperties[i].GetValue(obj,null)
                    .Equals(srcProperties[i].GetValue(this,null));
                if (isEqual == false)
                {
                    break;
                }
            }

            return isEqual;
        }

        /// <summary>
        /// 构建当前模型的SELECT命令，此命令是参数化命令
        /// </summary>
        /// <param name="condition">条件表达式,此条件表达式指的是SQL命令中的WHERE</param>
        /// <returns>参数化的SQL命令</returns>
        public virtual string SelectCommand(string condition)
        {
            StringBuilder select = new StringBuilder();

            try
            {
                string selectCols = string.Empty;
                PropertyInfo[] properties = this.GetType().GetProperties();
                
                for (int i = 0; i < properties.Length; i++)
                {
                    selectCols += properties[i].Name;
                    if (i + 1 < properties.Length)
                    {
                        selectCols += ", ";
                    }
                }

                if (selectCols.Length == 0)
                {
                    throw new CommandExpressionException("there_is_not_any_column_selected");
                }

                select.AppendFormat("SELECT {1} FROM {0} ", this.GetType().Name, selectCols);

                if (condition != null && condition.Length > 0)
                {
                    select.Append(" WHERE " + condition.Replace("#", ParameterSymbol));
                }
            }
            catch (CommandExpressionException e)
            {
                throw e;
            }

            return select.ToString();
        }

        /// <summary>
        /// 构建当前模型的SELECT命令，此命令是参数化命令
        /// </summary>
        /// <returns>参数化的SQL命令</returns>
        public virtual string SelectCommand()
        {
            return SelectCommand(null);
        }

        /// <summary>
        /// 构建当前模型的INSERT命令，此命令是参数化命令
        /// </summary>
        /// <returns>参数化的SQL命令</returns>
        public virtual string InsertCommand()
        {
            StringBuilder insert = new StringBuilder();
            try
            {
                string cols = string.Empty;
                string values = string.Empty;
                PropertyInfo[] properties = this.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmpValue = properties[i].GetValue(this, null);
                    
                    if (tmpValue != null)
                    {
                        if (tmpValue.GetType().Equals(typeof(DateTime)))
                        {
                            if (Convert.ToDateTime(tmpValue) == DateTime.MinValue)
                            { 
                                //值不为空，且为日期类型最小值(不做处理)
                            }
                            else
                            {
                                //值不为空，且不为日期类型最小值
                                cols += properties[i].Name;
                                values += ParameterSymbol + properties[i].Name;

                                if (i + 1 < properties.Length)
                                {
                                    cols += ", ";
                                    values += ", ";
                                }
                            }
                        }
                        else
                        {
                            //值不为空且不是日期数据
                            cols += properties[i].Name;
                            values += ParameterSymbol + properties[i].Name;

                            if (i + 1 < properties.Length)
                            {
                                cols += ", ";
                                values += ", ";
                            }
                        }
                    }
                }

                insert.AppendFormat("INSERT INTO {0} ({1}) VALUES ({2})", 
                    this.GetType().Name, cols, values);
            }
            catch (CommandExpressionException e)
            {
                throw e;
            }

            return insert.ToString();
        }

        /// <summary>
        /// 构建当前模型的DELETE命令，此命令是参数化命令
        /// </summary>
        /// <param name="condition">条件表达式</param>
        /// <returns></returns>
        public virtual string DeleteCommand(string condition)
        {
            StringBuilder delete = new StringBuilder();

            delete.AppendFormat("DELETE FROM {0} ", this.GetType().Name);

            if (condition != null && condition.Length > 0)
            {
                delete.Append(" WHERE " + condition.Replace("#", ParameterSymbol));
            }

            return delete.ToString();
        }

        /// <summary>
        /// 构建当前模型的DELETE命令，此命令是参数化命令
        /// </summary>
        /// <returns></returns>
        public virtual string DeleteCommand()
        {
            return DeleteCommand(null);
        }

        /// <summary>
        /// 构建当前模型的UPDATE命令，此命令是参数化命令
        /// </summary>
        /// <param name="condition">条件表达式,如果参数为NULL则表示此SQL没有WHERE子句</param>
        /// <returns></returns>
        public virtual string UpdateCommand(string condition)
        {
            StringBuilder update = new StringBuilder();

            try
            {
                string pairs = string.Empty;
                PropertyInfo[] properties = this.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmpValue = properties[i].GetValue(this, null);
                    object[] attributes = properties[i].GetCustomAttributes(false);

                    int keyfield = attributes.Count<object>(p => p.GetType().Equals(typeof(KeyPropertyAttribute)));

                    if (keyfield > 0)
                    {
                        continue;
                    }
                    else
                    {
                        if (tmpValue != null)
                        {
                            if (tmpValue.GetType().Equals(typeof(DateTime)))
                            {
                                if (Convert.ToDateTime(tmpValue) == DateTime.MinValue)
                                {
                                    //值不为空，且为日期类型最小值(不做处理)
                                }
                                else
                                {
                                    //值不为空，且不为日期类型最小值
                                    pairs += string.Format("{0}={1}{0}", properties[i].Name.ToLower(), ParameterSymbol);

                                    if (i + 1 < properties.Length)
                                    {
                                        pairs += ", ";
                                    }
                                }
                            }
                            else
                            {
                                //值不为空且不是日期数据
                                pairs += string.Format("{0}={1}{0}", properties[i].Name.ToLower(), ParameterSymbol);

                                if (i + 1 < properties.Length)
                                {
                                    pairs += ", ";
                                }
                            }
                        }
                    }
                }

                update.AppendFormat("UPDATE {0} SET {1}", this.GetType().Name, pairs.Trim().TrimEnd(','));

                if (condition != null && condition.Length > 0)
                {
                    update.Append(" WHERE " + condition.Replace("#", ParameterSymbol));
                }
            }
            catch (CommandExpressionException e)
            {
                throw e;
            }

            return update.ToString();
        }

        /// <summary>
        /// 构建当前模型的UPDATE命令，此命令是参数化命令
        /// </summary>
        /// <returns></returns>
        public virtual string UpdateCommand()
        {
            return UpdateCommand(null);
        }

        /// <summary>
        /// 将模型转换为对象数组,只转换模型中值不为空的公共属性，对于日期对象只转换值不为DateTime.MinValue的公共属性
        /// </summary>
        /// <returns></returns>
        public virtual object[] ToParamObjects()
        {
            ArrayList tmp = new ArrayList();

            try
            {
                PropertyInfo[] properties = this.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmpValue = properties[i].GetValue(this, null);
                    if (tmpValue != null)
                    {
                        if (tmpValue.GetType().Equals(typeof(DateTime)))
                        {
                            if (Convert.ToDateTime(tmpValue) == DateTime.MinValue)
                            {
                                //值不为空，且为日期类型最小值(不做处理)
                            }
                            else
                            {
                                //值不为空，且不为日期类型最小值
                                tmp.Add(tmpValue);
                            }
                        }
                        else
                        {
                            //值不为空且不是日期数据
                            tmp.Add(tmpValue);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return tmp.ToArray();
        }

        public virtual IDictionary<string, object> ToParamDictionary()
        {
            IDictionary<string, object> items = new Dictionary<string, object>();

            try
            {
                PropertyInfo[] properties = this.GetType().GetProperties();

                for (int i = 0; i < properties.Length; i++)
                {
                    object tmpValue = properties[i].GetValue(this, null);
                    if (tmpValue != null)
                    {
                        if (tmpValue.GetType().Equals(typeof(DateTime)))
                        {
                            if (Convert.ToDateTime(tmpValue) == DateTime.MinValue)
                            {
                                //值不为空，且为日期类型最小值(不做处理)
                            }
                            else
                            {
                                //值不为空，且不为日期类型最小值
                                items.Add(properties[i].Name.ToLower(),tmpValue);
                            }
                        }
                        else
                        {
                            //值不为空且不是日期数据
                            items.Add(properties[i].Name.ToLower(), tmpValue);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }

            return items;
        }

        /// <summary>
        /// 进行模型数据验证
        /// </summary>
        /// <returns></returns>
        public virtual IDictionary<string, string> ProcessValidation()
        {
            PropertyInfo[] properties = this.GetType().GetProperties();
            IDictionary<string, string> errors = new Dictionary<string, string>();

            for (int i = 0; i < properties.Length; i++)
            {
                //取得模型对象每个公共属性的自定义属性标识
                object[] attributes = properties[i].GetCustomAttributes(false);

                //当前公共属性的值
                object currentValue = properties[i].GetValue(this,null);

                for (int j = 0; j < attributes.Length; j++)
                {
                    IValidator validationAttr = attributes[j] as IValidator;
                    if (validationAttr != null)
                    { 
                        string tmpMsg = validationAttr.DoValidate(currentValue);
                        if (tmpMsg.Length > 0)
                        {
                            errors.Add(properties[i].Name, tmpMsg);
                        }
                    }
                }
            }

            return errors;
        }

        public virtual object GetKeyPropertyValue()
        {
            IDictionary<string, PropertyInfo> properties = TypeUtil.GetProperties(this.GetType());
            object result = null;

            foreach (PropertyInfo item in properties.Values)
            {
                result = TypeUtil.IsPropertyContainsAttribute(this.GetType(), item, typeof(KeyPropertyAttribute));

                if (result != null)
                {
                    break;
                }
            }

            return result;
        }
    }
}
