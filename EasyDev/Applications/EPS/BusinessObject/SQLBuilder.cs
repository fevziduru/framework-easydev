using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using EasyDev.EPS.Portal;
using EasyDev.EPS.Attributes;

namespace EasyDev.EPS.BusinessObject
{
        public sealed class DBName
        {
                public static readonly string ORACLE = "ORACLE";

                public static readonly string MSSQL = "MSSQL";
        }

        /// <summary>
        /// 根据MODEL对象生成SQL
        /// 出于简单处理考虑，这里处理的主属性列类型要是字符串类型。
        /// 因为字符串类型作为主属性是最为常见的做法，所以这里对主属性的处理按字符串进行
        /// 所以在设计数据库的时候要考虑这一点
        /// </summary>
        public class SQLBuilder
        {
                //private static readonly string ORACLE = "Oracle";

                //private static readonly string MSSQLSERVER = "MSSQL";

                #region SELECT SQL

                public static string BuildSelectAll(IModel model)
                {
                        return BuildSelect(model, null, false);
                }

                public static string BuildSelectInstance(IModel model)
                {
                        return BuildSelect(model, null, true);
                }

                /// <summary>
                /// 创建SELECT命令
                /// </summary>
                /// <param name="model"></param>
                /// <param name="schemaName"></param>
                /// <param name="iscondition"></param>
                /// <returns></returns>
                public static string BuildSelect(IModel model, string schemaName, bool iscondition)
                {
                        StringBuilder select = new StringBuilder();
                        select.Append("SELECT ");

                        PropertyInfo[] properties = model.GetType().GetProperties();
                        for (int i = 0; i < properties.Length; i++)
                        {
                                select.Append(properties[i].Name);
                                if (i + 1 < properties.Length)
                                {
                                        select.Append(", ");
                                }
                        }

                        if (schemaName != null && schemaName.Length > 0)
                        {
                                select.AppendFormat(" FROM {0}", schemaName + "." + model.GetType().Name);
                        }
                        else
                        {
                                select.AppendFormat(" FROM {0}", model.GetType().Name);
                        }

                        if (iscondition)
                        {
                                select.Append(" WHERE " + BuildSQLConditionByModel(model));
                        }

                        return select.ToString();
                }
                #endregion

                #region INSERT SQL
                //public static string BuildInsertByModel<TModel>(TModel model, string schemaName) where TModel : IModel, new()
                //{
                //    return BuildInsertByModel(model, schemaName);
                //}

                //public static string BuildInsertByModel<TModel>(TModel model) where TModel : IModel, new()
                //{
                //    return BuildInsertByModel<TModel>(model, null);
                //}

                public static string BuildInsertByModel(IModel model)
                {
                        return BuildInsertByModel(model, null);
                }

                public static string BuildInsertByModel(IModel model, string schemaName)
                {
                        StringBuilder insert = new StringBuilder();
                        string cols = string.Empty;
                        string values = string.Empty;
                        string insertFormat = "INSERT INTO {0} ({1}) VALUES({2})";
                        PropertyInfo[] properties = model.GetType().GetProperties();
                        object[] classAttributes = model.GetType().GetCustomAttributes(false);
                        string dbName = DBName.ORACLE;

                        if (classAttributes.Length > 0)
                        {
                                PropertyInfo[] modelProperties = classAttributes[0].GetType().GetProperties();
                                if (modelProperties.Length > 0)
                                {
                                        dbName = modelProperties[0].GetValue(classAttributes[0], null).ToString();
                                }
                        }

                        #region 对不同类型的转换处理
                        for (int i = 0; i < properties.Length; i++)
                        {
                                if (properties[i].PropertyType == typeof(int) ||
                                    properties[i].PropertyType == typeof(decimal) ||
                                    properties[i].PropertyType == typeof(double) ||
                                    properties[i].PropertyType == typeof(float))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                cols += properties[i].Name;
                                                values += properties[i].GetValue(model, null).ToString();
                                                if (i + 1 < properties.Length)
                                                {
                                                        cols += ",";
                                                        values += ",";
                                                }
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(string) ||
                                    properties[i].PropertyType == typeof(Guid) ||
                                    properties[i].PropertyType == typeof(bool))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                cols += properties[i].Name;
                                                values += "'" + properties[i].GetValue(model, null).ToString() + "'";
                                                if (i + 1 < properties.Length)
                                                {
                                                        cols += ",";
                                                        values += ",";
                                                }
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(DateTime))
                                {
                                        if (Convert.ToDateTime(properties[i].GetValue(model, null)) != DateTime.MinValue)
                                        {
                                                cols += properties[i].Name;
                                                if (DBName.ORACLE.Equals(dbName, StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                        values += "TO_DATE('" + properties[i].GetValue(model, null).ToString() + "','YYYY-MM-DD hh24:mi:ss')";
                                                }
                                                else
                                                {
                                                        values += "'" + properties[i].GetValue(model, null).ToString() + "'";
                                                }

                                                if (i + 1 < properties.Length)
                                                {
                                                        cols += ",";
                                                        values += ",";
                                                }
                                        }
                                }
                        }
                        #endregion

                        if (cols.Length > 0)
                        {
                                cols = cols.TrimEnd(new char[] { ',' });
                        }

                        if (values.Length > 0)
                        {
                                values = values.TrimEnd(new char[] { ',' });
                        }

                        if (schemaName != null && schemaName.Length > 0)
                        {
                                insert.AppendFormat(insertFormat,
                                    schemaName + "." + model.GetType().Name, cols, values);
                        }
                        else
                        {
                                insert.AppendFormat(insertFormat,
                                    model.GetType().Name, cols, values);
                        }

                        return insert.ToString();
                }
                #endregion

                #region UPDATE SQL

                public static string BuildUpdateByModel(IModel model)
                {
                        return BuildUpdateByModel(model, null);
                }

                public static string BuildUpdateByModel(IModel model, string schemaName)
                {
                        StringBuilder update = new StringBuilder();
                        string whereClause = string.Empty;

                        string dbName = GetDatabaseType(model);

                        if (schemaName != null && schemaName.Length > 0)
                        {
                                update.AppendFormat("UPDATE {0} SET ", schemaName + "." + model.GetType().Name);
                        }
                        else
                        {
                                update.AppendFormat("UPDATE {0} SET ", model.GetType().Name);
                        }

                        PropertyInfo[] properties = model.GetType().GetProperties();

                        for (int i = 0; i < properties.Length; i++)
                        {
                                object[] attributes = properties[i].GetCustomAttributes(false);

                                //if (attributes.Length > 0 && (whereClause.Length > 0 && properties[i].PropertyType != typeof(DateTime)
                                //    && properties[i].GetValue(model, null) != null) ||
                                //    (whereClause.Length > 0 && (properties[i].PropertyType == typeof(DateTime) &&
                                //        (Convert.ToDateTime(properties[i].GetValue(model, null))) != DateTime.MinValue)))
                                //{
                                //    whereClause += " AND ";
                                //}

                                #region 对不同类型的转换处理
                                if (properties[i].PropertyType == typeof(int) ||
                                            properties[i].PropertyType == typeof(decimal) ||
                                            properties[i].PropertyType == typeof(double) ||
                                            properties[i].PropertyType == typeof(float))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                update.AppendFormat("{0}={1}",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());

                                                if (i + 1 < properties.Length)
                                                {
                                                        update.Append(",");
                                                }
                                        }

                                        if (attributes.Length > 0 && attributes[0].GetType() == typeof(KeyPropertyAttribute))
                                        {
                                                whereClause += string.Format(" {0}={1} AND",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(string) ||
                                            properties[i].PropertyType == typeof(Guid) ||
                                            properties[i].PropertyType == typeof(bool))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                update.AppendFormat("{0}='{1}'",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());

                                                if (i + 1 < properties.Length)
                                                {
                                                        update.Append(",");
                                                }
                                        }

                                        if (attributes.Length > 0 && attributes[0].GetType() == typeof(KeyPropertyAttribute))
                                        {
                                                whereClause += string.Format(" {0}='{1}' AND",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(DateTime))
                                {
                                        if (Convert.ToDateTime(properties[i].GetValue(model, null)) != DateTime.MinValue)
                                        {
                                                if (DBName.ORACLE.Equals(dbName, StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                        update.AppendFormat("{0}=TO_DATE('{1}','YYYY-MM-DD hh24:mi:ss')",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }
                                                else
                                                {
                                                        update.AppendFormat("{0}='{1}'",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }

                                                if (i + 1 < properties.Length)
                                                {
                                                        update.Append(",");
                                                }
                                        }

                                        if (attributes.Length > 0 && attributes[0].GetType() == typeof(KeyPropertyAttribute))
                                        {
                                                if (DBName.ORACLE.Equals(dbName, StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                        whereClause += string.Format(" {0}=TO_DATE('{1}','YYYY-MM-DD hh24:mi:ss') AND",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }
                                                else
                                                {
                                                        whereClause += string.Format(" {0}='{1}' AND",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }
                                        }
                                }
                        }

                        if (update.ToString().Length > 0)
                        {
                                update.Remove(update.Length - 1, 1);
                        }
                                #endregion

                        if (whereClause.Length > 0)
                        {
                                update.Append(" WHERE " + whereClause.Remove(whereClause.Length - "AND".Length));
                        }

                        return update.ToString();
                }

                private static string GetDatabaseType(IModel model)
                {
                        object[] classAttributes = model.GetType().GetCustomAttributes(false);
                        string dbName = DBName.ORACLE;

                        if (classAttributes.Length > 0)
                        {
                                PropertyInfo[] modelProperties = classAttributes[0].GetType().GetProperties();
                                if (modelProperties.Length > 0)
                                {
                                        dbName = modelProperties[0].GetValue(classAttributes[0], null).ToString();
                                }
                        }
                        return dbName;
                }
                #endregion

                #region DELETE SQL

                public static string BuildDeleteByModel(IModel model)
                {
                        return BuildDeleteByModel(model, null);
                }

                public static string BuildDeleteByModel(IModel model, string schemaName)
                {
                        StringBuilder delete = new StringBuilder();
                        try
                        {
                                string deleteWhere = string.Empty;
                                string termFormat = "{0}='{1}'";
                                PropertyInfo[] properties = model.GetType().GetProperties();
                                delete.AppendFormat("DELETE FROM {0} ",
                                    schemaName.Length == 0 ? model.GetType().Name : schemaName + "." + model.GetType().Name);

                                for (int i = 0; i < properties.Length; i++)
                                {
                                        object[] attrs = properties[i].GetCustomAttributes(false);

                                        if (attrs.Length > 0)
                                        {
                                                if (deleteWhere.Length > 0)
                                                {
                                                        deleteWhere += " AND ";
                                                }

                                                if (attrs[0].GetType().Equals(typeof(KeyPropertyAttribute)))
                                                {
                                                        deleteWhere += string.Format(termFormat, properties[i].Name, properties[i].GetValue(model, null));
                                                }
                                        }
                                }

                                delete.Append(" WHERE " + deleteWhere);
                        }
                        catch (Exception e)
                        {
                                throw e;
                        }

                        return delete.ToString();
                }
                #endregion

                public static string BuildSQLConditionByModel(IModel model)
                {
                        StringBuilder whereClause = new StringBuilder();
                        string dbName = GetDatabaseType(model);
                        PropertyInfo[] properties = model.GetType().GetProperties();

                        #region 不同类型转换处理
                        for (int i = 0; i < properties.Length; i++)
                        {
                                if ((whereClause.Length > 0 && properties[i].PropertyType != typeof(DateTime)
                                    && properties[i].GetValue(model, null) != null) ||
                                    (whereClause.Length > 0 && (properties[i].PropertyType == typeof(DateTime) &&
                                        (Convert.ToDateTime(properties[i].GetValue(model, null))) != DateTime.MinValue)))
                                {
                                        whereClause.Append(" AND ");
                                }

                                if (properties[i].PropertyType == typeof(int) ||
                                            properties[i].PropertyType == typeof(decimal) ||
                                            properties[i].PropertyType == typeof(double) ||
                                            properties[i].PropertyType == typeof(float))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                whereClause.AppendFormat("{0}='{1}'",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(string) ||
                                            properties[i].PropertyType == typeof(Guid) ||
                                            properties[i].PropertyType == typeof(bool))
                                {
                                        if (properties[i].GetValue(model, null) != null)
                                        {
                                                whereClause.AppendFormat("{0}='{1}'",
                                                    properties[i].Name, properties[i].GetValue(model, null).ToString());
                                        }
                                }
                                else if (properties[i].PropertyType == typeof(DateTime))
                                {

                                        if (Convert.ToDateTime(properties[i].GetValue(model, null)) != DateTime.MinValue)
                                        {
                                                if (DBName.ORACLE.Equals(dbName, StringComparison.CurrentCultureIgnoreCase))
                                                {
                                                        whereClause.AppendFormat("{0}=TO_DATE('{1}','YYYY-MM-DD hh24:mi:ss')",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }
                                                else
                                                {
                                                        whereClause.AppendFormat("{0}='{1}'",
                                                            properties[i].Name, properties[i].GetValue(model, null).ToString());
                                                }
                                        }
                                }
                        }
                        #endregion

                        return whereClause.ToString();
                }


        }
}
