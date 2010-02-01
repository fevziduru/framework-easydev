using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;
using System.Data;
using System.Xml.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Web.UI;
using System.Reflection;

namespace EasyDev.Util
{
    /// <summary>
    /// 扩展工具方法
    /// </summary>
    public static class UtilityExtension
    {
        #region 加压解压方法
        /// <summary>
        /// 采用GZIP算法对数据进行压缩
        /// </summary>
        /// <param name="toCompressData">要压缩的数据</param>
        /// <returns>UTF8编码的压缩结果</returns>
        public static string Compress(this byte[] toCompressData)
        {
            MemoryStream stream = new MemoryStream();
            GZipStream zipStream = new GZipStream(stream, CompressionMode.Compress, false);
            try
            {
                zipStream.Write(toCompressData, 0, toCompressData.Length);
            }
            catch (CompressException e)
            {
                throw e;
            }
            finally
            {
                if (zipStream != null)
                {
                    zipStream.Flush();
                }

                if (stream != null)
                {
                    stream.Close();
                }
            }

            return Encoding.UTF8.GetString(stream.ToArray());
        }

        /// <summary>
        /// 对UTF8编码的压缩数据进行解压缩
        /// </summary>
        /// <param name="toExtractData">要解压的数据</param>
        /// <returns>UTF8编码的解压缩数据</returns>
        public static string Extract(this byte[] toExtractData)
        {
            GZipStream zipStream = new GZipStream(new MemoryStream(toExtractData), CompressionMode.Decompress, true);

            StreamReader reader = null;
            string result = "";
            try
            {
                reader = new StreamReader(zipStream);
                result = reader.ReadToEnd();
            }
            catch (ExtractException e)
            {
                throw e;
            }
            finally
            {
                if (zipStream != null)
                {
                    zipStream.Flush();
                }
            }

            return result;
        }

        /// <summary>
        /// 采用GZIP算法对数据进行压缩
        /// </summary>
        /// <param name="toCompressData">要压缩的数据</param>
        /// <returns>UTF8编码的压缩结果</returns>
        public static string Compress(this string toCompressData)
        {
            return Compress(Encoding.UTF8.GetBytes(toCompressData));
        }

        /// <summary>
        /// 对UTF8编码的压缩数据进行解压缩
        /// </summary>
        /// <param name="toExtractData">要解压的数据</param>
        /// <returns>UTF8编码的解压缩数据</returns>
        public static string Extract(this string toExtractData)
        {
            return Extract(Encoding.UTF8.GetBytes(toExtractData));
        }
        #endregion

        #region 数据转换方法
        /// <summary>
        /// 转换为字符串，如果为空则返回string.Empty
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static string ParseString(this object obj)
        {
            try
            {
                return Convert.ToString(obj).Trim();
            }
            catch { }

            return string.Empty;
        }

        /// <summary>
        /// 将对象转换成DateTime，如果无法转换则返回值为0001-1-1 00:00:00的DateTime（new DateTime();）
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static DateTime ParseDateTime(this object obj)
        {
            try
            {
                DateTime d = DateTime.Parse(ConvertUtil.ToStringOrDefault(obj));
                return d;
            }
            catch
            {
                return new DateTime();
            }
        }

        /// <summary>
        /// 将对象转换成int，如果无法转换则返回0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static int ParseInt(this object obj)
        {
            //int result = 0;

            //int.TryParse(obj.ParseString(), out result);
            //return result;

            try
            {
                return Convert.ToInt32(obj);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 将对象转换成decimal，如果无法转换则返回0
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static decimal ParseDecimal(this object obj)
        {
            try
            {
                return Convert.ToDecimal(obj);
            }
            catch
            {
                return 0;
            }
        }

        /// <summary>
        /// 转换成纬度，正数为北纬，负数为南纬，默认值为0度
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static decimal ParseLatitude(this object obj)
        {
            decimal value = 0;
            try
            {
                value = Convert.ToDecimal(obj);
                if (value > 90)
                {
                    value = 90;
                }
                else if (value < -90)
                {
                    value = -90;
                }
            }
            catch
            { }
            return value;
        }

        /// <summary>
        /// 转换成经度，正数为东经，负数为西经，默认值为0度
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public static decimal ParseLongitude(this object obj)
        {
            decimal value = 0;
            try
            {
                value = Convert.ToDecimal(obj);
                if (value > 180)
                {
                    value = 180;
                }
                else if (value < -180)
                {
                    value = -180;
                }
            }
            catch
            { }
            return value;
        }

        /// <summary>
        /// 转换成bool类型，特别的，"Y"转换成True，"N"转换成False
        /// </summary>
        /// <param name="p"></param>
        /// <returns></returns>
        public static bool ParesBool(this object p)
        {
            string str = p.ParseString();
            if ("N".Equals(str))
            {
                return false;
            }
            if ("Y".Equals(str))
            {
                return true;
            }
            try
            {
                return Convert.ToBoolean(p);
            }
            catch { }
            return false;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="d"></param>
        /// <returns></returns>
        public static decimal[] ResolveReal(this decimal d)
        {
            List<decimal> res = new List<decimal>();
            decimal integer = Math.Floor(d);
            if (integer >= 0.0M)
            {
                res.Add(integer);
            }
            else
            {
                return res.ToArray();
            }
            decimal dec = d - integer;
            res.Add(dec);
            return res.ToArray();
        }
        #endregion

        #region 数据集工具方法
        /// <summary>
        /// 获取数据集中的数据表
        /// </summary>
        /// <param name="dsFrom">数据集</param>
        /// <param name="tableName">数据表名称</param>
        /// <returns></returns>
        public static DataTable Table(this DataSet dsFrom, string tableName)
        {
            DataTable dt = null;
            try
            {
                if (dsFrom.Tables.Contains(tableName))
                {
                    dt = dsFrom.Tables[tableName].Copy();  //以复制的方式取得数据集中的表，不破坏原有的数据集
                    //对于有自关联的表，要先删除其自关联关系，否则无法移除表
                    //if (dsFrom.Relations.Count > 0)
                    //{
                    //    dsFrom.Relations.Clear();
                    //}

                    //dsFrom.Tables.Remove(tableName);
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取数据集中的数据表
        /// </summary>
        /// <param name="dsFrom">源数据集</param>
        /// <param name="index">数据表索引</param>
        /// <returns></returns>
        public static DataTable Table(this DataSet dsFrom, int index)
        {
            DataTable dt = null;
            try
            {
                if (dsFrom.Tables[index] != null)
                {
                    dt = dsFrom.Tables[index].Copy();  //以复制的方式取得数据集中的表，不破坏原有的数据集
                    //对于有自关联的表，要先删除其自关联关系，否则无法移除表
                    //if (dsFrom.Relations.Count > 0)
                    //{
                    //    dsFrom.Relations.Clear();
                    //}

                    //dsFrom.Tables.RemoveAt(index);
                    return dt;
                }
                else
                {
                    return null;
                }
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取数据集中的一行
        /// </summary>
        /// <param name="dsFrom">数据集</param>
        /// <param name="tableName">数据表名称</param>
        /// <param name="rowNum">行号</param>
        /// <returns></returns>
        public static DataRow SingleRow(this DataSet dsFrom, string tableName, int rowNum)
        {
            try
            {
                if (dsFrom.Tables.Contains(tableName) && dsFrom.Tables[tableName].Rows.Count > 0)
                {
                    return dsFrom.Tables[tableName].Rows[rowNum];
                }
                else
                {
                    return null;
                }
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取数据表中的一行
        /// </summary>
        /// <param name="dtFrom">数据表</param>
        /// <param name="rowNum">行号</param>
        /// <returns></returns>
        public static DataRow SingleRow(this DataTable dtFrom, int rowNum)
        {
            try
            {
                if (dtFrom.Rows.Count > 0)
                {
                    return dtFrom.Rows[rowNum];
                }
                else
                {
                    return null;
                }
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取数据集中某数据表的第一行
        /// </summary>
        /// <param name="dsFrom">数据集</param>
        /// <param name="tableName">数据表名称</param>
        /// <returns></returns>
        public static DataRow FirstRow(this DataSet dsFrom, string tableName)
        {
            try
            {
                return dsFrom.SingleRow(tableName, 0);
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 获取数据表中的第一行
        /// </summary>
        /// <param name="dtFrom">数据表</param>
        /// <returns></returns>
        public static DataRow FirstRow(this DataTable dtFrom)
        {
            try
            {
                return dtFrom.SingleRow(0);
            }
            catch (ConvertException e)
            {
                throw e;
            }
        }
        #endregion

        #region 序列化方法
        /// <summary>
        /// 实现从对象到XML串的序列化
        /// </summary>
        /// <param name="o"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string XmlSerialization(this object o, Type type)
        {
            string xmlStr;
            MemoryStream streamStoreObj = new MemoryStream();

            XmlSerializer xmlS = new XmlSerializer(type);
            xmlS.Serialize(streamStoreObj, o);
            xmlStr = streamStoreObj.QuickReader(true);

            return xmlStr;
        }

        /// <summary>
        /// 实现从对象到二进制串的序列化
        /// </summary>
        /// <param name="o"></param>
        /// <returns></returns>
        public static string BinaryStringSerialization(this object o)
        {
            string result;

            BinaryFormatter bFormatter = new BinaryFormatter();
            MemoryStream stream = new MemoryStream();
            bFormatter.Serialize(stream, o);

            result = stream.QuickReader(true);

            return result;
        }

        /// <summary>
        /// 实现从对象到LOS串的序列化 
        /// </summary>
        /// <param name="o"></param>
        /// <returns></returns>
        public static string LOSStringSerialization(this object o)
        {
            string result;

            LosFormatter los = new LosFormatter();
            MemoryStream stream = new MemoryStream();
            los.Serialize(stream, o);

            result = stream.QuickReader(true);

            return result;
        }

        /// <summary>
        /// 从流中反序列化XML数据
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        /// <param name="type"></param>
        public static object XmlDeserialization(this Stream stream, Type type)
        {
            object result;
            XmlSerializer xmlS = new XmlSerializer(type);
            result = xmlS.Deserialize(stream);

            return result;
        }

        /// <summary>
        /// 从流中反序列化LOS数据
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static object LOSDeserialization(this Stream stream)
        {
            object result;
            LosFormatter los = new LosFormatter();
            result = los.Deserialize(stream);

            return result;
        }

        /// <summary>
        /// 将LOS数据反序列化为对象
        /// </summary>
        /// <param name="losString"></param>
        /// <returns></returns>
        public static object LOSDeserialization(this string losString)
        {
            object result;
            LosFormatter los = new LosFormatter();
            result = los.Deserialize(losString);

            return result;
        }

        /// <summary>
        /// 从流中反序化二进制串数据
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        public static object BinaryDeserialization(this Stream stream)
        {
            object result;

            BinaryFormatter bFormatter = new BinaryFormatter();
            result = bFormatter.Deserialize(stream);

            return result;
        }
        #endregion

        #region 流处理方法
        /// <summary>
        /// 快速流读取
        /// </summary>
        /// <param name="stream">包含数据的流对象</param>
        /// <param name="isCloseStream">读完数据后是否关闭流</param>
        /// <returns></returns>
        public static string QuickReader(this Stream stream, bool isClose)
        {
            try
            {
                string tmp = string.Empty;
                using (BufferedStream bstream = new BufferedStream(stream))
                {
                    using (StreamReader reader = new StreamReader(bstream))
                    {
                        tmp = reader.ReadToEnd();
                    }
                }

                return tmp;
            }
            catch (IOException e)
            {
                throw e;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static byte[] QuickReaderWithEncoding(this Stream stream, Encoding encoder)
        {
            return encoder.GetBytes(stream.QuickReader(true));
        }

        public static byte[] QuickReaderWithEncoding(this Stream stream, Encoding encoder, bool isClose)
        {
            return encoder.GetBytes(stream.QuickReader(isClose));
        }

        /// <summary>
        /// 从流中读取数据并进行BASE64编码
        /// </summary>
        /// <param name="stream">包含数据的流对象</param>
        /// <param name="isCloseStream">读完数据后是否关闭流</param>
        /// <returns></returns>
        public static string QuickReaderAsBase64(this Stream stream, bool isClose)
        {
            return Convert.ToBase64String(QuickReaderWithEncoding(stream, Encoding.UTF8, isClose));
        }

        /// <summary>
        /// 从流中读取数据并进行BASE64编码
        /// </summary>
        /// <param name="stream">包含数据的流对象</param>
        /// <param name="isCloseStream">读完数据后是否关闭流</param>
        /// <returns></returns>
        public static string QuickReaderAsBase64(this Stream stream)
        {
            return Convert.ToBase64String(QuickReaderWithEncoding(stream, Encoding.UTF8, true));
        }

        /// <summary>
        /// 从流中读取数据并进行UTF8编码
        /// </summary>
        /// <param name="stream">包含数据的流对象</param>
        /// <param name="isCloseStream">读完数据后是否关闭流</param>
        /// <returns></returns>
        public static byte[] QuickReaderAsUTF8(this Stream stream, bool isClose)
        {
            return stream.QuickReaderWithEncoding(Encoding.UTF8, isClose);
        }

        /// <summary>
        /// 从流中读取数据并进行UTF8编码
        /// </summary>
        /// <param name="stream">包含数据的流对象</param>
        /// <param name="isCloseStream">读完数据后是否关闭流</param>
        /// <returns></returns>
        public static byte[] QuickReaderAsUTF8(this Stream stream)
        {
            return stream.QuickReaderWithEncoding(Encoding.UTF8, true);
        }

        /// <summary>
        /// 快速流写入器
        /// </summary>
        /// <param name="stream">要写入数据的基础流对象</param>
        /// <param name="data">要写入的数据</param>
        /// <param name="isCloseStream">是否关闭流对象</param>
        /// <returns></returns>
        public static void QuickWriter(this Stream stream, string data, bool isCloseStream)
        {
            try
            {
                using (BufferedStream bstream = new BufferedStream(stream))
                {
                    using (StreamWriter writer = new StreamWriter(bstream))
                    {
                        writer.AutoFlush = true;
                        writer.Write(data);
                    }
                }
            }
            catch (IOException e)
            {
                throw e;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// 快速流写入器
        /// </summary>
        /// <param name="stream">要写入数据的基础流对象</param>
        /// <param name="data">要写入的数据</param>
        /// <returns></returns>
        public static void QuickWriter(this Stream stream, string data)
        {
            stream.QuickWriter(data, true);
        }

        /// <summary>
        /// 快速流写入器
        /// </summary>
        /// <param name="stream">要写入数据的基础流对象</param>
        /// <param name="data">要写入的数据</param>
        /// <param name="isCloseStream">是否关闭流对象</param>
        /// <returns></returns>
        public static void QuickWriter(this Stream stream, byte[] data, bool isClose)
        {
            stream.QuickWriter(Encoding.UTF8.GetString(data), isClose);
        }

        /// <summary>
        /// 快速流写入器
        /// </summary>
        /// <param name="stream">要写入数据的基础流对象</param>
        /// <param name="data">要写入的数据</param>
        /// <returns></returns>
        public static void QuickWriter(this Stream stream, byte[] data)
        {
            stream.QuickWriter(Encoding.UTF8.GetString(data), true);
        }
        #endregion

        #region 类型反射工具方法
        /// <summary>
        /// 
        /// </summary>
        /// <param name="t"></param>
        /// <returns></returns>
        public static IDictionary<string, PropertyInfo> Properties(this Type t)
        {
            PropertyInfo[] properties = t.GetProperties();
            IDictionary<string, PropertyInfo> result = new Dictionary<string, PropertyInfo>();

            for (int i = 0; i < properties.Length; i++)
            {
                result.Add(properties[i].Name, properties[i]);
            }

            return result;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tObject"></param>
        /// <param name="propName"></param>
        /// <returns></returns>
        public static object ValueOfProperty(this Type tObject, string propName)
        {
            IDictionary<string, PropertyInfo> items = tObject.Properties();
            object result = null;

            foreach (KeyValuePair<string, PropertyInfo> item in items)
            {
                if (item.Key.Equals(propName, StringComparison.CurrentCultureIgnoreCase))
                {
                    result = item.Value.GetValue(tObject, null);
                }
            }

            return result;
        }

        /// <summary>
        /// 判断类型t中是否有attr这个自定义属性
        /// </summary>
        /// <param name="t"></param>
        /// <param name="attr"></param>
        /// <returns></returns>
        public static bool HasAttribute(this Type t, Type attr)
        {
            return t.GetCustomAttributes(false).FirstOrDefault(p => p.GetType().Equals(attr)) != null;
        }

        /// <summary>
        /// 判断类型中的指定属性(Property)是否有某个特性(Attribute)
        /// </summary>
        /// <param name="tobject"></param>
        /// <param name="propName"></param>
        /// <param name="tattr"></param>
        /// <returns></returns>
        public static object HasAttribute(this Type tobject, string propName, Type tattr)
        {
            IDictionary<string, PropertyInfo> properties = tobject.Properties();
            PropertyInfo prop = properties.FirstOrDefault(p => p.Key == propName).Value;
            object[] attrs = prop.GetCustomAttributes(false);
            object result = null;

            if (attrs != null)
            {
                result = attrs.FirstOrDefault(p => p.GetType().Equals(tattr));
            }

            return result;
        }

        /// <summary>
        /// 判断类型TObject中的名为propName的属性是否有类型为TAttribute的属性
        /// 返回空则不存在
        /// </summary>
        /// <param name="property"></param>
        /// <param name="type"></param>
        /// <param name="tattr"></param>
        /// <returns></returns>
        public static object HasAttribute(this PropertyInfo property, Type type, Type tattr)
        {
            return type.HasAttribute(property.Name, tattr);
        }

        /// <summary>
        /// 从类型tObject的名为propName的成员中找到类型为tAttr的自定义属性，并取得这个自定义属性中名为fieldName的公共成员的值
        /// </summary>
        /// <param name="tObject"></param>
        /// <param name="propName"></param>
        /// <param name="tAttr"></param>
        /// <param name="fieldName"></param>
        /// <returns></returns>
        public static object ValueOfAttribute(this Type tObject, string propName, Type tAttr, string fieldName)
        {
            object attr = tObject.HasAttribute(tAttr);

            if (attr != null)
            {
                return attr.GetType().GetProperty(fieldName).GetValue(attr, null);
            }
            else
            {
                return null;
            }
        }
        #endregion
    }
}
