using System;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Xml.Serialization;
using System.Web.UI;

namespace EasyDev.Util
{
        [Obsolete("此工具类中的方法都已经有对应的扩展方法实现")]
        public class SerializeUtil
        {
                /// <summary>
                /// 实现从对象到XML串的序列化
                /// </summary>
                /// <param name="o"></param>
                /// <param name="type"></param>
                /// <returns></returns>
                public static string SerializeToXML(object o, Type type)
                {
                        return o.XmlSerialization(type);
                }

                /// <summary>
                /// 实现从对象到二进制串的序列化
                /// </summary>
                /// <param name="o"></param>
                /// <returns></returns>
                public static string SerializeToBinaryString(object o)
                {
                        return o.BinaryStringSerialization();
                }

                /// <summary>
                /// 实现从对象到LOS串的序列化 
                /// </summary>
                /// <param name="o"></param>
                /// <returns></returns>
                public static string SerializeToLOSString(object o)
                {
                        return o.LOSStringSerialization();
                }

                /// <summary>
                /// 从流中反序列化XML数据
                /// </summary>
                /// <param name="stream"></param>
                /// <returns></returns>
                /// <param name="type"></param>
                public static object DeserializeFromXMLStream(Stream stream, Type type)
                {
                        return stream.XmlDeserialization(type);
                }

                /// <summary>
                /// 从流中反序列化LOS数据
                /// </summary>
                /// <param name="stream"></param>
                /// <returns></returns>
                public static object DeserializeFromLOSStream(Stream stream)
                {
                        return stream.LOSDeserialization();
                }

                /// <summary>
                /// 将LOS数据反序列化为对象
                /// </summary>
                /// <param name="losString"></param>
                /// <returns></returns>
                public static object DeserializeFromLOSString(string losString)
                {
                        return losString.LOSDeserialization();
                }

                /// <summary>
                /// 从流中反序化二进制串数据
                /// </summary>
                /// <param name="stream"></param>
                /// <returns></returns>
                public static object DeserializeFromBinaryStream(Stream stream)
                {
                        return stream.BinaryDeserialization();
                }
        }
}
