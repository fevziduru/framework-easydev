using System;
using System.IO;
using System.Text;

namespace EasyDev.Util
{
        [Obsolete("此工具类中的方法都已经有对应的扩展方法实现")]
        public class StreamUtil
        {
                /// <summary>
                /// 快速流读取
                /// </summary>
                /// <param name="stream">包含数据的流对象</param>
                /// <param name="isCloseStream">读完数据后是否关闭流</param>
                /// <returns></returns>
                public static string QuickStreamReader(Stream stream, bool isCloseStream)
                {
                        return stream.QuickReader(isCloseStream);
                }

                /// <summary>
                /// 从流中读取数据并进行BASE64编码
                /// </summary>
                /// <param name="stream">包含数据的流对象</param>
                /// <param name="isCloseStream">读完数据后是否关闭流</param>
                /// <returns></returns>
                public static string QuickStreamReaderWithBase64(Stream stream, bool isCloseStream)
                {
                        return System.Convert.ToBase64String(QuickStreamReaderWithUTF8Encoded(stream, isCloseStream));
                }

                /// <summary>
                /// 从流中读取数据并进行UTF8编码
                /// </summary>
                /// <param name="stream">包含数据的流对象</param>
                /// <param name="isCloseStream">读完数据后是否关闭流</param>
                /// <returns></returns>
                public static byte[] QuickStreamReaderWithUTF8Encoded(Stream stream, bool isCloseStream)
                {
                        return Encoding.UTF8.GetBytes(QuickStreamReader(stream, isCloseStream));
                }

                /// <summary>
                /// 快速流写入器
                /// </summary>
                /// <param name="stream">要写入数据的基础流对象</param>
                /// <param name="data">要写入的数据</param>
                /// <param name="isCloseStream">是否关闭流对象</param>
                /// <returns></returns>
                public static void QuickStreamWriter(Stream stream, string data, bool isCloseStream)
                {
                        stream.QuickWriter(data, isCloseStream);
                }

                /// <summary>
                /// 快速流写入器
                /// </summary>
                /// <param name="stream">要写入数据的基础流对象</param>
                /// <param name="data">要写入的数据</param>
                /// <param name="isCloseStream">是否关闭流对象</param>
                /// <returns></returns>
                public static void QuickStreamWriter(Stream stream, byte[] data, bool isCloseStream)
                {
                        stream.QuickWriter(data, isCloseStream);
                }
        }
}
