using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Compression;
using System.IO;
using com.dlut.ssdut.zucky.DotNetFetion;

namespace EasyDev.Util
{
        [Obsolete("此工具类中的方法都已经有对应的扩展方法实现")]
        public class CompressUtil
        {
                /// <summary>
                /// 采用GZIP算法对数据进行压缩
                /// </summary>
                /// <param name="toCompressData">要压缩的数据</param>
                /// <returns>UTF8编码的压缩结果</returns>
                public static string Compress(byte[] toCompressData)
                {
                        return toCompressData.Compress();
                }

                /// <summary>
                /// 对UTF8编码的压缩数据进行解压缩
                /// </summary>
                /// <param name="toExtractData">要解压的数据</param>
                /// <returns>UTF8编码的解压缩数据</returns>
                public static string Extract(byte[] toExtractData)
                {
                        return toExtractData.Extract();
                }

                /// <summary>
                /// 采用GZIP算法对数据进行压缩
                /// </summary>
                /// <param name="toCompressData">要压缩的数据</param>
                /// <returns>UTF8编码的压缩结果</returns>
                public static string Compress(string toCompressData)
                {
                        return toCompressData.Compress();
                }

                /// <summary>
                /// 对UTF8编码的压缩数据进行解压缩
                /// </summary>
                /// <param name="toExtractData">要解压的数据</param>
                /// <returns>UTF8编码的解压缩数据</returns>
                public static string Extract(string toExtractData)
                {
                        return toExtractData.Extract();
                }
        }
}
