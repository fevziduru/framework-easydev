using System;
using System.Reflection;
using System.IO;
using System.Threading;
using System.Xml;
using EasyDev.Resources;

namespace EasyDev.Resources
{
        /// <summary>
        /// 国际化对象管理器
        /// </summary>
        public class GlobalizationManager
        {
                private static object _lockHelper = new object();

                private static GlobalizationManager _instance = null;

                private static Stream resouceStream = null;

                /// <summary>
                /// 初始化国际化资源对象，加载语言资源所在的程序集
                /// </summary>
                /// <param name="assemblyName">国际化资源所在的程序集名称</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <param name="culture">国际化资源的语言包名称，对应标准的国际化语言代码</param>
                private GlobalizationManager(string assemblyName, string baseName, string culture)
                {
                        string path = string.Format(@"{0}.{1}_{2}.xml", assemblyName, baseName, culture);
                        resouceStream = Assembly.Load(assemblyName).GetManifestResourceStream(path);
                }

                /// <summary>
                /// 初始化国际化资源对象
                /// </summary>
                /// <param name="assemblyType">国际化资源所在的程序集类型</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <param name="culture">国际化资源的语言包名称，对应标准的国际化语言代码</param>
                private GlobalizationManager(Type assemblyType, string baseName, string culture)
                {
                        Assembly ass = Assembly.GetAssembly(assemblyType);
                        string path = string.Format(@"{0}.{1}_{2}.xml", ass.GetName(), baseName, culture);
                        resouceStream = ass.GetManifestResourceStream(path);
                }

                /// <summary>
                /// 实例化国际化对象
                /// </summary>
                /// <param name="assemblyName">国际化资源所在的程序集名称</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <param name="culture">国际化资源的语言包名称，对应标准的国际化语言代码</param>
                /// <returns></returns>
                public static GlobalizationManager GetInstance(string assemblyName, string resLocation, string culture)
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new GlobalizationManager(assemblyName, resLocation, culture);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 实例化国际化对象
                /// </summary>
                /// <param name="assemblyType">国际化资源所在的程序集类型</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <param name="culture">国际化资源的语言包名称，对应标准的国际化语言代码</param>
                /// <returns></returns>
                public static GlobalizationManager GetInstance(Type assemblyType, string resLocation, string culture)
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new GlobalizationManager(assemblyType, resLocation, culture);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 根据当前线程中的Culture信息实例化国际化对象
                /// </summary>
                /// <param name="assemblyType">国际化资源所在的程序集类型</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <returns></returns>
                public static GlobalizationManager GetInstance(Type assemblyType, string resLocation)
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new GlobalizationManager(assemblyType, resLocation, Thread.CurrentThread.CurrentCulture.Name);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 根据当前线程中的Culture信息实例化国际化对象
                /// </summary>
                /// <param name="assemblyName">国际化资源所在的程序集名称</param>
                /// <param name="resLocation">国际化资源所在的文件名称</param>
                /// <returns></returns>
                public static GlobalizationManager GetInstance(string assemblyName, string resLocation)
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(_lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new GlobalizationManager(assemblyName, resLocation, Thread.CurrentThread.CurrentCulture.Name);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(_lockHelper);
                                }
                        }

                        return _instance;
                }

                /// <summary>
                /// 取得国际化器
                /// </summary>
                /// <returns></returns>
                public IGlobalizer GetGlobalizer()
                {
                        string languageInfo = string.Empty;
                        StreamReader reader = null;
                        IGlobalizer globalizer = null;
                        XmlDocument doc = new XmlDocument();

                        try
                        {
                                if (resouceStream != null)
                                {
                                        resouceStream.Seek(0, SeekOrigin.Begin);
                                        reader = new StreamReader(resouceStream);

                                        //读出国际化语言信息
                                        languageInfo = reader.ReadToEnd();
                                        doc.LoadXml(languageInfo);

                                        globalizer = new Globalizer(doc);
                                }
                                else
                                {
                                        throw new GlobalizeException("_cannot_find_language_resource");
                                }
                        }
                        catch (GlobalizeException e)
                        {
                                throw e;
                        }

                        return globalizer;
                }

                ~GlobalizationManager()
                {
                        if (resouceStream != null)
                        {
                                resouceStream.Dispose();
                                resouceStream.Close();
                        }
                }
        }
}
