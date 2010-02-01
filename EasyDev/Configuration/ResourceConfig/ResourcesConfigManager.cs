using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Collections;
using System.Web;
using System.Threading;

namespace EasyDev.Configuration
{
        /// <summary>
        /// 资源配置管理器
        /// </summary>
        public class ResourcesConfigManager
        {
                /// <summary>
                /// 配置文件位置
                /// </summary>
                private string configPath = string.Empty;

                /// <summary>
                /// 资源配置集合
                /// </summary>
                private IDictionary<string, ResourceItemConfig> resources = null;

                private static readonly object lockHelper = new object();

                private static ResourcesConfigManager _instance = null;

                public string DefaultResrouceAssembly
                {
                        get;
                        set;
                }

                private ResourcesConfigManager()
                {
                        this.configPath = AppDomain.CurrentDomain.BaseDirectory + @"Config\\EasyDev.Resources.Config";
                        this.resources = new Dictionary<string, ResourceItemConfig>();
                        this.Initialize();
                }

                /// <summary>
                /// 取得资源管理器实例
                /// </summary>
                /// <returns></returns>
                public static ResourcesConfigManager GetInstance()
                {

                        if (_instance == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(lockHelper, 100))
                                        {
                                                if (_instance == null)
                                                {
                                                        _instance = new ResourcesConfigManager();
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(lockHelper);
                                }
                        }

                        return _instance;
                }

                private void Initialize()
                {
                        //缓存配置信息
                        InitResourceConfig();
                }

                private void InitResourceConfig()
                {
                        try
                        {
                                XmlDocument doc = new XmlDocument();
                                doc.Load(this.configPath);
                                XmlNode resRoot = doc.SelectSingleNode("EasyDev.Resources.Config/Resources");
                                if (resRoot != null)
                                {
                                        if (resRoot.Attributes["DefaultAssembly"] != null)
                                        {
                                                this.DefaultResrouceAssembly = resRoot.Attributes["DefaultAssembly"].Value;
                                        }
                                        else
                                        {
                                                throw new System.Exception("_can't_find_DefaultAssembly_in_resources_node");
                                        }

                                        XmlNodeList resourceNodes = resRoot.SelectNodes("Resource");
                                        if (resourceNodes != null && resourceNodes.Count > 0)
                                        {
                                                IEnumerator itr_Nodes = resourceNodes.GetEnumerator();
                                                while (itr_Nodes.MoveNext())
                                                {
                                                        ResourceItemConfig newResourceItem = new ResourceItemConfig();
                                                        XmlNode currentNode = (XmlNode)itr_Nodes.Current;
                                                        string name = currentNode.SelectSingleNode("Name").InnerText;
                                                        if (name.Length == 0)
                                                        {
                                                                throw new System.Exception("_resource_must_have_a_name");
                                                        }
                                                        else
                                                        {
                                                                newResourceItem.Name = name;
                                                                newResourceItem.Assembly = currentNode.SelectSingleNode("Assembly").InnerText;
                                                                newResourceItem.BaseName = newResourceItem.Assembly + "." + name;
                                                                newResourceItem.Culture = currentNode.SelectSingleNode("Culture").InnerText;

                                                                KeyValuePair<string, ResourceItemConfig> item = new KeyValuePair<string, ResourceItemConfig>(name, newResourceItem);
                                                                this.resources.Add(item);
                                                        }
                                                }
                                        }
                                }
                                else
                                {
                                        throw new System.Exception("_can't_find_Resources_node_in_config_file");
                                }
                        }
                        catch (System.Exception e)
                        {
                                throw e;
                        }
                }

                /// <summary>
                /// 根据名称取得资源对象
                /// </summary>
                /// <param name="name"></param>
                /// <returns></returns>
                public ResourceItemConfig GetResourceItemByName(string name)
                {
                        return this.resources[name];
                }

                /// <summary>
                /// 判断对应名称的资源配置项在配置中是否存在
                /// </summary>
                /// <param name="name"></param>
                /// <returns></returns>
                public bool IsResourceItemExist(string name)
                {
                        return this.resources.Keys.Contains(name);
                }
        }
}
