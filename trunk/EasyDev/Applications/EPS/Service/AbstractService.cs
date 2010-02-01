using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.EPS.BusinessObject;
using System.Web;
using EasyDev.EPS.Portal;
using WebCaching = System.Web.Caching;
using System.Threading;

namespace EasyDev.EPS
{
        public abstract class AbstractService : IService
        {
                private static readonly string DEFAULT_BO = "DEFAULT_BO";

                /// <summary>
                /// 业务对象缓存
                /// </summary>        
                private WebCaching.Cache BOCache { get; set; }

                /// <summary>
                /// 通用业务对象
                /// </summary>
                protected virtual IBO DefaultBO { get; set; }

                /// <summary>
                /// 构造方法
                /// </summary>
                public AbstractService()
                {
                        Initialize();
                }

                /// <summary>
                /// 初始化方法，此方法会实例化一个集合类型对象，这个对象用于存放缓存的业务对象
                /// </summary>
                protected virtual void Initialize()
                {
                        BOCache = HttpRuntime.Cache;
                }

                /// <summary>
                /// 获取通用业务对象
                /// </summary>
                /// <typeparam name="TModel"></typeparam>
                /// <returns></returns>
                protected virtual GenericBO<TModel> GetDefaultBO<TModel>() where TModel : IModel, new()
                {
                        GenericBO<TModel> tmpbo = BOCache.Get(DEFAULT_BO) as GenericBO<TModel>;


                        if (tmpbo == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(BOCache, 100))
                                        {
                                                if (tmpbo == null)
                                                {
                                                        tmpbo = BOFactory.CreateBO<GenericBO<TModel>>();
                                                        BOCache.Insert(DEFAULT_BO, tmpbo);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(BOCache);
                                }

                        }

                        return tmpbo;
                }

                /// <summary>
                /// 获取业务对象
                /// </summary>
                /// <typeparam name="TBO"></typeparam>
                /// <returns></returns>
                protected virtual TBO GetBO<TBO>() where TBO : AbstractBO, new()
                {
                        object tmpbo = BOCache.Get(typeof(TBO).Name);

                        if (tmpbo == null)
                        {
                                try
                                {
                                        if (Monitor.TryEnter(BOCache, 100))
                                        {
                                                if (tmpbo == null)
                                                {
                                                        tmpbo = BOFactory.CreateBO<TBO>();
                                                        BOCache.Insert(typeof(TBO).Name, tmpbo);
                                                }
                                        }
                                }
                                catch (ArgumentException e)
                                {
                                        throw e;
                                }
                                finally
                                {
                                        Monitor.Exit(BOCache);
                                }
                        }

                        return (tmpbo as TBO);
                }

                protected virtual object GetBO(Type boType)
                {
                        return Activator.CreateInstance(boType);
                }
        }
}
