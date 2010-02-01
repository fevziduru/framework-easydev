using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EasyDev.BL;
using System.ComponentModel;
using System.Data;

namespace EasyDev.Presentation
{
    public class GenericAction : IAction
    {
        private NativeServiceManager serviceManager = null;

        protected NativeServiceManager ServiceManager
        {
            get
            {
                return this.serviceManager;
            }
        }

        #region Events
        private static object _saving = new object();
        private static object _saved = new object();
        private static object _deleting = new object();
        private static object _deleted = new object();
        private static object _loading = new object();
        private static object _loaded = new object();
        protected EventHandlerList Events = new EventHandlerList();

        /// <summary>
        /// 加载数据后处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Loaded
        {
            add
            {
                Events.AddHandler(_loaded, value);
            }
            remove
            {
                Events.RemoveHandler(_loaded, value);
            }
        }

        /// <summary>
        /// 加载数据后处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Loading
        {
            add
            {
                Events.AddHandler(_loading, value);
            }
            remove
            {
                Events.RemoveHandler(_loading, value);
            }
        }

        /// <summary>
        /// 删除后处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Deleted
        {
            add
            {
                Events.AddHandler(_deleted, value);
            }
            remove
            {
                Events.RemoveHandler(_deleted, value);
            }
        }

        /// <summary>
        /// 删除前处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Deleting
        {
            add
            {
                Events.AddHandler(_deleting, value);
            }
            remove
            {
                Events.RemoveHandler(_deleting, value);
            }
        }

        /// <summary>
        /// 保存前处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Saving
        {
            add
            {
                Events.AddHandler(_saving, value);
            }
            remove
            {
                Events.RemoveHandler(_saving, value);
            }
        }

        /// <summary>
        /// 保存后处理事件
        /// </summary>
        public event EventHandler<OperationEventArgs> Saved
        {
            add
            {
                Events.AddHandler(_saved, value);
            }
            remove
            {
                Events.RemoveHandler(_saved, value);
            }
        }
        #endregion

        public GenericAction()
        {
            this.serviceManager = ServiceManagerFactory.CreateServiceManager<NativeServiceManager>();

            InitializeAction();
        }

        #region IAction Members

        /// <summary>
        /// 初始化Action
        /// </summary>
        public virtual void InitializeAction() { }
        
        #endregion

        protected delegate void OnSaveCallbackHandler(DataSet ds);
        protected delegate void OnDeleteCallbackHander(DataSet ds);
        protected delegate DataSet OnLoadByPrimaryCallbackHandler(object key);

        /// <summary>
        /// 保存处理方法，此方法会触发保存前处理事件和保存后处理事件
        /// </summary>
        /// <param name="ds">要保存的数据集</param>
        /// <param name="callback">保存处理方法代理</param>
        protected virtual void OnSave(DataSet ds,OnSaveCallbackHandler save)
        {
            OperationEventArgs args = new OperationEventArgs(ds);

            if (Events[_saving] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_saving])(this, args);
            }

            save(ds);
            args.Data = ds;

            if (Events[_saved] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_saved])(this, args);
            }
        }

        /// <summary>
        /// 删除处理方法，此方法会触发删除前处理事件和删除后处理事件
        /// </summary>
        /// <param name="ds">要删除的数据</param>
        /// <param name="delete">删除处理方法代理</param>
        protected virtual void OnDelete(DataSet ds, OnDeleteCallbackHander delete)
        {
            OperationEventArgs args = new OperationEventArgs(ds);
            if (Events[_deleting] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_deleting])(this, args);
            }

            delete(ds);
            args.Data = ds;

            if (Events[_deleted] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_deleted])(this, args);
            }
        }

        /// <summary>
        /// 加载数据处理方法
        /// </summary>
        /// <param name="ds"></param>
        /// <param name="load"></param>
        protected virtual DataSet OnLoadByPrimary(object key, OnLoadByPrimaryCallbackHandler load)
        {
            OperationEventArgs args = new OperationEventArgs(key);
            if (Events[_loading] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_loading])(this, args);
            }

            args.Data = load(key);

            if (Events[_loaded] != null)
            {
                ((EventHandler<OperationEventArgs>)Events[_loaded])(this, args);
            }

            return args.Data;
        }
    }
}
