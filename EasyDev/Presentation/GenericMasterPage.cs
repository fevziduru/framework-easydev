using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI;
using EasyDev.Resources;
using EasyDev.Presentation;
using System.Data;

namespace EasyDev.Presentation
{
    public class GenericMasterPage : MasterPage
    {
        private ActionFactory actionFactory = null;

        /// <summary>
        /// 本地化/国际化对象创建工厂
        /// </summary>
        private ResourceManagerFactory resMgrFactory = null;

        #region 事件
        private static object _recoveryUI = new object();

        /// <summary>
        /// 恢复界面事件
        /// </summary>
        public event EventHandler<EventArgs> OnRecoveryUI
        {
            add
            {
                Events.AddHandler(_recoveryUI, value);
            }
            remove
            {
                Events.RemoveHandler(_recoveryUI, value);
            }
        }

        private static object _initPage = new object();

        /// <summary>
        /// 初始化页面对象事件
        /// </summary>
        public event EventHandler<EventArgs> OnInitializePage
        {
            add
            {
                Events.AddHandler(_initPage, value);
            }
            remove
            {
                Events.RemoveHandler(_initPage, value);
            }
        }

        private static object _loadData = new object();

        /// <summary>
        /// 加载数据事件 
        /// </summary>
        public event EventHandler<EventArgs> OnLoadData
        {
            add
            {
                Events.AddHandler(_loadData, value);
            }
            remove
            {
                Events.RemoveHandler(_loadData, value);
            }
        }

        private static object _initUI = new object();

        /// <summary>
        /// 初始化界面UI事件
        /// </summary>
        public event EventHandler<EventArgs> OnInitializeUI
        {
            add
            {
                Events.AddHandler(_initUI, value);
            }
            remove
            {
                Events.RemoveHandler(_initUI, value);
            }
        }

        #endregion

        #region 公共属性
        /// <summary>
        /// 本地化/国际化对象
        /// </summary>
        public ResourceManagerFactory ResourceManagerCreator
        {
            get
            {
                return this.resMgrFactory;
            }
        }

        /// <summary>
        /// 资源管理器
        /// </summary>
        public IResourceManager ResourceManager
        {
            get;
            set;
        }

        /// <summary>
        /// Action工厂
        /// </summary>
        public ActionFactory ActionFactory
        {
            get
            {
                return this.actionFactory;
            }
        }

        public GenericMasterPage()
        {
            this.actionFactory = ActionFactory.CreateInstance();
            this.resMgrFactory = ResourceManagerFactory.GetInstance();      //国际化资源对象   
        }

        /// <summary>
        /// 界面元素数据
        /// </summary>
        protected DataSet UIData
        {
            get
            {
                DataSet dstmp = ViewState["Page_UIData"] as DataSet;
                if (dstmp == null)
                {
                    return new DataSet();
                }
                return dstmp;
            }
            set
            {
                if (ViewState["Page_UIData"] == null)
                {
                    ViewState.Add("Page_UIData", value);
                }
                else
                {
                    ViewState["Page_UIData"] = value;
                }
            }
        }

        #endregion

        /// <summary>
        /// 从界面元素获取数据
        /// </summary>
        protected virtual void GetDataFromUI() { }

        /// <summary>
        /// 页面加载事件
        /// </summary>
        /// <param name="e"></param>
        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            if (Events[_initPage] != null)
            {
                (Events[_initPage] as EventHandler<EventArgs>)(this, new EventArgs());
            }

            if (Events[_recoveryUI] != null)
            {
                (Events[_recoveryUI] as EventHandler<EventArgs>)(this, e);
            }

            if (!Page.IsPostBack)
            {
                if (Events[_loadData] != null)
                {
                    (Events[_loadData] as EventHandler<EventArgs>)(this, e);
                }

                if (Events[_initUI] != null)
                {
                    (Events[_initUI] as EventHandler<EventArgs>)(this, e);
                }
            }
            else
            {

            }
        }
    }
}
