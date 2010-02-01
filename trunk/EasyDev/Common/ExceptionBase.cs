using System;
using System.Collections.Generic;
using System.Text;

namespace EasyDev.Common
{
        /// <summary>
        /// 异常级别
        /// </summary>
        public sealed class ExceptionLevel
        {
                public static readonly string DEBUG = "DEBUG";

                public static readonly string WARNING = "WARNING";

                public static readonly string ERROR = "ERROR";
        }

        /// <summary>
        /// 异常处理公共基类
        /// </summary>
        public class ExceptionBase : System.Exception
        {
                /// <summary>
                /// 异常主题
                /// </summary>
                private string topic = "系统提示";

                /// <summary>
                /// 异常详细内容
                /// </summary>
                private string content = "底层框架产生了异常,但是没有设置异常信息描述";

                /// <summary>
                /// 异常发生的位置
                /// </summary>
                private string location = "";

                /// <summary>
                /// 异常发生时间
                /// </summary>
                private DateTime occurTime = DateTime.Now;

                /// <summary>
                /// 异常类型
                /// </summary>
                private string level = ExceptionLevel.DEBUG;

                #region 属性
                /// <summary>
                /// 异常类型
                /// </summary>
                public string Level
                {
                        get
                        {
                                return level;
                        }
                        set
                        {
                                level = value;
                        }
                }

                /// <summary>
                /// 异常主题
                /// </summary>
                public string Topic
                {
                        get
                        {
                                return this.topic;
                        }
                        set
                        {
                                this.topic = value;
                        }
                }

                /// <summary>
                /// 异常详细内容
                /// </summary>
                public string Content
                {
                        get
                        {
                                return this.content;
                        }
                        set
                        {
                                this.content = value;
                        }
                }

                /// <summary>
                /// 异常发生时间
                /// </summary>
                public DateTime OccurTime
                {
                        get
                        {
                                return this.occurTime;
                        }
                        set
                        {
                                this.occurTime = value;
                        }
                }

                /// <summary>
                /// 异常发生的位置
                /// </summary>
                public string Location
                {
                        get
                        {
                                return this.location;
                        }
                        set
                        {
                                this.location = value;
                        }
                }
                #endregion

                public ExceptionBase(string msg)
                        : base(msg)
                { }
        }
}
