using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace EasyDev.Security
{
    /// <summary>
    /// 用户验证接口
    /// </summary>
    public interface IUserAuthentication  
    {
        /// <summary>
        /// 用户登入方法
        /// </summary>
        /// <param name="context"></param>
        void CheckIn(IDictionary context);

        /// <summary>
        /// 登出方法
        /// </summary>
        void CheckOut();

        /// <summary>
        /// 验证某用户对某资源是否有某种操作类型的权限
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="resourceId"></param>
        /// <param name="operationId"></param>
        /// <returns></returns>
        bool Authentication(string userName, string resourceId, string operationId);

        /// <summary>
        /// 验证某用户对某资源是否有任一操作类型的权限
        /// </summary>
        /// <param name="userName"></param>
        /// <param name="resourceId"></param>
        /// <returns></returns>
        bool Authentication(string userName, string resourceId);
    }
}
