using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.UI.Design;
using System.ComponentModel;

namespace EasyDev.Presentation.UI.WebControls.Design
{
    /// <summary>
    /// 
    /// </summary>
    public class DropDownTreeDesign : ControlDesigner
    {
        private DropDownTree dropdownTree = new DropDownTree();

        /// <summary>
        /// 
        /// </summary>
        /// <param name="component"></param>
        public override void Initialize(IComponent component)
        {
            base.Initialize(component);

            dropdownTree = (DropDownTree)component;
        }

        /// <summary>
        /// 
        /// </summary>
        public override bool AllowResize
        {
            get
            {
                return true;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public override string GetDesignTimeHtml()
        {
            return dropdownTree.DesignerSupport();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        protected override string GetEmptyDesignTimeHtml()
        {
            return base.GetEmptyDesignTimeHtml();
        }
    }
}
