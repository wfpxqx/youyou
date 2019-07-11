using System;
using DevExpress.Xpf.Core;

namespace CASCRS_Voucher_Import.Common.CommCheck
{
	public static class UICheck
    {
        /// <summary>
        /// 检测DXTabControl中是否存在重复的DXTabItem
        /// </summary>
        /// <param name="Header">TabItem的Header字符串</param>
        /// <param name="tabControl">当前DXTabControl</param>
        /// <returns>存在标识</returns>
        public static bool CheckDXTabControlRepeatItem(string Header, DXTabControl tabControl, out DXTabItem CurrentTabItem)
        {
            bool HasItem = false;
			CurrentTabItem = null;
			try
            {
                foreach (DXTabItem tabItem in tabControl.Items)
                {
                    if (Header.Equals(tabItem.Header.ToString()))
                    {
                        HasItem = true;
						CurrentTabItem = tabItem;
						break;
                    }
                }
                return HasItem; 
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 检测DXTabControl中是否看见，若不可见在添加DXTabControl是可见
        /// </summary>
        /// <param name="tabControl">当前DXTabControl</param>
        /// <returns>可见标识</returns>
        public static bool CheckDXTabControlVisible(DXTabControl tabControl)
        {
            bool Visible = false;
            try
            {
                int count = tabControl.Items.Count;
                if (count > 0) Visible = true;
                return Visible;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
