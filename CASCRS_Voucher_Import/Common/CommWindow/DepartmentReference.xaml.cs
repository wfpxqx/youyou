using System;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CASCRS_Voucher_Import.Common.CommDB;

namespace CASCRS_Voucher_Import.Common.CommWindow
{
    /// <summary>
    /// DepartmentReference.xaml 的交互逻辑
    /// </summary>
    public partial class DepartmentReference : Window
    {
        private DataTable TreeViewDataSource = null;
        public string Value { get; set; }
        public string[] DepRefInfo { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public DepartmentReference()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 加载
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            try
            {
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.AppendLine("SELECT");
                sbSQL.AppendLine("     cDepCode");
                sbSQL.AppendLine("    ,cDepName");
                sbSQL.AppendLine("    ,iDepGrade");
                sbSQL.AppendLine("    ,bDepEnd");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    Department");
				sbSQL.AppendLine("ORDER BY");
                sbSQL.AppendLine("    cDepCode ASC");
                string strSQL = sbSQL.ToString();
                SetTreeViewNode(strSQL);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 设置树状图节点
        /// </summary>
        /// <param name="DataSourceSQL">SQL数据源</param>
        private void SetTreeViewNode(string DataSourceSQL)
        {
            try
            {
                if (TreeViewDataSource != null) TreeViewDataSource = null;
                if (trvDepRef.Items.Count > 0) trvDepRef.Items.Clear();
                TreeViewDataSource = DbOperation.GetDataTable(DataSourceSQL, 1);
                int RowCount = TreeViewDataSource.Rows.Count;
                if (RowCount == 0)
                {
                    MessageBox.Show("未能找到部门信息。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                string strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '05'", 2).Rows[0].ItemArray[0].ToString();
                for (int i = 0; i < RowCount; i++)
                {
                    string strDepCode = TreeViewDataSource.Rows[i].ItemArray[0].ToString();
                    string strDepName = TreeViewDataSource.Rows[i].ItemArray[1].ToString();
                    TreeViewItem tviDep = new TreeViewItem() { Name = "TVI" + strDepCode, Header = string.Format("({0}){1}", strDepCode, strDepName) };
                    int intDepGrade = Convert.ToInt32(TreeViewDataSource.Rows[i].ItemArray[2].ToString());
                    if (intDepGrade == 1)
                    {
                        if (Value != null)
                        {
                            int len = Convert.ToInt32(strCvalue.Substring(0, 1));
                            string strDeptLv1 = Value.Substring(0, len);
                            if (strDepCode.Equals(strDeptLv1)) tviDep.IsExpanded = true;
                            else tviDep.IsExpanded = false;
                            if (strDepCode.Equals(Value)) tviDep.IsSelected = true;
                        }
                        RegisterName("TVI" + strDepCode, tviDep); 
                        trvDepRef.Items.Add(tviDep);
                         
                    }
                    else
                    {
                        
                        int len = 0;
                        int j = 1; 
                        while (j < intDepGrade)
                        {
                            len = len + Convert.ToInt32(strCvalue.Substring(j - 1, 1));
                            j++;
                        }
                        string strParentDepCode = strDepCode.Substring(0, len);
                        object objParentDep = trvDepRef.FindName("TVI" + strParentDepCode);
                        if (objParentDep != null)
                        {
                            if (Value != null)
                            {
                                if (Value.Length > strDepCode.Length)
                                {
                                    if (strDepCode.Equals(Value.Substring(0, strDepCode.Length)))
                                    {
                                        tviDep.IsExpanded = true;
                                    }
                                    else
                                    {
                                        tviDep.IsExpanded = false;
                                    }
                                }
                                else if (Value.Length == strDepCode.Length)
                                {
                                    if (strDepCode.Equals(Value))
                                    {
                                        tviDep.IsSelected = true;
                                    }
                                }
                            }
                            RegisterName("TVI" + strDepCode, tviDep);
                            TreeViewItem triParentDep = (TreeViewItem)objParentDep;
                            triParentDep.Items.Add(tviDep);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 选中按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnSelected_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object objCurrentItem = trvDepRef.SelectedItem;
                if (objCurrentItem == null)
                    return;
                TreeViewItem tviCurrentItem = (TreeViewItem)objCurrentItem;
                int ItemCount = tviCurrentItem.Items.Count;
                if (ItemCount > 0)
                    return;
                string strHeader = tviCurrentItem.Header.ToString();
                DepRefInfo = strHeader.Remove(0, 1).Split(')');
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 部门引用鼠标双击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void trvDepRef_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                btnSelected_Click(new object(), new RoutedEventArgs());
                
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 关闭按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
