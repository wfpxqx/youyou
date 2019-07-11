using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using DevExpress.Xpf.Grid;
using CASCRS_Voucher_Import.Common.CommDB;

namespace CASCRS_Voucher_Import.Common.CommWindow
{
    /// <summary>
    /// ReferenceWindow2.xaml 的交互逻辑
    /// </summary>
    public partial class ReferenceWindow2 : Window
    {
        private bool blBvalue = false;
        private string strCvalue = null;
        private string ValueClass = null;
		private ObservableCollection<GridColumn> Columns = null;
		public string BaseTitle { get; set; }
        public DataTable TreeViewDataSource { get; set; }
        public DataTable GridDataSource { get; set; }
        public string[] GridHeader { get; set; }
        public string ProjBigClass { get; set; }
        public DataRow SelectedRow { get; set; }
        public string Value { get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public ReferenceWindow2()
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
                Title = BaseTitle;
                DataRow[] drc = null;
                int count = GridDataSource.Columns.Count;
                if (!string.IsNullOrWhiteSpace(Value))
                {
					drc = GridDataSource.Select(string.Format("{0} = '{1}'", GridDataSource.Columns[0].ColumnName, Value));
					if (!"会计科目参照".Equals(BaseTitle))
                    {
                        if (drc.Length > 0)
							ValueClass = drc[0][count - 1].ToString();
                    }
                }
                if ("会计科目参照".Equals(BaseTitle))
                {
                    blBvalue = true;
                    int intYear = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Year;
                    
                    //strCvalue = DbOperation.GetDataTable(string.Format("SELECT CODINGRULE FROM GradeDef_Base WHERE KEYWORD = 'code' AND iYear = '2017'"), 2).Rows[0].ItemArray[0].ToString();
                    //strCvalue = DbOperation.GetDataTable(string.Format("SELECT CODINGRULE FROM GradeDef_Base WHERE KEYWORD = 'code' AND iYear = {0}", intYear), 2).Rows[0].ItemArray[0].ToString();
                    strCvalue = DbOperation.GetDataTable(string.Format("SELECT CODINGRULE FROM GradeDef_Base WHERE KEYWORD = 'code'") , 2).Rows[0].ItemArray[0].ToString();
                }
                else if ("职员参照".Equals(BaseTitle))
                {
                    blBvalue = true;
                    strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '05'", 2).Rows[0].ItemArray[0].ToString();
                }
                else if ("客户参照".Equals(BaseTitle))
                {
                    blBvalue = Convert.ToBoolean(DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '40'", 2).Rows[0].ItemArray[0]);
                    strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '02'", 2).Rows[0].ItemArray[0].ToString();
                }
                else if ("供应商参照".Equals(BaseTitle))
                {
                    blBvalue = Convert.ToBoolean(DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '41'", 2).Rows[0].ItemArray[0]);
                    strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '04'", 2).Rows[0].ItemArray[0].ToString();
                }
                else if ("项目参照".Equals(BaseTitle))
                {
                    blBvalue = true;
                    strCvalue = DbOperation.GetDataTable(string.Format("SELECT crule FROM fitem WHERE citem_class = '{0}'", ProjBigClass), 2).Rows[0].ItemArray[0].ToString();
                }
				else if ("结算方式参照".Equals(BaseTitle))
				{
					blBvalue = true;
					strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '06'", 2).Rows[0].ItemArray[0].ToString();
				}
				else if ("存货参照".Equals(BaseTitle))
				{
					blBvalue = true;
					strCvalue = DbOperation.GetDataTable("SELECT cValue FROM AccInformation WHERE cSysID = 'AA' AND cID = '01'", 2).Rows[0].ItemArray[0].ToString();
				}
                if (blBvalue)
					SetTreeViewNode();
                if (drc != null) 
                {
                    if (drc.Length > 0)
                    {
                        if ("会计科目参照".Equals(BaseTitle))
                        {
                            DataTable dtTemp = drc.CopyToDataTable();
                            SetGridControl(dtTemp);
                        }
                        else
                        {
                            DataTable dtTemp = drc.CopyToDataTable();
                            DataRow[] drc2 = GridDataSource.Select(string.Format("{0} = '{1}' AND {2} <> '{3}'", GridDataSource.Columns[count - 1].ColumnName, ValueClass, GridDataSource.Columns[0].ColumnName, Value));
                            if (drc2.Length > 0) dtTemp.Merge(drc2.CopyToDataTable(), true, MissingSchemaAction.Ignore);
                            SetGridControl(dtTemp);
                        }
                    }
                    else
                    {
                        SetGridControl();     
                    }
                }
                else
                {
                    SetGridControl();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 设置树状图节点
        /// </summary>
        private void SetTreeViewNode()
        {
            try
            {
                if (TreeViewDataSource == null) return;
                TreeViewItem tviMain = (TreeViewItem)trvRef.Items[0];
                tviMain.Items.Clear();
                int RowCount = TreeViewDataSource.Rows.Count;
                if (RowCount == 0)
                {
                    MessageBox.Show("未能找到相关信息。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
				if ("会计科目参照".Equals(BaseTitle))
				{
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "Asset", Header = "资产" });
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "Debt", Header = "负债" });
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "Common", Header = "共同" });
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "Benefit", Header = "权益" });
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "Cost", Header = "成本" });
					tviMain.Items.Add(new TreeViewItem() { IsExpanded = true, Name = "PR", Header = "损益" });
				}
				for (int i = 0; i < RowCount; i++)
                {
                    string strCode = TreeViewDataSource.Rows[i].ItemArray[0].ToString();
                    string strName = TreeViewDataSource.Rows[i].ItemArray[1].ToString();
                    TreeViewItem tviRef = new TreeViewItem() { Name = "TVI" + strCode, Header = string.Format("({0}){1}", strCode, strName), IsExpanded = false};
                    int intGrade = Convert.ToInt32(TreeViewDataSource.Rows[i].ItemArray[2].ToString());
                    if (intGrade == 1)
                    {
                        if ("会计科目参照".Equals(BaseTitle))
                        {
                            if (Value != null)
                            {
                                int len = Convert.ToInt32(strCvalue.Substring(0, 1));
                                string strCodeLv1 = Value.Substring(0, len);
                                if (strCode.Equals(strCodeLv1)) tviRef.IsExpanded = true;
                                else tviRef.IsExpanded = false;
                                if (strCode.Equals(Value)) tviRef.IsSelected = true;
                            }
                        }
                        else
                        {
                            if (ValueClass != null)
                            {
                                int len = Convert.ToInt32(strCvalue.Substring(0, 1));
                                string strCodeLv1 = ValueClass.Substring(0, len);
                                if (strCode.Equals(strCodeLv1)) tviRef.IsExpanded = true;
                                else tviRef.IsExpanded = false;
                                if (strCode.Equals(ValueClass)) tviRef.IsSelected = true;
                            }
                        }
                        
                        RegisterName("TVI" + strCode, tviRef); 
                        if ("会计科目参照".Equals(BaseTitle))
                        {
                            string strCClass = TreeViewDataSource.Rows[i].ItemArray[4].ToString();
                            foreach (TreeViewItem tvi in tviMain.Items)
                            {
                                if (strCClass.Equals(tvi.Header))
                                {
                                    tvi.Items.Add(tviRef);
                                    break;
                                }
                            }
                        }
                        else tviMain.Items.Add(tviRef);
                    }
                    else
                    {
                        int len = 0;
                        int j = 1;
                        while (j < intGrade)
                        {
                            len = len + Convert.ToInt32(strCvalue.Substring(j - 1, 1));
                            j++;
                        }
                        string strParentCode = strCode.Substring(0, len);
                        object objParentDep = trvRef.FindName("TVI" + strParentCode);
                        if (objParentDep != null)
                        {
                            if ("会计科目参照".Equals(BaseTitle))
                            {
                                if (Value != null)
                                {
                                    if (Value.Length > strCode.Length)
                                    {
                                        if (strCode.Equals(Value.Substring(0, strCode.Length)))
                                        {
                                            tviRef.IsExpanded = true;
                                        }
                                        else
                                        {
                                            tviRef.IsExpanded = false;
                                        }
                                    }
                                    else if (Value.Length == strCode.Length)
                                    {
                                        if (strCode.Equals(Value))
                                        {
                                            tviRef.IsSelected = true;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (ValueClass != null)
                                {
                                    if (ValueClass.Length > strCode.Length)
                                    {
                                        if (strCode.Equals(ValueClass.Substring(0, strCode.Length)))
                                        {
                                            tviRef.IsExpanded = true;
                                        }
                                        else
                                        {
                                            tviRef.IsExpanded = false;
                                        }
                                    }
                                    else if (ValueClass.Length == strCode.Length)
                                    {
                                        if (strCode.Equals(ValueClass))
                                        {
                                            tviRef.IsSelected = true;
                                        }
                                    }
                                }
                            }
                            RegisterName("TVI" + strCode, tviRef);
                            TreeViewItem triParentDep = (TreeViewItem)objParentDep;
                            triParentDep.Items.Add(tviRef);
                        }
						else
						{
							if ("会计科目参照".Equals(BaseTitle))
							{
								string strCClass = TreeViewDataSource.Rows[i].ItemArray[4].ToString();
								foreach (TreeViewItem tvi in tviMain.Items)
								{
									if (strCClass.Equals(tvi.Header))
									{
										tvi.Items.Add(tviRef);
										break;
									}
								}
							}
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
        /// 设置Grid
        /// </summary>
        /// <param name="dtTemp">临时数据源</param>
        private void SetGridControl(DataTable dtTemp = null)
        {
            try
            {
                if (GridDataSource == null) return;
                int RowCount = GridDataSource.Rows.Count;
                if (RowCount == 0)
                {
                    MessageBox.Show("未能找到相关信息。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (dtTemp != null) gcRef.ItemsSource = dtTemp;
                else gcRef.ItemsSource = GridDataSource;
                int HeaderCount = GridDataSource.Columns.Count;
                for (int i = 0; i < HeaderCount; i++)
                {
                    gcRef.Columns[i].Header = GridHeader[i];
                    gcRef.Columns[i].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
                    gcRef.Columns[i].ReadOnly = true;
                }
                gcRef.SelectionMode = MultiSelectMode.Row;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnSelected_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object objSelectedRow = gcRef.CurrentItem;
                if (objSelectedRow == null)
                    return;
                SelectedRow = ((DataRowView)objSelectedRow).Row;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 刷新按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                trvRef.Items.Refresh();
                gcRef.RefreshData();
            }
            catch (Exception ex)
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

        /// <summary>
        /// Grid鼠标双击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void gcRef_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                btnSelected_Click(new object(), new RoutedEventArgs());    
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 树状图鼠标双击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void trvRef_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
			try
			{
				object objSelectedItem = trvRef.SelectedItem;
				if (objSelectedItem == null)
					return;
				TreeViewItem tviSelectedItem = (TreeViewItem)objSelectedItem;
				string Header = tviSelectedItem.Header.ToString();
				if ("会计科目参照".Equals(BaseTitle))
				{
					int ItemCount = tviSelectedItem.Items.Count;
					if (ItemCount > 0)
						return;
					string strCode = Header.Remove(0, 1).Split(')')[0];
					DataRow[] drc = GridDataSource.Select(string.Format("ccode = '{0}'", strCode));
					if (drc.Length == 0)
						return;
					SelectedRow = drc[0];
					Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}    
        }

        /// <summary>
        /// 搜索按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
				Columns = new ObservableCollection<GridColumn>();
				foreach (GridColumn gc in gcRef.Columns)
				{
					if (gc.FieldName == "autoId")
						continue;
					Columns.Add(gc);
				}
				SearchWindow winSearch = new SearchWindow() { Columns = Columns };
				winSearch.ShowDialog();
				string strCondition = winSearch.strResult;
				DataRow[] drc = GridDataSource.Select(strCondition);
				if (drc.Length > 0)
					SetGridControl(drc.CopyToDataTable());
				else
					SetGridControl(GridDataSource.Clone());
			}
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

		private void trvRef_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
		{
			try
			{
				object objSelectedItem = trvRef.SelectedItem;
				if (objSelectedItem == null)
					return;
				TreeViewItem tviSelectedItem = (TreeViewItem)objSelectedItem;
				string Header = tviSelectedItem.Header.ToString();
				if ("全部分类".Equals(Header))
				{
					gcRef.ItemsSource = GridDataSource;
				}
				else
				{
					if ("会计科目参照".Equals(BaseTitle))
					{
						TreeViewItem tviParent = (TreeViewItem)tviSelectedItem.Parent;
						if ("全部分类".Equals(tviParent.Header.ToString()))
						{
							string strCClass = tviSelectedItem.Header.ToString();
							DataRow[] drcSelected = GridDataSource.Select(string.Format("cclass = '{0}'", strCClass));
							if (drcSelected.Length == 0)
							{
								DataTable dtCopy = GridDataSource.Copy();
								dtCopy.Rows.Clear();
								gcRef.ItemsSource = dtCopy;

							}
							else
							{
								gcRef.ItemsSource = drcSelected.CopyToDataTable();
							}
							int HeaderCount2 = GridDataSource.Columns.Count;
							for (int i = 0; i < HeaderCount2; i++)
							{
								gcRef.Columns[i].Header = GridHeader[i];
								gcRef.Columns[i].ReadOnly = true;
							}
							gcRef.SelectionMode = MultiSelectMode.Row;
							return;
						}
					}
					string strCode = Header.Remove(0, 1).Split(')')[0];
					DataRow[] drsSelected = null;
					if ("会计科目参照".Equals(BaseTitle))
					{
						drsSelected = GridDataSource.Select(string.Format("{0} LIKE '{1}%'", GridDataSource.Columns[0].ColumnName, strCode));
					}
					else
					{
						int ColCount = GridDataSource.Columns.Count;
						drsSelected = GridDataSource.Select(string.Format("{0} LIKE '{1}%'", GridDataSource.Columns[ColCount - 1].ColumnName, strCode));
					}
					if (drsSelected == null)
						return;
					if (drsSelected.Length == 0)
					{
						DataTable dtCopy = GridDataSource.Copy();
						dtCopy.Rows.Clear();
						gcRef.ItemsSource = dtCopy;

					}
					else
					{
						gcRef.ItemsSource = drsSelected.CopyToDataTable();
					}
				}
				int HeaderCount = GridDataSource.Columns.Count;
				for (int i = 0; i < HeaderCount; i++)
				{
					gcRef.Columns[i].Header = GridHeader[i];
					gcRef.Columns[i].ReadOnly = true;
				}
				gcRef.SelectionMode = MultiSelectMode.Row;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
	}
}
