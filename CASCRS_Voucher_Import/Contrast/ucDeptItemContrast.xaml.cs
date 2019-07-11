using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CASCRS_Voucher_Import.Common.CommDB;
using CASCRS_Voucher_Import.Common.CommWindow;
using DevExpress.Xpf.Editors.Settings;
using DevExpress.Xpf.Grid;
using Microsoft.ApplicationInsights;

namespace CASCRS_Voucher_Import.Contrast
{
	/// <summary>
	/// ucDeptItemContrast.xaml 的交互逻辑
	/// </summary>
	public partial class ucDeptItemContrast : UserControl
	{
		private TelemetryClient tc = new TelemetryClient();
		private ButtonEditSettings btnRef = null;
		private int LOGIN_YEAR = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Year;
		private int LOGIN_MONTH = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Month;
		//private string MIDDLE_ACC_ID = ConfigurationManager.AppSettings["MIDDLE_ACC_ID"];
		//private string MIDDLE_ACC_YEAR = ConfigurationManager.AppSettings["MIDDLE_ACC_YEAR"];
		//private string TARGET_ACC_ID = ConfigurationManager.AppSettings["TARGET_ACC_ID"];
		//private string TARGET_ACC_YEAR = ConfigurationManager.AppSettings["TARGET_ACC_YEAR"];
		private SqlDataAdapter sdaDeptItemCst = null;
		private DataTable DeptItemCstDataSource = null;
		private DataTable dtDepartment = null;
		private DataTable dtFitem = null;
		private DataTable dtFitemss = null;
		public ucDeptItemContrast()
		{
			InitializeComponent();
		}

		private void UserControl_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				//ApplicationInsights跟踪
				tc.InstrumentationKey = "30b22ad1-3cdd-462f-885f-fce870bb83c4";
				tc.Context.User.Id = Environment.UserName;
				tc.Context.Session.Id = Guid.NewGuid().ToString();
				tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
				tc.TrackPageView("中科腐蚀研究院凭证导入工具");

				DeptItemCstDataSource = DbOperation.GetDataTable(out sdaDeptItemCst, "SELECT * FROM DeptItemContrast");
				SetDataTable();
				SetDeptItemContrastGrid(DeptItemCstDataSource);
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void SetDataTable()
		{
			try
			{
				dtDepartment = DbOperation.GetDataTable("SELECT cDepCode, cDepName FROM Department", 1);
				StringBuilder sbSQL = new StringBuilder();
				sbSQL.AppendLine("SELECT");
				sbSQL.AppendLine("     citem_class");
				sbSQL.AppendLine("    ,citem_name");
				sbSQL.AppendLine("    ,ctable");
				sbSQL.AppendLine("    ,cClasstable");
				sbSQL.AppendLine("FROM");
				sbSQL.AppendLine("    fitem");
				string strSQL = sbSQL.ToString();
				dtFitem = DbOperation.GetDataTable(strSQL, 2);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private void SetDeptItemContrastGrid(DataTable dtDeptItemCst)
		{
			try
			{
				gcDeptItemContrast.ItemsSource = dtDeptItemCst;
				gcDeptItemContrast.Columns[0].Header = "唯一标识";
				gcDeptItemContrast.Columns[1].Header = "部门编码";
				gcDeptItemContrast.Columns[2].Header = "部门名称";
				gcDeptItemContrast.Columns[3].Header = "项目大类编码";
				gcDeptItemContrast.Columns[4].Header = "项目大类名称";
				gcDeptItemContrast.Columns[5].Header = "项目编码";
				gcDeptItemContrast.Columns[6].Header = "项目名称";
				if (btnRef == null)
				{
					btnRef = new ButtonEditSettings() { ShowNullText = false, ShowText = true };
					btnRef.DefaultButtonClick += BtnRef_DefaultButtonClick;
				}
				gcDeptItemContrast.Columns[0].Visible = false;
				gcDeptItemContrast.Columns[1].EditSettings = btnRef;
				gcDeptItemContrast.Columns[3].EditSettings = btnRef;
				gcDeptItemContrast.Columns[5].EditSettings = btnRef;
				gcDeptItemContrast.Columns[2].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
				gcDeptItemContrast.Columns[4].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
				gcDeptItemContrast.Columns[6].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
				tvDeptItemContrast.BestFitMode = DevExpress.Xpf.Core.BestFitMode.AllRows;
				tvDeptItemContrast.BestFitColumns();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private void BtnRef_DefaultButtonClick(object sender, RoutedEventArgs e)
		{
			try
			{
				object objCurrentItem = gcDeptItemContrast.CurrentItem;
				if (objCurrentItem == null)
					return;
				ColumnBase cb = gcDeptItemContrast.CurrentColumn;
				if (cb == null)
					return;
				tvDeptItemContrast.CancelRowEdit();
				DataRow dr = ((DataRowView)objCurrentItem).Row;
				string strHeader = cb.Header.ToString();
				if (strHeader == "部门编码")
				{
					DepartmentReference win = new DepartmentReference();
					string strValue = dr["deptId"].ToString();
					if (!string.Empty.Equals(strValue))
						win.Value = strValue;
					win.ShowDialog();
					string[] strDepRefInfo = win.DepRefInfo;
					if (strDepRefInfo == null)
						return;
					dr.BeginEdit();
					dr["deptId"] = strDepRefInfo[0];
					dr["deptName"] = strDepRefInfo[1];
					dr.EndEdit();
				}
				else if (strHeader == "项目大类编码")
				{
					ReferenceWindow win = new ReferenceWindow() { BaseTitle = "项目大类参照", DataSource = dtFitem, GridHeader = new string[] { "项目大类码", "项目大类名", "项目表名", "项目大类表名" } };
					win.ShowDialog();
					DataRow drResult = win.Result;
					if (drResult == null)
						return;
					dr.BeginEdit();
					dr["itemCClass"] = drResult[0];
					dr["itemCName"] = drResult[1];
					dr.EndEdit();
				}
				else if(strHeader == "项目编码")
				{
					string strProjectBigClass = dr["itemCClass"].ToString();
					if (string.Empty.Equals(strProjectBigClass))
						return;
					DataRow drFitem = dtFitem.Select(string.Format("citem_class = '{0}'", strProjectBigClass))[0];
					string strCtable = drFitem[2].ToString();
					string strCclasstable = drFitem[3].ToString();
					int RowCount = DbOperation.GetDataTable(string.Format("select * from sysobjects where id = object_id('{0}') and type = 'u'", strCclasstable), 2).Rows.Count;
					int RowCount2 = DbOperation.GetDataTable(string.Format("select * from sysobjects where id = object_id('{0}') and type = 'u'", strCtable), 2).Rows.Count;
					StringBuilder sbSQL = new StringBuilder();
					string strSQL = null;
					DataTable dtCtable = null;
					DataTable dtCclasstable = null;
					if (RowCount > 0)
					{
						sbSQL.Clear();
						sbSQL.AppendLine("SELECT");
						sbSQL.AppendLine("     cItemCcode");
						sbSQL.AppendLine("    ,cItemCname");
						sbSQL.AppendLine("    ,iItemCgrade");
						sbSQL.AppendLine("    ,bItemCend");
						sbSQL.AppendLine("FROM");
						sbSQL.AppendLine(string.Format("    {0}", strCclasstable));
						strSQL = sbSQL.ToString();
						dtCclasstable = DbOperation.GetDataTable(strSQL, 2);
					}
					if (RowCount2 > 0)
					{
						sbSQL.Clear();
						sbSQL.AppendLine("SELECT");
						sbSQL.AppendLine("     citemcode");
						sbSQL.AppendLine("    ,citemname");
						sbSQL.AppendLine("    ,citemccode");
						sbSQL.AppendLine("FROM");
						sbSQL.AppendLine(string.Format("    {0}", strCtable));
						strSQL = sbSQL.ToString();
						dtCtable = DbOperation.GetDataTable(strSQL, 2);
					}
					ReferenceWindow2 win = new ReferenceWindow2() { BaseTitle = "项目参照", TreeViewDataSource = dtCclasstable, GridDataSource = dtCtable, GridHeader = new string[] { "项目编码", "项目名称", "项目大类编码" }, ProjBigClass = strProjectBigClass };
					string strValue = dr["itemId"].ToString();
					if (!string.Empty.Equals(strValue))
						win.Value = strValue;
					win.ShowDialog();
					DataRow drResult = win.SelectedRow;
					if (drResult == null)
						return;
					dr.BeginEdit();
					dr["itemId"] = drResult[0];
					dr["itemName"] = drResult[1];
					dr.EndEdit();
				}
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnAdd_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				tvDeptItemContrast.CancelRowEdit();
				int RowCount = DeptItemCstDataSource.Rows.Count;
				if (RowCount > 0)
				{
					DataRow drLast = DeptItemCstDataSource.Rows[RowCount - 1];
					string strDeptId = drLast["deptId"].ToString();
					if (string.IsNullOrWhiteSpace(strDeptId))
					{
						MessageBox.Show("部门编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
					string strItemCClass = drLast["itemCClass"].ToString();
					if (string.IsNullOrWhiteSpace(strItemCClass))
					{
						MessageBox.Show("项目大类编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
					string strItemId = drLast["itemId"].ToString();
					if (string.IsNullOrWhiteSpace(strItemCClass))
					{
						MessageBox.Show("项目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
					DataRow[] drc = DeptItemCstDataSource.Select(string.Format("deptId = '{0}' AND itemCClass = '{1}' AND itemId = '{2}'", strDeptId, strItemCClass, strItemId));
					if (drc.Length > 1)
					{
						MessageBox.Show(string.Format("部门编码{0}项目大类编码{1}项目编码{2}的行存在若干重复行", strDeptId, strItemCClass, strItemId), "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
				}
				DataRow dr = DeptItemCstDataSource.NewRow();
				dr[0] = Guid.NewGuid();
				DeptItemCstDataSource.Rows.Add(dr);
				SaveData();
				tvDeptItemContrast.MoveLastRow();
				gcDeptItemContrast.RefreshData();
				gcDeptItemContrast.Focus();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void SaveData()
		{
			try
			{
				DbOperation.UpdateDataSource(sdaDeptItemCst, DeptItemCstDataSource);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private void btnDelete_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				object objCurrentItem = gcDeptItemContrast.CurrentItem;
				if (objCurrentItem == null)
					return;
				MessageBoxResult mbrResult = MessageBox.Show("是否要删除该行数据？", "信息", MessageBoxButton.YesNo, MessageBoxImage.Information);
				if (mbrResult == MessageBoxResult.No)
					return;
				DataRow dr = ((DataRowView)objCurrentItem).Row;
				dr.Delete();
				SaveData();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void gcDeptItemContrast_CurrentItemChanged(object sender, CurrentItemChangedEventArgs e)
		{
			try
			{
				tvDeptItemContrast.CancelRowEdit();
				if (e == null)
					return;
				if (e.OldItem == null)
					return;
				DataRow dr = ((DataRowView)e.OldItem).Row;
				if (dr.RowState == DataRowState.Deleted || dr.RowState == DataRowState.Detached)
				{
					SaveData();
					return;
				}
				string strDeptId = dr["deptId"].ToString();
				if (string.IsNullOrWhiteSpace(strDeptId))
				{
					MessageBox.Show("部门编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				string strItemCClass = dr["itemCClass"].ToString();
				if (string.IsNullOrWhiteSpace(strItemCClass))
				{
					MessageBox.Show("项目大类编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				string strItemId = dr["itemId"].ToString();
				if (string.IsNullOrWhiteSpace(strItemCClass))
				{
					MessageBox.Show("项目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				DataRow[] drc = DeptItemCstDataSource.Select(string.Format("deptId = '{0}' AND itemCClass = '{1}' AND itemId = '{2}'", strDeptId, strItemCClass, strItemId));
				if (drc.Length > 1)
				{
					MessageBox.Show(string.Format("部门编码{0}项目大类编码{1}项目编码{2}的行存在若干重复行", strDeptId, strItemCClass, strItemId), "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				SaveData();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void tvDeptItemContrast_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
		{
			try
			{
				if (DeptItemCstDataSource == null)
					return;
				if (DeptItemCstDataSource.Rows.Count != 1)
					return;
				SaveData();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void tvDeptItemContrast_CellValueChanged(object sender, CellValueChangedEventArgs e)
		{
			try
			{
				if (e == null)
					return;
				if (e.Row == null)
					return;
				if (e.Column == null)
					return;
				DataRow dr = ((DataRowView)e.Row).Row;
				dr.EndEdit();
				GridColumn gc = e.Column;
				string strHeader = gc.Header.ToString();
				string strFieldName = gc.FieldName;
				if (strHeader == "部门编码")
				{
					DataRow[] drc = dtDepartment.Select(string.Format("cDepCode = '{0}'", e.Value));
					dr.BeginEdit();
					if (drc.Length == 0)
						dr["deptName"] = DBNull.Value;
					else
						dr["deptName"] = drc[0][1];
					dr.EndEdit();
				}
				else if (strHeader == "项目大类编码")
				{
					DataRow[] drc = dtFitem.Select(string.Format("citem_class = '{0}'", e.Value));
					dr.BeginEdit();
					if (drc.Length == 0)
					{
						dr["itemCName"] = DBNull.Value;
						dr["itemId"] = DBNull.Value;
						dr["itemName"] = DBNull.Value;
					}	
					else
						dr["itemCName"] = drc[0][1];
					dr.EndEdit();
				}
				else if (strHeader == "项目编码")
				{
					if (dtFitemss == null)
					{
						string strItemCClass = ((DataRowView)e.Row).Row["itemCClass"].ToString();
						DataRow[] drcFitem = dtFitem.Select(string.Format("citem_class = '{0}'", strItemCClass));
						if (drcFitem.Length == 0)
							return;
						DataRow drFitem = drcFitem[0];
						string strCtable = drFitem[2].ToString();
						int RowCount2 = DbOperation.GetDataTable(string.Format("select * from sysobjects where id = object_id('{0}') and type = 'u'", strCtable), 2).Rows.Count;
						if (RowCount2 == 0)
							return;
						StringBuilder sbSQL = new StringBuilder();
						sbSQL.Clear();
						sbSQL.AppendLine("SELECT");
						sbSQL.AppendLine("     citemcode");
						sbSQL.AppendLine("    ,citemname");
						sbSQL.AppendLine("    ,citemccode");
						sbSQL.AppendLine("FROM");
						sbSQL.AppendLine(string.Format("    {0}", strCtable));
						string strSQL = sbSQL.ToString();
						dtFitemss = DbOperation.GetDataTable(strSQL, 2);
					}
					DataRow[] drc = dtFitemss.Select(string.Format("citemcode = '{0}'", e.Value));
					dr.BeginEdit();
					if (drc.Length == 0)
					{
						dr["itemName"] = DBNull.Value;
					}
					else
					{
						dr["itemName"] = drc[0][1];
					}
					dr.EndEdit();
				}
				if (DeptItemCstDataSource == null)
					return;
				if (DeptItemCstDataSource.Rows.Count != 1)
					return;
				SaveData();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
	}
}
