using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using CASCRS_Voucher_Import.Common.CommDB;

namespace CASCRS_Voucher_Import
{
	/// <summary>
	/// SystemUser.xaml 的交互逻辑
	/// </summary>
	public partial class SystemUser : Window
	{
		private DataTable SUDataSource = null;
		private SqlDataAdapter sdaSU = null;
		public SystemUser()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				SUDataSource = DbOperation.GetDataTable(out sdaSU, "SELECT * FROM SystemUser");
				gcSystemUser.ItemsSource = SUDataSource;
				gcSystemUser.Columns[0].Header = "用户ID";
				gcSystemUser.Columns[1].Header = "用户显示名";
				gcSystemUser.Columns[2].Header = "用户密码";
				DevExpress.Xpf.Editors.Settings.PasswordBoxEditSettings pbs = new DevExpress.Xpf.Editors.Settings.PasswordBoxEditSettings();
				gcSystemUser.Columns[2].EditSettings = pbs;
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnAdd_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if(SUDataSource.Rows.Count > 0)
				{
					DataRow drLast = SUDataSource.Rows[SUDataSource.Rows.Count - 1];
					if (drLast.RowState == DataRowState.Deleted)
					{
						DataRow dr2 = SUDataSource.NewRow();
						dr2[0] = string.Empty;
						SUDataSource.Rows.Add(dr2);
						return;
					}
					string strUserId = drLast[0].ToString();
					if (string.IsNullOrWhiteSpace(strUserId))
					{
						MessageBox.Show("用户ID不能为空", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
					DataRow[] drcCurrentRows = SUDataSource.Select(string.Format("userId = '{0}'", strUserId), null, DataViewRowState.CurrentRows);
					if (drcCurrentRows.Length > 1)
					{
						MessageBox.Show("用户ID不能重复。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
				}
				DataRow dr = SUDataSource.NewRow();
				dr[0] = string.Empty;
				SUDataSource.Rows.Add(dr);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnDelete_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				object objCurrentItem = gcSystemUser.CurrentItem;
				if (objCurrentItem == null)
					return;
				DataRow dr = ((DataRowView)objCurrentItem).Row;
				string strUserId = dr[0].ToString();
				if (strUserId == "admin")
				{
					MessageBox.Show("anmin帐户不能删除", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				dr.Delete();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnSave_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (SUDataSource.Rows.Count > 0)
				{
					DataRow drLast = SUDataSource.Rows[SUDataSource.Rows.Count - 1];
					if (drLast.RowState == DataRowState.Deleted)
					{
						DbOperation.UpdateDataSource(sdaSU, SUDataSource);
						MessageBox.Show("保存完成。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
						return;
					}
					string strUserId = drLast[0].ToString();
					if (string.IsNullOrWhiteSpace(strUserId))
					{
						MessageBox.Show("用户ID不能为空", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
					DataRow[] drcCurrentRows = SUDataSource.Select(string.Format("userId = '{0}'", strUserId), null, DataViewRowState.CurrentRows);
					if (drcCurrentRows.Length > 1)
					{
						MessageBox.Show("用户ID不能重复。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
						return;
					}
				}
				DbOperation.UpdateDataSource(sdaSU, SUDataSource);
				MessageBox.Show("保存完成。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

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

		private void gcSystemUser_CurrentItemChanged(object sender, DevExpress.Xpf.Grid.CurrentItemChangedEventArgs e)
		{
			try
			{
				object objCurrentItem = gcSystemUser.CurrentItem;
				if (objCurrentItem == null)
					return;
				DataRow dr = ((DataRowView)objCurrentItem).Row;
				string strUserId = dr[0].ToString();
				if (strUserId == "admin")
				{
					gcSystemUser.View.AllowEditing = false;
				}
				else
				{
					gcSystemUser.View.AllowEditing = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
	}
}
