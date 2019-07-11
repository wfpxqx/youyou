using System;
using System.Configuration;
using System.Data.SqlClient;
using System.Text;
using System.Windows;
using CASCRS_Voucher_Import.Common.CommDB;

namespace CASCRS_Voucher_Import
{
	/// <summary>
	/// SystemConfig.xaml 的交互逻辑
	/// </summary>
	public partial class SystemConfig : Window
	{
		public bool ConfigCancelled { get; set; }
		public SystemConfig()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				ConfigCancelled = true;
				txtServer.Text = ConfigurationManager.AppSettings["SERVER_NAME"].ToString();
				txtDbPassword.Text = ConfigurationManager.AppSettings["DB_PASSWORD"].ToString();
				txtMiddleAccId.Text = ConfigurationManager.AppSettings["MIDDLE_ACC_ID"].ToString();
				txtMiddleAccYear.Text = ConfigurationManager.AppSettings["MIDDLE_ACC_Year"].ToString();
				txtTargetAccId.Text = ConfigurationManager.AppSettings["TARGET_ACC_ID"].ToString();
				txtTargetAccYear.Text = ConfigurationManager.AppSettings["TARGET_ACC_Year"].ToString();
				txtDbName.Text = ConfigurationManager.AppSettings["DB_NAME"].ToString();
				if(string.IsNullOrWhiteSpace(txtDbName.Text))
				{
					txtDbName.Text = "CASCRS_VOUCHER";
				}
				txtServer.Focus();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnConfirm_Click(object sender, RoutedEventArgs e)
		{
			SqlConnection objSqlConn = null;
			SqlConnection objSqlConn1 = null;
			SqlConnection objSqlConn2 = null;
			try
			{
				string strServer = txtServer.Text;
				string strDbName = txtDbName.Text;
				string strDbUserPassword = txtDbPassword.Text;
				string strMiddleAccId = txtMiddleAccId.Text;
				string strMiddleAccYear = txtMiddleAccYear.Text;
				string strTargetAccId = txtTargetAccId.Text;
				string strTargetAccYear = txtTargetAccYear.Text;
				if (strServer == string.Empty)
				{
					MessageBox.Show("服务器名不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				if (strMiddleAccId == string.Empty)
				{
					MessageBox.Show("中间账套编号不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				if (strMiddleAccYear == string.Empty)
				{
					MessageBox.Show("中间账套年份不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				if (strTargetAccId == string.Empty)
				{
					MessageBox.Show("目标账套编号不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				if (strTargetAccYear == string.Empty)
				{
					MessageBox.Show("目标账套年份不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				if (strDbName == string.Empty)
				{
					MessageBox.Show("数据库名不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				StringBuilder sbConnectionStrings = new StringBuilder();
				sbConnectionStrings.Append(string.Format("Data Source={0};", strServer));
				sbConnectionStrings.Append(string.Format("Initial Catalog={0};", strDbName));
				sbConnectionStrings.Append("Integrated Security=False;");
				sbConnectionStrings.Append("User ID=sa;");
				sbConnectionStrings.Append(string.Format("Password={0};", strDbUserPassword));
				sbConnectionStrings.Append("Connect Timeout=15;Encrypt=False;TrustServerCertificate=False");
				string ConnectionStrings = sbConnectionStrings.ToString();
				string ConnectionStrings1 = sbConnectionStrings.ToString().Replace(strDbName, string.Format("UFDATA_{0}_{1}", strMiddleAccId, strMiddleAccYear));
				string ConnectionStrings2 = sbConnectionStrings.ToString().Replace(strDbName, string.Format("UFDATA_{0}_{1}", strTargetAccId, strTargetAccYear));
				objSqlConn = new SqlConnection(ConnectionStrings);
				objSqlConn.Open();
				objSqlConn1 = new SqlConnection(ConnectionStrings1);
				objSqlConn1.Open();
				objSqlConn2 = new SqlConnection(ConnectionStrings2);
				objSqlConn2.Open();
				Configuration objConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
				objConfig.AppSettings.Settings["CASCRS_DB_CONN"].Value = ConnectionStrings;
				objConfig.AppSettings.Settings["CASCRS_DB_U8_CONN1"].Value = ConnectionStrings1;
				objConfig.AppSettings.Settings["CASCRS_DB_U8_CONN2"].Value = ConnectionStrings2;
				objConfig.AppSettings.Settings["SERVER_NAME"].Value = strServer;
				objConfig.AppSettings.Settings["DB_PASSWORD"].Value = strDbUserPassword;
				objConfig.AppSettings.Settings["DB_NAME"].Value = strDbName;
				objConfig.AppSettings.Settings["MIDDLE_ACC_YEAR"].Value = strMiddleAccYear;
				objConfig.AppSettings.Settings["MIDDLE_ACC_ID"].Value = strMiddleAccId;
				objConfig.AppSettings.Settings["TARGET_ACC_YEAR"].Value = strTargetAccYear;
				objConfig.AppSettings.Settings["TARGET_ACC_ID"].Value = strTargetAccId;
				objConfig.AppSettings.Settings["FIRST_RUN"].Value = "0";
				objConfig.Save();
				ConfigurationManager.RefreshSection("appSettings");
				DbOperation.connectionString = ConnectionStrings;
				DbOperation.connectionString1 = ConnectionStrings1;
				DbOperation.connectionString2 = ConnectionStrings2;
				MessageBox.Show("配置成功。", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
				ConfigCancelled = false;
				Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
			finally
			{
				if (objSqlConn != null)
				{
					objSqlConn.Close();
					objSqlConn.Dispose();
					objSqlConn = null;
				}
				if (objSqlConn1 != null)
				{
					objSqlConn1.Close();
					objSqlConn1.Dispose();
					objSqlConn1 = null;
				}
				if (objSqlConn2 != null)
				{
					objSqlConn2.Close();
					objSqlConn2.Dispose();
					objSqlConn2 = null;
				}
			}
		}

		private void btnCancel_Click(object sender, RoutedEventArgs e)
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
