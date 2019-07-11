using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using CASCRS_Voucher_Import.Common.CommDB;

namespace CASCRS_Voucher_Import
{
	/// <summary>
	/// Login.xaml 的交互逻辑
	/// </summary>
	public partial class Login : Window
	{
		public bool LoginCancelled { get; set; }
		public Login()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				LoginCancelled = true;
				dteLoginDate.DateTime = DateTime.Now;
				txtLoginId.Focus();

				//测试用
				//txtLoginId.Text = "admin";
				//txtLoginPassword.Text = "admin";
				//dteLoginDate.DateTime = Convert.ToDateTime("2017-9-10");
				//btnConfirm_Click(new object(), new RoutedEventArgs());
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnConfirm_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				string strLoginId = txtLoginId.Text;
				string strLoginPassword = txtLoginPassword.Text;
				if (string.IsNullOrWhiteSpace(strLoginId))
				{
					MessageBox.Show("用户名不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				DataTable dtLogin = DbOperation.GetDataTable(string.Format("SELECT * FROM SystemUser WHERE userId = '{0}' AND userPassword = '{1}'", strLoginId, strLoginPassword));
				int count = dtLogin.Rows.Count;
				if (count == 0)
				{
					MessageBox.Show("用户名或密码错误，请重新输入。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					return;
				}
				string LoginDate = dteLoginDate.DateTime.ToString("yyyy-MM-dd");
				Configuration objConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
				objConfig.AppSettings.Settings["USER_ID"].Value = strLoginId;
				objConfig.AppSettings.Settings["USER_NAME"].Value = dtLogin.Rows[0][1].ToString();
				objConfig.AppSettings.Settings["LOGIN_DATE"].Value = LoginDate;
				objConfig.Save();
				ConfigurationManager.RefreshSection("appSettings");
				LoginCancelled = false;
				Close();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnSystemConfig_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				SystemConfig win = new SystemConfig();
				win.ShowDialog();
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void txtLoginId_PreviewKeyUp(object sender, KeyEventArgs e)
		{
			try
			{
				if (e.Key == Key.Enter)
				{
					txtLoginPassword.Focus();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void txtLoginPassword_PreviewKeyUp(object sender, KeyEventArgs e)
		{
			try
			{
				if (e.Key == Key.Enter)
				{
					btnConfirm_Click(new object(), new RoutedEventArgs());
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
	}
}
