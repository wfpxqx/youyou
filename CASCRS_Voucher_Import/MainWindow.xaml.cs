using System;
using System.Configuration;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using CASCRS_Voucher_Import.Common.CommCheck;
using CASCRS_Voucher_Import.Contrast;
using DevExpress.Xpf.Core;
using Microsoft.ApplicationInsights;

namespace CASCRS_Voucher_Import
{
	/// <summary>
	/// MainWindow.xaml 的交互逻辑
	/// </summary>
	public partial class MainWindow : Window
	{
		private TelemetryClient tc = new TelemetryClient();
		public MainWindow()
		{
			InitializeComponent();
		}

		private void Window_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
				Visibility = Visibility.Collapsed;
				Process[] pros = Process.GetProcesses();
				Process proCurrent = Process.GetCurrentProcess();
				int n = 0;
				foreach (Process pro in pros)
					if (pro.ProcessName.Equals(proCurrent.ProcessName))
						n++;
				if (n > 1)
				{
					MessageBox.Show("本程序正在运行中，请关闭后重试。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
					Environment.Exit(0);
				}
				//ApplicationInsights跟踪
				tc.InstrumentationKey = "30b22ad1-3cdd-462f-885f-fce870bb83c4";
				tc.Context.User.Id = Environment.UserName;
				tc.Context.Session.Id = Guid.NewGuid().ToString();
				tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
				tc.TrackPageView("中科腐蚀研究院凭证导入工具");
				bool IsFirstRun = Convert.ToBoolean(Convert.ToInt32(ConfigurationManager.AppSettings["FIRST_RUN"]));
				if (IsFirstRun)
				{
					SystemConfig winCfg = new SystemConfig();
					winCfg.ShowDialog();
					bool ConfigCancelled = winCfg.ConfigCancelled;
					if (ConfigCancelled)
						Environment.Exit(0);
				}
				Login winLg = new Login();
				winLg.ShowDialog();
				bool LoginCancelled = winLg.LoginCancelled;
				if (LoginCancelled)
					Environment.Exit(0);
				Visibility = Visibility.Visible;
                MainTreeView_PreviewMouseDoubleClick(null, null);

            }
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnSystemConfig_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				SystemConfig winCfg = new SystemConfig();
				winCfg.ShowDialog();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnExit_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				Environment.Exit(0);
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnSystemUser_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				SystemUser winSU = new SystemUser();
				winSU.ShowDialog();
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
      
		private void MainTreeView_PreviewMouseDoubleClick(object sender, MouseButtonEventArgs e)
		{
            try
            {
                object objCurrentItem = MainTreeView.SelectedItem;
                if (objCurrentItem == null)
                    return;
                TreeViewItem tviCurrentItem = (TreeViewItem)objCurrentItem;
                if (tviCurrentItem.Items.Count > 0)
                    return;
                string strHeader = tviCurrentItem.Header.ToString();
                DXTabItem tabItem = null;
                bool Visible = UICheck.CheckDXTabControlVisible(MainTabControl);
                if (!Visible)
                    MainTabControl.Visibility = Visibility.Visible;
                bool HasItem = UICheck.CheckDXTabControlRepeatItem(strHeader, MainTabControl, out tabItem);
                if (HasItem)
                {
                    //tabItem.IsSelected = true;
                    //tabItem.Focus();
                    //return;
                    MainTabControl.Items.Remove(tabItem);
                }
                switch (strHeader)
                {
                    case "会计科目对照":
                        tabItem = new DXTabItem() { Header = strHeader, Content = new ucCodeContrast() };
                        break;
                    case "部门-项目核算对照":
                        tabItem = new DXTabItem() { Header = strHeader, Content = new ucDeptItemContrast() };
                        break;
                    case "凭证导入":
                        tabItem = new DXTabItem() { Header = strHeader, Content = new ucVoucherImport() };
                        break;
                    case "凭证生成":
                        tabItem = new DXTabItem() { Header = strHeader, Content = new ucVoucherGenerate() };
                        break;
                    case "目标账套科目设置":
                        tabItem = new DXTabItem() { Header = strHeader, Content = new ucCodeAdd() };
                        break;
                }
                if (tabItem == null)
                    return;
                tabItem.Style = (Style)FindResource("DXTabItem_Style");
                MainTabControl.Items.Add(tabItem);
                MainTabControl.SelectedItem = tabItem;
            }
            catch (Exception ex)
            {
                tc.TrackException(ex);
                tc.TrackTrace(ex.Message);
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
		}

		private void btnTabClose_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				Button btn = sender as Button;
				string strHeader = btn.Tag.ToString();
				foreach (DXTabItem tabItem in MainTabControl.Items)
				{
					string strCurrentHeader = tabItem.Header.ToString();
					if (strCurrentHeader == strHeader)
					{
						MainTabControl.Items.Remove(tabItem);
						break;
					}
				}
				bool Visible = UICheck.CheckDXTabControlVisible(MainTabControl);
				if (!Visible)
					MainTabControl.Visibility = Visibility.Collapsed;
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
