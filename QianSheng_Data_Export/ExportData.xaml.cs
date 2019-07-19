using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.IO;
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
using Microsoft.ApplicationInsights;
using QianSheng_Data_Export.Common.BLL;
using QianSheng_Data_Export.Common.CommDB;
using QianSheng_Data_Export.Common.CommExcelOpt;

namespace QianSheng_Data_Export
{
	/// <summary>
	/// Login.xaml 的交互逻辑
	/// </summary>
	public partial class Login : Window
	{
        #region 已有的事件和方法
        public bool LoginCancelled { get; set; }
        private TelemetryClient tc = new TelemetryClient();
        public Login()
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
                tc.InstrumentationKey = "{986F10E7-32B8-4776-AA26-78FCE47C0B10}";// "30b22ad1-3cdd-462f-885f-fce870bb83c4";
                tc.Context.User.Id = Environment.UserName;
                tc.Context.Session.Id = Guid.NewGuid().ToString();
                tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
                tc.TrackPageView("千盛数据导出工具");
                LoginCancelled = true;
                txtTargetAccId.Text = ConfigurationManager.AppSettings["TARGET_ACC_ID"].ToString();
                txtTargetAccYear.Text = ConfigurationManager.AppSettings["TARGET_ACC_Year"].ToString();
                txtServer.Text = ConfigurationManager.AppSettings["SERVER_NAME"].ToString();
                txtDbPassword.Text = ConfigurationManager.AppSettings["DB_PASSWORD"].ToString();
                txtTargetAccId.Text = ConfigurationManager.AppSettings["TARGET_ACC_ID"].ToString();
                txtTargetAccYear.Text = ConfigurationManager.AppSettings["TARGET_ACC_Year"].ToString();

                Visibility = Visibility.Visible;

                txtTargetAccId.Focus();


            }
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnConfirm_Click(object sender, RoutedEventArgs e)
		{
            btnConfirm.IsEnabled = false;
            SystemConfig();
            btnConfirm.IsEnabled = true;
        }


        #endregion
        #region 验证
       
        private void SystemConfig()
        {
            SqlConnection objSqlConn = null;
            try
            {
                string strServer = txtServer.Text;
                string strDbUserPassword = txtDbPassword.Text;
                string strTargetAccId = txtTargetAccId.Text;
                string strTargetAccYear = txtTargetAccYear.Text;
                string strDbName = string.Format("UFDATA_{0}_{1}", strTargetAccId, strTargetAccYear);// txtDbName.Text;
                DirectoryInfo desPath = null ;
                if (strServer == string.Empty)
                {
                    MessageBox.Show("服务器名不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (strTargetAccId == string.Empty)
                {
                    MessageBox.Show("导出账套编号不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (strTargetAccYear == string.Empty)
                {
                    MessageBox.Show("导出账套年份不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (txtPath.Text == string.Empty)
                {
                    MessageBox.Show("导出文件夹不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                else
                {
                    DirectoryInfo dir = new DirectoryInfo(txtPath.Text);
                    if (!dir.Exists)
                    {
                        MessageBox.Show("导出文件夹不存在。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    else
                    {
                        try
                        {
                            desPath = dir.CreateSubdirectory(strTargetAccId + "\\" + strTargetAccYear);
                        }
                        catch
                        {

                        }
                    }
                }


                StringBuilder sbConnectionStrings = new StringBuilder();
                sbConnectionStrings.Append(string.Format("Data Source={0};", strServer));
                sbConnectionStrings.Append(string.Format("Initial Catalog={0};", strDbName));
                sbConnectionStrings.Append("Integrated Security=False;");
                sbConnectionStrings.Append("User ID=sa;");
                sbConnectionStrings.Append(string.Format("Password={0};", strDbUserPassword));
                sbConnectionStrings.Append("Connect Timeout=15;Encrypt=False;TrustServerCertificate=False");
                string ConnectionStrings = sbConnectionStrings.ToString();
                objSqlConn = new SqlConnection(ConnectionStrings);
                objSqlConn.Open();
                Configuration objConfig = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
                objConfig.AppSettings.Settings["CASCRS_DB_CONN"].Value = ConnectionStrings;
                objConfig.AppSettings.Settings["SERVER_NAME"].Value = strServer;
                objConfig.AppSettings.Settings["DB_PASSWORD"].Value = strDbUserPassword;
                objConfig.AppSettings.Settings["DB_NAME"].Value = strDbName;
                objConfig.AppSettings.Settings["TARGET_ACC_YEAR"].Value = strTargetAccYear;
                objConfig.AppSettings.Settings["TARGET_ACC_ID"].Value = strTargetAccId;
                objConfig.AppSettings.Settings["FIRST_RUN"].Value = "0";
                objConfig.Save();
                ConfigurationManager.RefreshSection("appSettings");
                DbOperation.connectionString = ConnectionStrings;

                ExportDataToCSV(desPath.FullName);
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
              
            }
        }
        #endregion
        #region 导出相关的方法
        private void ExportDataToCSV(string pullPath)
        {
            try
            {
                int iyear = int.Parse(txtTargetAccYear.Text);
                DataTable dtResult = GetExportData.GetAccount(iyear);
                if (!pullPath.EndsWith(@"\"))
                    pullPath += "\\";
                string fileName = "kj_ztjk_.csv";
                CSVHelper.SaveCSV(dtResult, pullPath + fileName);

                dtResult = GetExportData.GetCode(iyear);
                fileName = "kj_kmjk_.csv";
                CSVHelper.SaveCSV(dtResult, pullPath + fileName);

                dtResult = GetExportData.GetVoucherSum(iyear);
                fileName = "kj_yejk_.csv";
                CSVHelper.SaveCSV(dtResult, pullPath + fileName);

                dtResult = GetExportData.GetVoucher(iyear);
                if (dtResult.Rows.Count > 10000000)//超过1000万行，按月导
                {
                    DataRow[] drs;
                    DataTable dtMResult;
                    for (int iperiod = 1; iperiod < 13; iperiod++)
                    {
                        drs = dtResult.Select(string.Format("KJQJ={0}", iperiod));
                        if (drs.Length > 0)
                        {
                            dtMResult = drs.CopyToDataTable();
                            fileName = string.Format("kj_pzjk_{0}.csv", iperiod.ToString().PadLeft(2, '0'));
                            CSVHelper.SaveCSV(dtResult, pullPath + fileName);
                        }
                    }
                }
                else
                {
                    fileName = "kj_pzjk_.csv";
                    CSVHelper.SaveCSV(dtResult, pullPath + fileName);
                }
                MessageBox.Show("导出成功！", "信息", MessageBoxButton.OK, MessageBoxImage.Information);

            } 
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        #endregion 
    }
}
