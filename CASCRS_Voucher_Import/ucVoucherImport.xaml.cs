using System;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows;

using CASCRS_Voucher_Import.Common.CommDB;
using DevExpress.Xpf.Grid;
using Microsoft.ApplicationInsights;

using System.IO;

using System.Windows.Forms;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

using NPOI.HSSF.UserModel;

using U8Voucher.Common.CommExcelOpt;

namespace CASCRS_Voucher_Import
{
	/// <summary>
	/// ucVoucherImport.xaml 的交互逻辑
	/// </summary>
	public partial class ucVoucherImport : System.Windows.Controls.UserControl
    {
		private TelemetryClient tc = new TelemetryClient();
		private int LOGIN_YEAR = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Year;
		private DataTable dtVoucherHeader = null;
		private DataTable dtVoucherDetail = null;
        private DataTable dtVoucherCash = null;
		private DataTable dtCodeCst = null;
		private DataTable dtDeptItemCst = null;
		private DataTable dtTargetCode = null;
		private DataTable dtMiddleCode = null;
		private DataTable dtDepartment = null;
		private bool IsImportEnded = false;
        public DataTable ExcelDataSource { get; set; }
        public ucVoucherImport()
		{
			InitializeComponent();
		}

		private void UserControl_Loaded(object sender, RoutedEventArgs e)
		{
			try
			{
                //	//ApplicationInsights跟踪
                //	tc.InstrumentationKey = "30b22ad1-3cdd-462f-885f-fce870bb83c4";
                //	tc.Context.User.Id = Environment.UserName;
                //	tc.Context.Session.Id = Guid.NewGuid().ToString();
                //	tc.Context.Device.OperatingSystem = Environment.OSVersion.ToString();
                //	tc.TrackPageView("中科腐蚀研究院凭证导入工具");

                //StringBuilder sbSQL = new StringBuilder();
                //	sbSQL.AppendLine("SELECT DISTINCT");
                //	sbSQL.AppendLine("	  REPLACE((csign+' - '+CONVERT(NVARCHAR(10), ino_id+10000)), csign+' - 1', csign+' - 0') AS ino_id");
                //	sbSQL.AppendLine("	 ,iperiod");
                //	sbSQL.AppendLine("	 ,iyear");
                //	sbSQL.AppendLine("	 ,ino_id AS real_inoid");
                //	sbSQL.AppendLine("INTO");
                //	sbSQL.AppendLine("    TMP_VoucherHeader");
                //	sbSQL.AppendLine("FROM");
                //	sbSQL.AppendLine("    dbo.GL_accvouch");
                //	sbSQL.AppendLine("WHERE");
                //	sbSQL.AppendLine("    cDefine11 IS NULL");
                //	sbSQL.AppendLine("AND");
                //	sbSQL.AppendLine("	  iflag IS NULL");
                //	sbSQL.AppendLine("AND");
                //	sbSQL.AppendLine("	  ino_id IS NOT NULL");
                //	sbSQL.AppendLine("AND");
                //	sbSQL.AppendLine("	  cdigest LIKE '%RD%'");
                //	sbSQL.AppendLine("AND");
                //	sbSQL.AppendLine(string.Format("	  iyear = {0}", LOGIN_YEAR));
                //	sbSQL.AppendLine("ORDER BY");
                //	sbSQL.AppendLine("    iyear, iperiod, ino_id");
                //	DbOperation.ExecuteNonQuery(sbSQL.ToString(), 1);
                //	dtVoucherHeader = DbOperation.GetDataTable("SELECT * FROM TMP_VoucherHeader", 1);

                //sbSQL.Clear();
                //sbSQL.AppendLine("SELECT");
                //sbSQL.AppendLine("    glav.*");
                //sbSQL.AppendLine("FROM");
                //sbSQL.AppendLine("    dbo.GL_accvouch AS glav");
                //sbSQL.AppendLine("INNER JOIN");
                //sbSQL.AppendLine("    dbo.TMP_VoucherHeader AS tmpvh");
                //sbSQL.AppendLine("ON");
                //sbSQL.AppendLine("    tmpvh.real_inoid = glav.ino_id");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("    tmpvh.iperiod = glav.iperiod");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("    tmpvh.iyear = glav.iyear");
                //sbSQL.AppendLine("WHERE");
                //sbSQL.AppendLine("    glav.cDefine11 IS NULL");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("	  glav.iflag IS NULL");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("	  glav.ino_id IS NOT NULL");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine(string.Format("	  glav.iyear = {0}", LOGIN_YEAR));
                //sbSQL.AppendLine("ORDER BY");
                //sbSQL.AppendLine("    glav.iyear, glav.iperiod, glav.ino_id");
                //dtVoucherDetail = DbOperation.GetDataTable(sbSQL.ToString(), 1);

                //sbSQL.Clear();
                //sbSQL.AppendLine("SELECT");
                //sbSQL.AppendLine("    glct.*");
                //sbSQL.AppendLine("FROM");
                //sbSQL.AppendLine("    GL_CashTable AS glct");
                //sbSQL.AppendLine("INNER JOIN");
                //sbSQL.AppendLine("    TMP_VoucherHeader AS tmpvh");
                //sbSQL.AppendLine("ON");
                //sbSQL.AppendLine("    tmpvh.real_inoid = glct.ino_id");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("    tmpvh.iperiod = glct.iperiod");
                //sbSQL.AppendLine("AND");
                //sbSQL.AppendLine("    tmpvh.iyear = glct.iyear");
                //dtVoucherCash = DbOperation.GetDataTable(sbSQL.ToString(), 1);
                //DbOperation.ExecuteNonQuery("DROP TABLE TMP_VoucherHeader", 1);
                //SetDataTable();
                //SetVoucherHeaderGrid(dtVoucherHeader);
                //SetVoucherDetailGrid(dtVoucherDetail);
            }
            catch (Exception ex)
            {
                tc.TrackException(ex);
                tc.TrackTrace(ex.Message);
                System.Windows.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

		private void SetDataTable()
		{
			try
			{
                //会计科目对照
				dtCodeCst = DbOperation.GetDataTable("SELECT * FROM CodeContrast where Flag=1");
                //部门项目对照
				dtDeptItemCst = DbOperation.GetDataTable("SELECT * FROM DeptItemContrast");
                // 2 目标帐套
				dtTargetCode = DbOperation.GetDataTable(string.Format("SELECT ccode, ccode_name, bitem, bdept FROM code WHERE iyear = {0}", LOGIN_YEAR), 2);
                //1  中间帐套
                dtMiddleCode = DbOperation.GetDataTable(string.Format("SELECT ccode, ccode_name, bitem, bdept FROM code WHERE iyear = {0}", LOGIN_YEAR), 1);
                 //1  中间帐套
                dtDepartment = DbOperation.GetDataTable("SELECT cDepCode, cDepName FROM Department", 1);
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private void SetVoucherHeaderGrid(DataTable dtVH)
		{
			try
			{
				gcVoucherHeader.ItemsSource = dtVH;
				gcVoucherHeader.Columns[0].Header = "凭证号";
				gcVoucherHeader.Columns[1].Header = "期间";
				gcVoucherHeader.Columns[2].Header = "年度";
				gcVoucherHeader.Columns[3].Header = "凭证编号";
				gcVoucherHeader.Columns[3].Visible = false;
				tvVoucherHeader.BestFitMode = DevExpress.Xpf.Core.BestFitMode.AllRows;
				tvVoucherHeader.BestFitColumns();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private	void SetVoucherDetailGrid(DataTable dtVD)
		{
			try
			{
				gcVoucherDetail.ItemsSource = dtVD;
				foreach (GridColumn gc in gcVoucherDetail.Columns)
				{
					gc.Visible = false;
				}
				gcVoucherDetail.Columns["cdigest"].Header = "摘要";
				gcVoucherDetail.Columns["ccode"].Header = "会计科目";
				gcVoucherDetail.Columns["md"].Header = "借方金额";
				gcVoucherDetail.Columns["mc"].Header = "贷方金额";
				gcVoucherDetail.Columns["cdept_id"].Header = "部门";
				gcVoucherDetail.Columns["citem_class"].Header = "项目大类编码";
				gcVoucherDetail.Columns["citem_id"].Header = "项目编码";
				gcVoucherDetail.Columns["cdigest"].Visible = true;
				gcVoucherDetail.Columns["ccode"].Visible = true;
				gcVoucherDetail.Columns["md"].Visible = true;
				gcVoucherDetail.Columns["mc"].Visible = true;
				gcVoucherDetail.Columns["cdept_id"].Visible = true;
				gcVoucherDetail.Columns["citem_class"].Visible = true;
				gcVoucherDetail.Columns["citem_id"].Visible = true;
				ObservableCollection<GridSummaryItem> oc = new ObservableCollection<GridSummaryItem>();
				oc.Add(new GridSummaryItem() { FieldName = "md", DisplayFormat = "借方金额总计：{0}", SummaryType = DevExpress.Data.SummaryItemType.Sum, Visible = true });
				oc.Add(new GridSummaryItem() { FieldName = "mc", DisplayFormat = "贷方金额总计：{0}", SummaryType = DevExpress.Data.SummaryItemType.Sum, Visible = true });
				gcVoucherDetail.TotalSummarySource = oc;
				//gcVoucherDetail.UpdateTotalSummary();
				tvVoucherDetail.BestFitMode = DevExpress.Xpf.Core.BestFitMode.VisibleRows;
				tvVoucherDetail.BestFitColumns();
			}
			catch (Exception ex)
			{
				throw ex;
			}
		}

		private void gcVoucherHeader_CurrentItemChanged(object sender, CurrentItemChangedEventArgs e)
		{
			try
			{
				if (e == null)
					return;
				if (e.NewItem == null)
					return;
				DataRow dr = ((DataRowView)e.NewItem).Row;
				if (dr.RowState == DataRowState.Deleted || dr.RowState == DataRowState.Detached)
					return;
				int iPeriod = Convert.ToInt32(dr["iperiod"]);
				int inoid = Convert.ToInt32(dr["real_inoid"]);
				DataRow[] drcVD = dtVoucherDetail.Select(string.Format("iperiod = {0} AND ino_id = {1}", iPeriod, inoid));
				if(drcVD.Length > 0)
				{
					SetVoucherDetailGrid(drcVD.CopyToDataTable());
				}
				else
				{
					SetVoucherDetailGrid(dtVoucherDetail.Clone());
				}
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
                System.Windows.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void gcVoucherDetail_CustomColumnDisplayText(object sender, CustomColumnDisplayTextEventArgs e)
		{
			try
			{
				if (e.Row == null)
					return;
				if (e.Column == null)
					return;
				if (e.Value == null)
					return;
				e.ShowAsNullText = true;
				GridColumn gc = e.Column;
				string strField = gc.FieldName;
				if (strField == "ccode")
				{
					DataRow[] drcCode = null;
					if (IsImportEnded)
					{
						if (dtTargetCode == null)
							return;
						drcCode = dtTargetCode.Select(string.Format("ccode = '{0}'", e.Value.ToString()));
					}
					else
					{
						if (dtMiddleCode == null)
							return;
						drcCode = dtMiddleCode.Select(string.Format("ccode = '{0}'", e.Value.ToString()));
					}
					if (drcCode.Length == 0)
						return;
					string strCodeName = drcCode[0]["ccode_name"].ToString();
					e.DisplayText = strCodeName;
				}
				else if (strField == "cdept_id")
				{
					if (dtDepartment == null)
						return;
					DataRow[] drcDept = dtDepartment.Select(string.Format("cDepCode = '{0}'", e.Value.ToString()));
					if (drcDept.Length == 0)
						return;
					string strDeptname = drcDept[0][1].ToString();
					e.DisplayText = strDeptname;
				}
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
                System.Windows.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}

		private void btnImport_Click(object sender, RoutedEventArgs e)
		{
			SqlConnection objSqlConn = null;
			try
			{
                foreach (DataRow drVD in dtVoucherDetail.Rows)
                
                //foreach (DataRow drVD in ExcelDataSource.Rows)
                {
					string strMiddleCode = drVD["ccode"].ToString();
					string strMiddleDeptId = drVD["cdept_id"].ToString();
					string strTargetCode = null;
					drVD.BeginEdit();
					DataRow[] drcCodeCst = dtCodeCst.Select(string.Format("middleCode = '{0}'", strMiddleCode));
					if (drcCodeCst.Length > 0)
						strTargetCode = drcCodeCst[0]["targetCode"].ToString();
					bool bItem = false;
					bool bDept = false;
					DataRow[] drcTargetCode = dtTargetCode.Select(string.Format("ccode = '{0}'", strTargetCode));
					if(drcTargetCode.Length > 0)
					{
						drVD["ccode"] = strTargetCode;
						bItem = Convert.ToBoolean(drcTargetCode[0]["bitem"]);
						bDept = Convert.ToBoolean(drcTargetCode[0]["bdept"]);
					}
					DataRow[] drcDeptItemCst = dtDeptItemCst.Select(string.Format("deptId = '{0}'", strMiddleDeptId));
					if (drcDeptItemCst.Length > 0)
					{
						if (!bDept)
							drVD["cdept_id"] = DBNull.Value;
						if(bItem)
						{
							drVD["citem_class"] = drcDeptItemCst[0]["itemCClass"];
							drVD["citem_id"] = drcDeptItemCst[0]["itemId"];
						}
					}
					drVD.EndEdit();
				}
				DataTable dtVDMerge = dtVoucherDetail.Clone();
				DataTable dtVDMergeCopy = dtVoucherDetail.Clone();
				string strCode = null;
				string strItemCCode = null;
				string strItemCode = null;
				objSqlConn = new SqlConnection(DbOperation.connectionString1);
				objSqlConn.Open();
				SqlCommand objSqlCmd = new SqlCommand() { Connection = objSqlConn };
				int inoid = 1;
				int iPeriod_Ref = -1;
				foreach (DataRow drVH in dtVoucherHeader.Rows)
				{
					int iPeriod = Convert.ToInt32(drVH["iperiod"]);
					int inoid2 = 1;
					int.TryParse(drVH["real_inoid"].ToString(), out inoid2);
					DataRow[] drcVD = dtVoucherDetail.Select(string.Format("iperiod = {0} AND ino_id = {1}", iPeriod, inoid2));
                    DataRow[] drcVC = dtVoucherCash.Select(string.Format("iperiod = {0} AND ino_id = {1}", iPeriod, inoid2));
                    dtVDMergeCopy.Rows.Clear();
					if(iPeriod_Ref != iPeriod)
					{
						object MaxInoId = DbOperation.GetDataTable(string.Format("SELECT MAX(ino_id) FROM GL_accvouch WHERE iyear = {0} AND iperiod = {1}", LOGIN_YEAR, iPeriod), 2).Rows[0][0];
						int.TryParse(MaxInoId.ToString(), out inoid);
					}
					inoid++;
					foreach (DataRow drVD in drcVD)
					{
						
						strCode = drVD["ccode"].ToString();
						strItemCCode = drVD["citem_class"].ToString();
						strItemCode = drVD["citem_id"].ToString();
						DataRow[] drcVDPrev = dtVDMergeCopy.Select(string.Format("ccode = '{0}' AND citem_class = '{1}' AND citem_id = '{2}'", strCode, strItemCCode, strItemCode));
						if (drcVDPrev.Length == 0)
						{
							object objInoid = drVD["ino_id"];
							drVD.BeginEdit();
							drVD["ino_id"] = inoid;
							drVD.EndEdit();
							dtVDMergeCopy.Rows.Add(drVD.ItemArray);
							drVD.BeginEdit();
							drVD["ino_id"] = objInoid;
							drVD.EndEdit();
						}
						else
						{
							DataRow drVDPrev = drcVDPrev[0];
							decimal decMdPrev = decimal.Zero;
							decimal decMcPrev = decimal.Zero;
							decimal decMd = decimal.Zero;
							decimal decMc = decimal.Zero;
							decimal.TryParse(drVDPrev["md"].ToString(), out decMdPrev);
							decimal.TryParse(drVDPrev["mc"].ToString(), out decMcPrev);
							decimal.TryParse(drVD["md"].ToString(), out decMd);
							decimal.TryParse(drVD["mc"].ToString(), out decMc);
							drVDPrev.BeginEdit();
							drVDPrev["md"] = decMdPrev + decMd;
							drVDPrev["mc"] = decMcPrev + decMc;
							drVDPrev.EndEdit();
						}
					}
                    foreach (DataRow drVC in drcVC)
                    {
                        drVC.BeginEdit();
                        drVC["ino_id"] = inoid;
                        drVC.EndEdit();
                    }
					dtVDMerge.Merge(dtVDMergeCopy, true, MissingSchemaAction.Ignore);
					objSqlCmd.CommandText = string.Format("UPDATE GL_accvouch SET cDefine11 = '已导入目标帐套' WHERE ino_id = {0} AND iperiod = {1} AND iyear = {2}",
						inoid2, iPeriod, LOGIN_YEAR);
					objSqlCmd.ExecuteNonQuery();
					iPeriod_Ref = iPeriod;
				}
				DbOperation.ExecuteSqlBulkCopy(dtVDMerge, "GL_accvouch", 2);
                DbOperation.ExecuteSqlBulkCopy(dtVoucherCash, "GL_CashTable", 2);
				UserControl_Loaded(new object(), new RoutedEventArgs());
				IsImportEnded = true;
                System.Windows.MessageBox.Show("导入完成", "信息", MessageBoxButton.OK, MessageBoxImage.Information);
			}
			catch (Exception ex)
			{
				tc.TrackException(ex);
				tc.TrackTrace(ex.Message);
                System.Windows.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
        /// <summary>
        /// 导入按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnImportt_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var a = 1;
                OpenFileDialog objOFD = new OpenFileDialog() { Filter = "Excel文件|*.xlsx;*.xls", InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Multiselect = false };
                if (objOFD.ShowDialog() == DialogResult.Cancel) return;
                var FilePath = objOFD.FileName;
                IWorkbook workbook = null;
                string FileExt = Path.GetExtension(FilePath);
                if (".xls".Equals(FileExt)) using (FileStream objFS = new FileStream(FilePath, FileMode.Open, FileAccess.Read)) workbook = new HSSFWorkbook(objFS);
                else using (FileStream objFS = new FileStream(FilePath, FileMode.Open, FileAccess.Read)) workbook = new XSSFWorkbook(objFS);
                var sheet = workbook.GetSheetAt(0);
                ExcelDataSource = ExcelOpt.Export2DataTable(sheet,3, true, false, FileExt);
                //DataRow[] drcDataSource = ExcelDataSource.Select("金额 = '0.00' OR 金额 = '0'");
                //if (drcDataSource.Length > 0) foreach (DataRow dr in drcDataSource) ExcelDataSource.Rows.Remove(dr);
                //DbOperation.ExecuteNonQuery("TRUNCATE TABLE ImportData");
                //DbOperation.ExecuteSqlBulkCopy(ExcelDataSource, "ImportData");
                DbOperation.ExecuteNonQuery("TRUNCATE TABLE TMP_VoucherHeader", 0);
                for (int i = 0; i < 3; i++)
                {
                    ExcelDataSource.Rows.Remove(ExcelDataSource.Rows[ExcelDataSource.Rows.Count - 1]);
                }
                DbOperation.ExecuteSqlBulkCopy(ExcelDataSource, "TMP_VoucherHeader",0);
                
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.AppendLine("SELECT DISTINCT");
                //sbSQL.AppendLine("	 convert(int, replace(ltrim(replace(substring(ino_Id,CHARINDEX('-',ino_id)+1,CHARINDEX('-',ino_id)+3),'0',' ')),' ','0'))  AS real_inoid");
                sbSQL.AppendLine("	substring(ino_Id,CHARINDEX('-',ino_id)+2,CHARINDEX('-',ino_id)+3) AS real_inoid");
                sbSQL.AppendLine("	 ,MONTH(dbill_date) as  iperiod");
                sbSQL.AppendLine("	 ,iYear");
                sbSQL.AppendLine("	 ,ino_id ");
                sbSQL.AppendLine("INTO");
                sbSQL.AppendLine("    TMP_VoucherHeader");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    CASCRS_VOUCHER.dbo.TMP_VoucherHeader");
                sbSQL.AppendLine("WHERE");                
                sbSQL.AppendLine("	  Remarks LIKE '%rd%'");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine(string.Format("	  iyear = {0}", LOGIN_YEAR));
                sbSQL.AppendLine("ORDER BY");
                sbSQL.AppendLine("  iyear, MONTH(dbill_date), ino_id");
                DbOperation.ExecuteNonQuery(sbSQL.ToString(), 1);
                dtVoucherHeader = DbOperation.GetDataTable("SELECT * FROM TMP_VoucherHeader", 1);

                sbSQL.Clear();
                sbSQL.AppendLine("SELECT");
                sbSQL.AppendLine("    glav.*");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    dbo.GL_accvouch AS glav");
                sbSQL.AppendLine("INNER JOIN");
                sbSQL.AppendLine("    dbo.TMP_VoucherHeader AS tmpvh");
                sbSQL.AppendLine("ON");
                sbSQL.AppendLine("    convert(int,tmpvh.real_inoid) = glav.ino_id");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("    tmpvh.iperiod = glav.iperiod");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("    tmpvh.iyear = glav.iyear");
                sbSQL.AppendLine("WHERE");
                sbSQL.AppendLine("    glav.cDefine11 IS NULL");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("	  glav.iflag IS NULL");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("	  glav.ino_id IS NOT NULL");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine(string.Format("	  glav.iyear = {0}", LOGIN_YEAR));
                sbSQL.AppendLine("ORDER BY");
                sbSQL.AppendLine("    glav.iyear, glav.iperiod, glav.ino_id");
                dtVoucherDetail = DbOperation.GetDataTable(sbSQL.ToString(), 1);

                sbSQL.Clear();
                sbSQL.AppendLine("SELECT");
                sbSQL.AppendLine("    glct.*");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    GL_CashTable AS glct");
                sbSQL.AppendLine("INNER JOIN");
                sbSQL.AppendLine("    TMP_VoucherHeader AS tmpvh");
                sbSQL.AppendLine("ON");
                sbSQL.AppendLine("    tmpvh.real_inoid = glct.ino_id");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("    tmpvh.iperiod = glct.iperiod");
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("    tmpvh.iyear = glct.iyear");
                dtVoucherCash = DbOperation.GetDataTable(sbSQL.ToString(), 1);
                DbOperation.ExecuteNonQuery("DROP TABLE TMP_VoucherHeader", 1);
                SetDataTable();
                SetVoucherHeaderGrid(dtVoucherHeader);
                SetVoucherDetailGrid(dtVoucherDetail);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
