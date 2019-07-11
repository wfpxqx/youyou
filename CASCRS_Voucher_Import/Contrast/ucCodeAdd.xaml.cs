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
    /// ucCodeAdd.xaml 的交互逻辑
    /// </summary>
    public partial class ucCodeAdd : UserControl
    {

        private TelemetryClient tc = new TelemetryClient();
        private ButtonEditSettings btnRef = null;
        private int LOGIN_YEAR = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Year;
        private int LOGIN_MONTH = Convert.ToDateTime(ConfigurationManager.AppSettings["LOGIN_DATE"]).Month;
        private string MIDDLE_ACC_ID = ConfigurationManager.AppSettings["MIDDLE_ACC_ID"];
        private string MIDDLE_ACC_YEAR = ConfigurationManager.AppSettings["MIDDLE_ACC_YEAR"];
        private string TARGET_ACC_ID = ConfigurationManager.AppSettings["TARGET_ACC_ID"];
        private string TARGET_ACC_YEAR = ConfigurationManager.AppSettings["TARGET_ACC_YEAR"];
        private SqlDataAdapter sdaCodeCst = null;
        private DataTable CodeCstDataSource = null;
        private DataTable dtMiddleCode = null;
        //private DataTable dtMiddleCodeClass = null;
        private DataTable dtTargetCode = null;
        private DataTable dtTargetCodeClass = null;

        public ucCodeAdd()
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

                CodeCstDataSource = DbOperation.GetDataTable(out sdaCodeCst, "SELECT * FROM CodeContrast where Flag=0");
                SetDataTable();
                SetCodeContrastGrid(CodeCstDataSource);
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
                StringBuilder sbSQL = new StringBuilder();
                sbSQL.AppendLine("SELECT");
                sbSQL.AppendLine("     ccode");
                sbSQL.AppendLine("    ,ccode_name");
                sbSQL.AppendLine("    ,igrade");
                sbSQL.AppendLine("    ,bend");
                sbSQL.AppendLine("    ,cclass");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    code");
                sbSQL.AppendLine("WHERE");
                sbSQL.AppendLine(string.Format("    iYear = {0}", LOGIN_YEAR));
                sbSQL.AppendLine("ORDER BY");
                sbSQL.AppendLine("    ccode");
                string strSQL = sbSQL.ToString();
                //dtMiddleCodeClass = DbOperation.GetDataTable(strSQL, 1);
                dtTargetCodeClass = DbOperation.GetDataTable(strSQL, 2);

                sbSQL.Clear();
                sbSQL.AppendLine("SELECT");
                sbSQL.AppendLine("     ccode");
                sbSQL.AppendLine("    ,ccode_name");
                sbSQL.AppendLine("    ,bdept");
                sbSQL.AppendLine("    ,bperson");
                sbSQL.AppendLine("    ,bcus");
                sbSQL.AppendLine("    ,bsup");
                sbSQL.AppendLine("    ,bitem");
                sbSQL.AppendLine("    ,cclass");
                sbSQL.AppendLine("FROM");
                sbSQL.AppendLine("    code");
                sbSQL.AppendLine("WHERE");
                sbSQL.AppendLine(string.Format("    iYear = {0}", LOGIN_YEAR));
                sbSQL.AppendLine("AND");
                sbSQL.AppendLine("    bend = 1");
                sbSQL.AppendLine("ORDER BY");
                sbSQL.AppendLine("    ccode");
                string strSQL2 = sbSQL.ToString();
                //dtMiddleCode = DbOperation.GetDataTable(strSQL2, 1);
                dtTargetCode = DbOperation.GetDataTable(strSQL2, 2);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void SetCodeContrastGrid(DataTable dtCodeCst)
        {
            try
            {
                gcCodeAdd.ItemsSource = dtCodeCst;
                gcCodeAdd.Columns[0].Header = "唯一标识";
                gcCodeAdd.Columns[1].Header = "借方科目编码";
                gcCodeAdd.Columns[2].Header = "借方科目名称";
                gcCodeAdd.Columns[3].Header = "贷方科目编码";
                gcCodeAdd.Columns[4].Header = "贷方科目名称";
                if (btnRef == null)
                {
                    btnRef = new ButtonEditSettings() { ShowNullText = false, ShowText = true };
                    btnRef.DefaultButtonClick += BtnRef_DefaultButtonClick;
                }
                gcCodeAdd.Columns[0].Visible = false;
                gcCodeAdd.Columns[1].Visible = false;
                gcCodeAdd.Columns[2].Visible = false;

                gcCodeAdd.Columns[5].Visible = false;
                //gcCodeAdd.Columns[1].EditSettings = btnRef;
                gcCodeAdd.Columns[3].EditSettings = btnRef;
                //gcCodeAdd.Columns[2].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
                gcCodeAdd.Columns[4].AllowEditing = DevExpress.Utils.DefaultBoolean.False;
                tvCodeAdd.BestFitMode = DevExpress.Xpf.Core.BestFitMode.AllRows;
                tvCodeAdd.BestFitColumns();
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
                object objCurrentItem = gcCodeAdd.CurrentItem;
                if (objCurrentItem == null)
                    return;
                ColumnBase cb = gcCodeAdd.CurrentColumn;
                if (cb == null)
                    return;
                tvCodeAdd.CancelRowEdit();
                DataRow dr = ((DataRowView)objCurrentItem).Row;
                string strHeader = cb.Header.ToString();
                if (strHeader == "借方科目编码")
                {
                    ReferenceWindow2 win = new ReferenceWindow2()
                    {
                        BaseTitle = "会计科目参照",
                        //TreeViewDataSource = dtMiddleCodeClass,
                        //GridDataSource = dtMiddleCode,
                        TreeViewDataSource = dtTargetCodeClass,
                        GridDataSource = dtTargetCode,
                        GridHeader = new string[]
                        {
                            "科目编码",
                            "科目名称",
                            "部门核算",
                            "职员核算",
                            "客户核算",
                            "供应商核算",
                            "项目核算",
                            "所属大类"
                        }
                    };
                    string strValue = dr["middleCode"].ToString();
                    if (!string.IsNullOrWhiteSpace(strValue))
                        win.Value = strValue;
                    win.ShowDialog();
                    DataRow drResult = win.SelectedRow;
                    if (drResult == null)
                        return;
                    dr.BeginEdit();
                    dr["middleCode"] = drResult[0];
                    dr["middleCodeName"] = drResult[1];
                    dr.EndEdit();
                }
                else if (strHeader == "贷方科目编码")
                {
                    ReferenceWindow2 win = new ReferenceWindow2()
                    {
                        BaseTitle = "会计科目参照",
                        TreeViewDataSource = dtTargetCodeClass,
                        GridDataSource = dtTargetCode,
                        GridHeader = new string[]
                        {
                            "科目编码",
                            "科目名称",
                            "部门核算",
                            "职员核算",
                            "客户核算",
                            "供应商核算",
                            "项目核算",
                            "所属大类"
                        }
                    };
                    string strValue = dr["targetCode"].ToString();
                    if (!string.IsNullOrWhiteSpace(strValue))
                        win.Value = strValue;
                    win.ShowDialog();
                    DataRow drResult = win.SelectedRow;
                    if (drResult == null)
                        return;
                    dr.BeginEdit();
                    dr["targetCode"] = drResult[0];
                    dr["targetCodeName"] = drResult[1];
                    dr.EndEdit();
                }
                gcCodeAdd.RefreshData();
                gcCodeAdd.Focus();
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
                tvCodeAdd.CancelRowEdit();
                int RowCount = CodeCstDataSource.Rows.Count;
                if (RowCount > 0)
                {
                    DataRow drLast = CodeCstDataSource.Rows[RowCount - 1];
                    string strMiddleCode = drLast["middleCode"].ToString();
                    if (string.IsNullOrWhiteSpace(strMiddleCode))
                    {
                        MessageBox.Show("借方科目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    string strTargetCode = drLast["targetCode"].ToString();
                    if (string.IsNullOrWhiteSpace(strTargetCode))
                    {
                        MessageBox.Show("贷方科目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                    DataRow[] drc = CodeCstDataSource.Select(string.Format("middleCode = '{0}' AND targetCode = '{1}'", strMiddleCode, strTargetCode));
                    if (drc.Length > 1)
                    {
                        MessageBox.Show(string.Format("借方科目编码{0}贷方科目编码{1}的行存在若干重复行", strMiddleCode, strTargetCode), "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                        return;
                    }
                }
                DataRow dr = CodeCstDataSource.NewRow();
                dr[0] = Guid.NewGuid();
                dr["Flag"] = 0;
                CodeCstDataSource.Rows.Add(dr);
                SaveData();
                tvCodeAdd.MoveLastRow();
                gcCodeAdd.RefreshData();
                gcCodeAdd.Focus();
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
                DbOperation.UpdateDataSource(sdaCodeCst, CodeCstDataSource);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void gcCodeAdd_CurrentItemChanged(object sender, CurrentItemChangedEventArgs e)
        {
            try
            {
                tvCodeAdd.CancelRowEdit();
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
                string strMiddleCode = dr["middleCode"].ToString();
                string strTargetCode = dr["targetCode"].ToString();
                if (string.IsNullOrWhiteSpace(strMiddleCode))
                {
                    MessageBox.Show("借方科目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (string.IsNullOrWhiteSpace(strTargetCode))
                {
                    MessageBox.Show("贷方科目编码不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                DataRow[] drc = CodeCstDataSource.Select(string.Format("middleCode = '{0}' AND targetCode = '{1}'", strMiddleCode, strTargetCode));
                if (drc.Length > 1)
                {
                    MessageBox.Show(string.Format("借方科目编码{0}贷方科目编码{1}的行存在若干重复行", strMiddleCode, strTargetCode), "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
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

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object objCurrentItem = gcCodeAdd.CurrentItem;
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

        private void tvCodeAdd_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            try
            {
                if (CodeCstDataSource == null)
                    return;
                if (CodeCstDataSource.Rows.Count != 1)
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

        private void tvCodeAdd_CellValueChanged(object sender, CellValueChangedEventArgs e)
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
                if (strHeader == "借方科目编码")
                {
                    DataRow[] drc = dtMiddleCode.Select(string.Format("ccode = '{0}'", e.Value));
                    dr.BeginEdit();
                    if (drc.Length == 0)
                        dr["middleCodeName"] = DBNull.Value;
                    else
                        dr["middleCodeName"] = drc[0][1];
                    dr.EndEdit();
                }
                else if (strHeader == "贷方帐套科目编码")
                {
                    DataRow[] drc = dtTargetCode.Select(string.Format("ccode = '{0}'", e.Value));
                    dr.BeginEdit();
                    if (drc.Length == 0)
                        dr["targetCodeName"] = DBNull.Value;
                    else
                        dr["targetCodeName"] = drc[0][1];
                    dr.EndEdit();
                }
                if (CodeCstDataSource == null)
                    return;
                if (CodeCstDataSource.Rows.Count != 1)
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
