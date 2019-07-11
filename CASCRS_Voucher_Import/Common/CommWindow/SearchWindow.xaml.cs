using DevExpress.Xpf.Editors;
using DevExpress.Xpf.Grid;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Text;
using System.Windows;
using System.Windows.Documents;

namespace CASCRS_Voucher_Import.Common.CommWindow
{
    /// <summary>
    /// SearchWindow.xaml 的交互逻辑
    /// </summary>
    public partial class SearchWindow : Window
    {
        //public string[] GridColNames { get; set; }
		public ObservableCollection<GridColumn> Columns { get; set; }
		public string strResult { get; set; }
        private DataTable ConditionSource = null;
		private Dictionary<string, string> dicField = null;

		/// <summary>
		/// 初始化
		/// </summary>
		public SearchWindow()
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
				txtValue.Visibility = Visibility.Visible;
				dteValue.Visibility = Visibility.Hidden;
				strResult = string.Empty;
                cbxField.Items.Clear();
				dicField = new Dictionary<string, string>();
				foreach (GridColumn gc in Columns)
				{
					dicField.Add(gc.FieldName, gc.HeaderCaption.ToString());
				}
				cbxField.ItemsSource = dicField;
                cbxField.IsTextEditable = false;
                cbxOperator.Items.Clear();
                cbxOperator.Items.Add("等于");
                cbxOperator.Items.Add("不等于");
                cbxOperator.Items.Add("包含");
                cbxOperator.Items.Add("不包含");
                cbxOperator.IsTextEditable = false;

                cbxLogicRelation.Items.Clear();
                cbxLogicRelation.Items.Add("并且");
                cbxLogicRelation.Items.Add("或者");
                cbxLogicRelation.IsTextEditable = false;

                ConditionSource = new DataTable();
                ConditionSource.Columns.Add("LogicRelation", Type.GetType("System.String"));
                ConditionSource.Columns.Add("HeaderCaption", Type.GetType("System.String"));
                ConditionSource.Columns.Add("operator", Type.GetType("System.String"));
                ConditionSource.Columns.Add("value", Type.GetType("System.String"));
				ConditionSource.Columns.Add("FieldName", Type.GetType("System.String"));
				gcCondtion.ItemsSource = ConditionSource;
                gcCondtion.Columns[0].Header = "逻辑关系";
                gcCondtion.Columns[1].Header = "名称";
                gcCondtion.Columns[2].Header = "符号";
                gcCondtion.Columns[3].Header = "值";
				gcCondtion.Columns[4].Header = "字段";
				gcCondtion.Columns[0].ReadOnly = true;
                gcCondtion.Columns[1].ReadOnly = true;
                gcCondtion.Columns[2].ReadOnly = true;
				gcCondtion.Columns[4].Visible = false;
			}
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 设置条件字符串
        /// </summary>
        private void SetConditionExpression()
        {
            try
            {
                if (ConditionSource == null)
                {
                    txtConditionExpression.Clear();
                    return;
                }

                int count = ConditionSource.Rows.Count;
                if (count == 0)
                {
                    txtConditionExpression.Clear();
                    return;
                }
                StringBuilder sb = new StringBuilder();
				StringBuilder sb2 = new StringBuilder();
				for (int i = 0; i < count; i++)
                {
                    string LogicRelation = ConditionSource.Rows[i].ItemArray[0].ToString();
                    string HeaderCaption = ConditionSource.Rows[i].ItemArray[1].ToString();
                    string Operator = ConditionSource.Rows[i].ItemArray[2].ToString();
                    string Value = ConditionSource.Rows[i].ItemArray[3].ToString();
					string FieldName = ConditionSource.Rows[i].ItemArray[4].ToString();
					if (!string.Empty.Equals(LogicRelation))
                    {
                        if (LogicRelation.Contains("并且")) sb.Append("AND");
                        else if (LogicRelation.Contains("或者")) sb.AppendFormat("OR");
                    }
                    if ("等于".Equals(Operator))
					{
						sb.AppendFormat("{0} = '{1}'", HeaderCaption, Value);
						sb2.AppendFormat("{0} = '{1}'", FieldName, Value);
					}
                    else if ("不等于".Equals(Operator))
					{
						sb.AppendFormat("{0} <> '{1}'", HeaderCaption, Value);
						sb2.AppendFormat("{0} <> '{1}'", FieldName, Value);
					}
                    else if ("包含".Equals(Operator))
					{
						sb.AppendFormat("{0} Like '%{1}%'", HeaderCaption, Value);
						sb2.AppendFormat("{0} Like '%{1}%'", FieldName, Value);
					}
                    else if ("不包含".Equals(Operator))
					{
						sb.AppendFormat("{0} Not Like '%{1}%'", HeaderCaption, Value);
						sb2.AppendFormat("{0} Not Like '%{1}%'", FieldName, Value);
					}
                }
                txtConditionExpression.Clear();
                txtConditionExpression.Text = sb.ToString();
				strResult = sb2.ToString();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 增加按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ConditionSource == null)
					return;
				if (string.Empty.Equals(txtValue.Text.Trim()) && string.Empty.Equals(dteValue.DateTime.ToString("yyyy-MM-dd")))
                {
                    MessageBox.Show("值不能为空。", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
				KeyValuePair<string, string> dicSelectItem = new KeyValuePair<string, string>();
				dicSelectItem = (KeyValuePair<string, string>)cbxField.SelectedItem;
				int count = ConditionSource.Rows.Count;
                DataRow dr = ConditionSource.NewRow();
                if (count == 0)
					dr[0] = string.Empty;
                else
					dr[0] = cbxLogicRelation.Text;
                dr[1] = dicSelectItem.Value;
                dr[2] = cbxOperator.Text;
				if(dteValue.Visibility == Visibility.Visible)
				{
					if(cbxField.Text == "凭证日期")
					{
						dr[3] = dteValue.DateTime.ToString("yyyy-MM-dd");
					}
					else if(cbxField.Text == "制单时间" || cbxField.Text == "报账时间")
					{
						dr[3] = dteValue.DateTime.ToString("yyyy-MM-dd HH:mm:ss");
					}
				}
				else
				{
					dr[3] = txtValue.Text;
				}
				dr[4] = dicSelectItem.Key;
                ConditionSource.Rows.Add(dr);
                SetConditionExpression();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 取反按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnReverse_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                object objCurrentItem = gcCondtion.CurrentItem;
                if (objCurrentItem == null) return;
                DataRowView drvCurrentItem = (DataRowView)objCurrentItem;
                DataRow drCurrentItem = drvCurrentItem.Row;
                drCurrentItem.BeginEdit();
                string Operator = drCurrentItem[2].ToString();
                if ("等于".Equals(Operator)) drCurrentItem[2] = "不等于";
                else if ("不等于".Equals(Operator)) drCurrentItem[2] = "等于";
                else if ("包含".Equals(Operator)) drCurrentItem[2] = "不包含";
                else if ("不包含".Equals(Operator)) drCurrentItem[2] = "包含";
                drCurrentItem.EndEdit();
                SetConditionExpression();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 删除按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (ConditionSource == null) return;
                object objCurrentItem = gcCondtion.CurrentItem;
                if (objCurrentItem == null) return;
                DataRowView drvCurrentItem = (DataRowView)objCurrentItem;
                DataRow drCurrentItem = drvCurrentItem.Row;
                ConditionSource.Rows.Remove(drCurrentItem);
                int count = ConditionSource.Rows.Count;
                if (count > 0)
                {
                    DataRow dr = ConditionSource.Rows[0];
                    dr.BeginEdit();
                    dr[0] = string.Empty;
                    dr.EndEdit();
                }
                SetConditionExpression();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 清空按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                try
                {
                    ConditionSource.Rows.Clear();
                    txtConditionExpression.Clear();
					strResult = null;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 执行按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnExecute_Click(object sender, RoutedEventArgs e)
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
        /// 关闭按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
				strResult = null;
				Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

		private void cbxField_SelectedIndexChanged(object sender, RoutedEventArgs e)
		{
			try
			{
				string Field = cbxField.Text;
				if(Field.Contains("日期") || Field.Contains("时间"))
				{
					txtValue.Visibility = Visibility.Hidden;
					dteValue.Visibility = Visibility.Visible;
					if(cbxField.Text == "凭证日期")
					{
						dteValue.DisplayFormatString = "yyyy-MM-dd";
					}
					else if(cbxField.Text == "制单时间" || cbxField.Text == "报账时间")
					{
						dteValue.DisplayFormatString = "yyyy-MM-dd HH:mm:ss";
					}
				}
				else
				{
					txtValue.Visibility = Visibility.Visible;
					dteValue.Visibility = Visibility.Hidden;
				}
					
			}
			catch (Exception ex)
			{
				MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
			}
		}
	}
}
