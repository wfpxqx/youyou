using System;
using System.Data;
using System.Windows;
using DevExpress.Utils;

namespace CASCRS_Voucher_Import.Common.CommWindow
{
    /// <summary>
    /// ReferenceWindow.xaml 的交互逻辑
    /// </summary>
    public partial class ReferenceWindow : Window
    {
        public string BaseTitle { get; set; }
        public string[] GridHeader { get; set; }
        public DataRow Result { get; set; }
        public DataTable DataSource{ get; set; }

        /// <summary>
        /// 初始化
        /// </summary>
        public ReferenceWindow()
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
                if (DataSource == null)
                    return;
                gcReference.ItemsSource = DataSource;
                var count = gcReference.Columns.Count;
                for (var i = 0; i < count; i++)
                {
                    gcReference.Columns[i].Header = GridHeader[i];
                    gcReference.Columns[i].AllowEditing = DefaultBoolean.False;
                    gcReference.Columns[i].ReadOnly = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// 选中按钮点击事件
        /// </summary>
        /// <param name="sender">发送对象</param>
        /// <param name="e">触发事件</param>
        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (gcReference.CurrentItem == null) return;
                var drv = (DataRowView)gcReference.CurrentItem;
                Result = drv.Row;
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
                Window_Loaded(new object(), new RoutedEventArgs());
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
        private void gcReference_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                btnSelect_Click(new object(), new RoutedEventArgs());
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
