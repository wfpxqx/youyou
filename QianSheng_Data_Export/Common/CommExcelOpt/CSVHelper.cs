using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace QianSheng_Data_Export.Common.CommExcelOpt
{
    /// <summary>
    /// 对CSV文件的操作
    /// </summary>
    public class CSVHelper
    {
        /// <summary>
        /// 写入CSV文件
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="fileName">文件全名</param>
        /// <returns>是否写入成功</returns>
        public static bool SaveCSV(DataTable dt, string fullFileName = "")
        {
         
            try
            {
                FileStream fs = new FileStream(fullFileName, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default);
                string data = "";

                //写出列名称
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    data += dt.Columns[i].ColumnName.ToString();
                    if (i < dt.Columns.Count - 1)
                    {
                        data += ",";
                    }
                }
                sw.WriteLine(data);

                //写出各行数据
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    data = "";
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        data += dt.Rows[i][j].ToString();
                        if (j < dt.Columns.Count - 1)
                        {
                            data += ",";
                        }
                    }
                    sw.WriteLine(data);
                }

                sw.Close();
                fs.Close();

                return true;
            }
            catch (Exception e)
            {
                System.Windows.MessageBox.Show(e.Message);
                return false;
            }

        }

        /// <summary>
        /// 打开CSV 文件
        /// </summary>
        /// <param name="fileName">文件全名</param>
        /// <returns>DataTable</returns>
        public static DataTable OpenCSV(string fileName = "")
        {
            if (fileName == "")
            {
                System.Windows.MessageBox.Show("目标CSV文件不存在,请手动选择数据文件.");
                System.Windows.Forms.OpenFileDialog ofd = new System.Windows.Forms.OpenFileDialog();
                ofd.Filter = "CSV文件(*.csv)|*.csv";
                ofd.DefaultExt = "csv";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileName = ofd.FileName;
                    return OpenCSV(fileName, 0, 0, 0, 0, true);
                }
                else
                {
                    return null;
                }
            }
            else
            {
                try
                {
                    return OpenCSV(fileName, 0, 0, 0, 0, true);
                }
                catch (Exception e)
                {
                    System.Windows.MessageBox.Show(e.Message);
                    return null;
                }

            }
        }

        /// <summary>
        /// 打开CSV 文件
        /// </summary>
        /// <param name="fileName">文件全名</param>
        /// <param name="firstRow">开始行</param>
        /// <param name="firstColumn">开始列</param>
        /// <param name="getRows">获取多少行</param>
        /// <param name="getColumns">获取多少列</param>
        /// <param name="haveTitleRow">是有标题行</param>
        /// <returns>DataTable</returns>
        public static DataTable OpenCSV(string fullFileName, Int16 firstRow = 0, Int16 firstColumn = 0, Int16 getRows = 0, Int16 getColumns = 0, bool haveTitleRow = true)
        {
            DataTable dt = new DataTable();
            FileStream fs = new FileStream(fullFileName, System.IO.FileMode.Open, System.IO.FileAccess.Read);
            StreamReader sr = new StreamReader(fs, System.Text.Encoding.Default);
            //记录每次读取的一行记录
            string strLine = "";
            //记录每行记录中的各字段内容
            string[] aryLine;
            //标示列数
            int columnCount = 0;
            //是否已建立了表的字段
            bool bCreateTableColumns = false;
            //第几行
            int iRow = 1;

            //去除无用行
            if (firstRow > 0)
            {
                for (int i = 1; i < firstRow; i++)
                {
                    sr.ReadLine();
                }
            }

            // { ",", ".", "!", "?", ";", ":", " " };
            string[] separators = { "," };
            //逐行读取CSV中的数据
            while ((strLine = sr.ReadLine()) != null)
            {
                strLine = strLine.Trim();
                aryLine = strLine.Split(separators, System.StringSplitOptions.RemoveEmptyEntries);

                if (bCreateTableColumns == false)
                {
                    bCreateTableColumns = true;
                    columnCount = aryLine.Length;
                    //创建列
                    for (int i = firstColumn; i < (getColumns == 0 ? columnCount : firstColumn + getColumns); i++)
                    {
                        DataColumn dc
                            = new DataColumn(haveTitleRow == true ? aryLine[i] : "COL" + i.ToString());
                        dt.Columns.Add(dc);
                    }

                    bCreateTableColumns = true;

                    if (haveTitleRow == true)
                    {
                        continue;
                    }
                }


                DataRow dr = dt.NewRow();
                //for (int j = firstColumn; j < (getColumns == 0 ? columnCount : firstColumn + getColumns); j++)
                //{
                //    dr[j - firstColumn] = aryLine[j];
                //}
                for (int j = firstColumn; j < aryLine.Length; j++)
                {
                    dr[j - firstColumn] = aryLine[j];
                }
                dt.Rows.Add(dr);

                iRow = iRow + 1;
                if (getRows > 0)
                {
                    if (iRow > getRows)
                    {
                        break;
                    }
                }

            }

            sr.Close();
            fs.Close();
            return dt;
        }
        public static DataTable ImportCSV(string filePath)
        {
            DataTable ds = new DataTable();
            using (StreamReader sw = new StreamReader(filePath, Encoding.Default))
            {
                string str = sw.ReadLine();
                string[] columns = str.Split(',');
                foreach (string name in columns)
                {
                    ds.Columns.Add(name);
                }
                string line = sw.ReadLine();
                while (line != null)
                {
                    string[] data = line.Split(',');
                    ds.Rows.Add(data);
                    line = sw.ReadLine();
                }
                sw.Close();
            }
            return ds;
        }

       
//导出数据到csv文件中。（这种方式速度快）   
        public static void ExportToCSV(DataTable ds, string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.Unicode))
            {
                for (int i = 0; i < ds.Columns.Count; i++)
                {
                    sw.Write(ds.Columns[i].ColumnName + "\t");
                }
                sw.Write("\n");
                for (int i = 0; i < ds.Rows.Count; i++)
                {
                    for (int j = 0; j < ds.Columns.Count; j++)
                    {
                        sw.Write(ds.Rows[i][j].ToString() + "\t");
                    }
                    sw.Write("\n");
                }
                sw.Close();
            }
            return;
        }
    }

}
