using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.Formula.Eval;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;

namespace U8Voucher.Common.CommExcelOpt
{
    public class ExcelOpt
    {
        /// <summary>
        /// 将指定sheet中的数据导入到datatable中
        /// </summary>
        /// <param name="sheet">指定需要导出的sheet</param>
        /// <param name="HeaderRowIndex">列头所在的行号，-1没有列头</param>
        /// <param name="needHeader"></param>
        /// <returns></returns>
        public static DataTable Export2DataTable(ISheet sheet, int HeaderRowIndex, bool needHeader, bool isVoucherGenerate = true, string FileExt = ".xlsx")
        {
            var dt = new DataTable();
            if (".xlsx".Equals(FileExt))
            {
                XSSFRow headerRow = null;
                int cellCount;
                try
                {
                    if (HeaderRowIndex < 0 || !needHeader)
                    {
                        headerRow = sheet.GetRow(0) as XSSFRow;
                        cellCount = headerRow.LastCellNum;
                        for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                        {
                            var column = new DataColumn(Convert.ToString(i));
                            dt.Columns.Add(column);
                        }
                    }
                    else
                    {
                        headerRow = sheet.GetRow(HeaderRowIndex) as XSSFRow;
                        cellCount = headerRow.LastCellNum;
                        for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                        {
                            var cell = headerRow.GetCell(i);
                            if (cell == null)
                            {
                                break;
                            }
                            else
                            {
                                DataColumn column = null;
                                if (i == 2)
                                {
                                    if (isVoucherGenerate)
                                        column = new DataColumn(headerRow.GetCell(i).ToString(), Type.GetType("System.Double"));
                                    else
                                        column = new DataColumn(headerRow.GetCell(i).ToString());
                                }
                                else
                                {
                                    column = new DataColumn(headerRow.GetCell(i).ToString());
                                }
                                dt.Columns.Add(column);
                            }
                        }
                    }
                    for (var i = HeaderRowIndex + 1; i <= sheet.LastRowNum; i++)
                    {
                        XSSFRow row = null;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i) as XSSFRow;
                        }
                        else
                        {
                            row = sheet.GetRow(i) as XSSFRow;
                        }
                        var dtRow = dt.NewRow();
                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                switch (row.GetCell(j).CellType)
                                {
                                    case CellType.Boolean:
                                        dtRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        dtRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                        break;
                                    case CellType.Formula:
                                        switch (row.GetCell(j).CachedFormulaResultType)
                                        {
                                            case CellType.Boolean:
                                                dtRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);

                                                break;
                                            case CellType.Error:
                                                dtRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);

                                                break;
                                            case CellType.Numeric:
                                                dtRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);

                                                break;
                                            case CellType.String:
                                                var strFORMULA = row.GetCell(j).StringCellValue;
                                                if (strFORMULA != null && strFORMULA.Length > 0)
                                                {
                                                    dtRow[j] = strFORMULA.ToString();
                                                }
                                                else
                                                {
                                                    dtRow[j] = null;
                                                }
                                                break;
                                            default:
                                                dtRow[j] = string.Empty;
                                                break;
                                        }
                                        break;
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                        {
                                            dtRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                        }
                                        else
                                        {
                                            dtRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                        }
                                        break;
                                    case CellType.String:
                                        var str = row.GetCell(j).StringCellValue;
                                        if (!string.IsNullOrEmpty(str))
                                        {
                                            dtRow[j] = Convert.ToString(str);
                                        }
                                        else
                                        {
                                            dtRow[j] = null;
                                        }
                                        break;
                                    default:
                                        dtRow[j] = string.Empty;
                                        break;
                                }
                            }
                        }
                        dt.Rows.Add(dtRow);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            else
            {
                HSSFRow headerRow = null;
                int cellCount;
                try
                {
                    if (HeaderRowIndex < 0 || !needHeader)
                    {
                        headerRow = sheet.GetRow(0) as HSSFRow;
                        cellCount = headerRow.LastCellNum;
                        for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                        {
                            var column = new DataColumn(Convert.ToString(i));
                            dt.Columns.Add(column);
                        }
                    }
                    else
                    {
                        headerRow = sheet.GetRow(HeaderRowIndex) as HSSFRow;
                        cellCount = headerRow.LastCellNum;
                        for (int i = headerRow.FirstCellNum; i <= cellCount; i++)
                        {
                            var cell = headerRow.GetCell(i);
                            if (cell == null)
                            {
                                break;
                            }
                            else
                            {
                                DataColumn column = null;
                                if (i == 2)
                                {
                                    if (isVoucherGenerate)
                                        column = new DataColumn(headerRow.GetCell(i).ToString(), Type.GetType("System.Double"));
                                    else
                                        column = new DataColumn(headerRow.GetCell(i).ToString());
                                }
                                else
                                {
                                    column = new DataColumn(headerRow.GetCell(i).ToString());
                                }
                                dt.Columns.Add(column);
                            }
                        }
                    }
                    for (var i = HeaderRowIndex + 1; i <= sheet.LastRowNum; i++)
                    {
                        HSSFRow row = null;
                        if (sheet.GetRow(i) == null)
                        {
                            row = sheet.CreateRow(i) as HSSFRow;
                        }
                        else
                        {
                            row = sheet.GetRow(i) as HSSFRow;
                        }
                        var dtRow = dt.NewRow();
                        for (int j = row.FirstCellNum; j <= cellCount; j++)
                        {
                            if (row.GetCell(j) != null)
                            {
                                switch (row.GetCell(j).CellType)
                                {
                                    case CellType.Boolean:
                                        dtRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);
                                        break;
                                    case CellType.Error:
                                        dtRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);
                                        break;
                                    case CellType.Formula:
                                        switch (row.GetCell(j).CachedFormulaResultType)
                                        {
                                            case CellType.Boolean:
                                                dtRow[j] = Convert.ToString(row.GetCell(j).BooleanCellValue);

                                                break;
                                            case CellType.Error:
                                                dtRow[j] = ErrorEval.GetText(row.GetCell(j).ErrorCellValue);

                                                break;
                                            case CellType.Numeric:
                                                dtRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);

                                                break;
                                            case CellType.String:
                                                var strFORMULA = row.GetCell(j).StringCellValue;
                                                if (strFORMULA != null && strFORMULA.Length > 0)
                                                {
                                                    dtRow[j] = strFORMULA.ToString();
                                                }
                                                else
                                                {
                                                    dtRow[j] = null;
                                                }
                                                break;
                                            default:
                                                dtRow[j] = string.Empty;
                                                break;
                                        }
                                        break;
                                    case CellType.Numeric:
                                        if (DateUtil.IsCellDateFormatted(row.GetCell(j)))
                                        {
                                            dtRow[j] = DateTime.FromOADate(row.GetCell(j).NumericCellValue);
                                        }
                                        else
                                        {
                                            dtRow[j] = Convert.ToDouble(row.GetCell(j).NumericCellValue);
                                        }
                                        break;
                                    case CellType.String:
                                        var str = row.GetCell(j).StringCellValue;
                                        if (!string.IsNullOrEmpty(str))
                                        {
                                            dtRow[j] = Convert.ToString(str);
                                        }
                                        else
                                        {
                                            dtRow[j] = null;
                                        }
                                        break;
                                    default:
                                        dtRow[j] = string.Empty;
                                        break;
                                }
                            }
                        }
                        dt.Rows.Add(dtRow);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return dt;
        }

        /// <summary>
        /// 将DataTable中的数据导入Excel文件中
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="file"></param>
        public static void DataTable2Excel(DataSet ds, string filepath, string sheetName)
        {
            try
            {
                string FileExt = Path.GetExtension(filepath);
                if (".xlsx".Equals(FileExt))
                {
                    XSSFWorkbook workbook = new XSSFWorkbook();
                    foreach (DataTable table in ds.Tables)
                    {
                        ISheet sheet = workbook.CreateSheet(sheetName);
                        IRow headerRow = sheet.CreateRow(0);
                        foreach (DataColumn column in table.Columns) headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
                        int rowIndex = 1;
                        foreach (DataRow row in table.Rows)
                        {
                            IRow dataRow = sheet.CreateRow(rowIndex);
                            foreach (DataColumn column in table.Columns) dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                            rowIndex++;
                        }
                    }
                    using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fs);
                        fs.Close();
                    }
                }
                else
                {
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    foreach (DataTable table in ds.Tables)
                    {
                        ISheet sheet = workbook.CreateSheet(sheetName);
                        IRow headerRow = sheet.CreateRow(0);
                        foreach (DataColumn column in table.Columns) headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);
                        int rowIndex = 1;
                        foreach (DataRow row in table.Rows)
                        {
                            IRow dataRow = sheet.CreateRow(rowIndex);
                            foreach (DataColumn column in table.Columns) dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                            rowIndex++;
                        }
                    }
                    using (FileStream fs = new FileStream(filepath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fs);
                        fs.Close();
                    }
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
