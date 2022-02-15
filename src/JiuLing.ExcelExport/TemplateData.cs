using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using JiuLing.ExcelExport.Items;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace JiuLing.ExcelExport
{
    /// <summary>
    /// 模板数据模式
    /// </summary>
    public class TemplateData
    {
        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="templateFile">模板文件的路径</param>
        /// <param name="destinationFile">待导出文件的路径</param>
        /// <param name="data"></param>
        /// <exception cref="FileNotFoundException">文件不存在</exception>
        /// <exception cref="FileLoadException">文件加载异常</exception>
        /// <exception cref="ArgumentException">参数异常</exception>
        public void Export(string templateFile, string destinationFile, DataSet data)
        {
            if (!File.Exists(templateFile))
            {
                throw new FileNotFoundException($"模板文件不存在：{templateFile}");
            }

            var destinationDirectory = Path.GetDirectoryName(destinationFile) ?? throw new ArgumentException($"文件路径不合法：{destinationFile}");
            if (!Directory.Exists(destinationDirectory))
            {
                Directory.CreateDirectory(destinationDirectory);
            }


            IWorkbook workbook;
            using (FileStream fs = new FileStream(templateFile, FileMode.Open, FileAccess.Read))
            {
                if (destinationFile.IndexOf(".xlsx") > 0)
                {
                    //07版本
                    workbook = new XSSFWorkbook(fs);
                }
                else if (destinationFile.IndexOf(".xls") > 0)
                {
                    //03版本  
                    workbook = new HSSFWorkbook(fs);
                }
                else
                {
                    throw new FileLoadException($"不支持的文件版本：{destinationFile}");
                }

                using (FileStream fsDestination = new FileStream(destinationFile, FileMode.Create, FileAccess.ReadWrite))
                {
                    int sheetCount = workbook.NumberOfSheets;
                    for (int index = 0; index < sheetCount; index++)
                    {
                        ISheet sheet = workbook.GetSheetAt(index);
                        WriteSheet(sheet, data);
                    }
                    workbook.Write(fsDestination);
                }
            }
        }

        private void WriteSheet(ISheet sheet, DataSet data)
        {
            //思路：从第一行第一列开始检查，匹配到模板时则进行绑定，然后继续检查后面的单元格
            int rowIndex = -1;
            while (true)
            {
                rowIndex += 1;
                if (rowIndex > sheet.LastRowNum)
                {
                    return;
                }
                IRow row = sheet.GetRow(rowIndex);
                if (row == null)
                {
                    continue;
                }

                //逐列扫描
                for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                {
                    ICell cell = row.GetCell(colIndex);
                    if (cell == null || cell.CellType != CellType.String)
                    {
                        //非字符串格式的单元格，认为不是模板值
                        continue;
                    }

                    string cellValue = cell.StringCellValue;
                    CellBindingInfo bindingInfo = TemplateUtils.GetCellBindingInfo(cellValue);

                    if (bindingInfo.BindingType == BindingTypeEnum.NotTemplate)
                    {
                        continue;
                    }
                    else if (bindingInfo.BindingType == BindingTypeEnum.Cell)
                    {
                        WriteOneCell(cell, data, bindingInfo.TableName, bindingInfo.ColumnName);
                        continue;
                    }
                    else if (bindingInfo.BindingType == BindingTypeEnum.List)
                    {
                        var insertRowCount = WriteListCells(sheet, row, colIndex, data, bindingInfo.TableName);
                        //列表绑定后，需要将原有的行索引移动到新增行之后
                        rowIndex += insertRowCount;

                        //如果发现有列表绑定，则认为后面的单元格全部是列表形式，跳过
                        break;
                    }
                    else
                    {
                        throw new ArgumentException($"无法识别的绑定格式，行：{rowIndex}，列：{colIndex}");
                    }
                }
            }
        }

        /// <summary>
        /// 写入列表形式的绑定
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="startColIndex"></param>
        /// <param name="data"></param>
        /// <param name="tableName"></param>
        /// <returns>返回新增列表的总行数</returns>
        /// <exception cref="ArgumentException"></exception>
        private int WriteListCells(ISheet sheet, IRow row, int startColIndex, DataSet data, string tableName)
        {
            var dt = data.Tables[tableName];
            if (dt == null)
            {
                throw new ArgumentException($"数据源中不包含{tableName}数据表");
            }

            var bindingMap = new Dictionary<string, int>();
            for (int colIndex = startColIndex; colIndex < row.LastCellNum; colIndex++)
            {
                ICell cell = row.GetCell(colIndex);
                if (cell == null || cell.CellType != CellType.String)
                {
                    continue;
                }

                var columnName = TemplateUtils.GetCellBindingColumnName(cell.StringCellValue, tableName);
                if (string.IsNullOrEmpty(columnName))
                {
                    continue;
                }

                var column = dt.Columns[columnName];
                if (column == null)
                {
                    throw new ArgumentException($"不存在的绑定：数据表{tableName}，列{columnName}");
                }
                bindingMap.Add(columnName, colIndex);
            }

            if (dt.Rows.Count == 0)
            {
                foreach (var bindingItem in bindingMap)
                {
                    var column = dt.Columns[bindingItem.Key];
                    if (column == null)
                    {
                        throw new ArgumentException($"不存在的绑定：数据表{tableName}，列{bindingItem.Key}");
                    }
                    SetCellValue(row.GetCell(bindingItem.Value), column.DataType, "");
                }
            }
            else
            {
                int startRowIndex = row.RowNum;
                for (int rowIndex = 0; rowIndex < dt.Rows.Count; rowIndex++)
                {
                    int targetIndex = startRowIndex + rowIndex;
                    if (rowIndex > 0)
                    {
                        sheet.CopyRow(startRowIndex, targetIndex);
                    }


                    IRow newRow = sheet.GetRow(targetIndex);
                    foreach (var bindingItem in bindingMap)
                    {
                        var column = dt.Columns[bindingItem.Key];
                        if (column == null)
                        {
                            throw new ArgumentException($"不存在的绑定：数据表{tableName}，列{bindingItem.Key}");
                        }

                        object value = dt.Rows[rowIndex][bindingItem.Key];
                        SetCellValue(newRow.GetCell(bindingItem.Value), column.DataType, value);

                    }
                }
            }
            return dt.Rows.Count;
        }
        /// <summary>
        /// 写入单元格形式的绑定
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="data"></param>
        /// <param name="tableName"></param>
        /// <param name="columnName"></param>
        /// <exception cref="ArgumentException"></exception>
        private void WriteOneCell(ICell cell, DataSet data, string tableName, string columnName)
        {
            var dt = data.Tables[tableName];
            if (dt == null)
            {
                throw new ArgumentException($"数据源中不包含{tableName}数据表");
            }

            var column = dt.Columns[columnName];
            if (column == null)
            {
                throw new ArgumentException($"不存在的绑定：数据表{tableName}，列{columnName}");
            }

            var value = dt.Rows[0][columnName];
            SetCellValue(cell, column.DataType, value);
        }

        private static void SetCellValue(ICell cell, Type type, object value)
        {
            switch (type.FullName)
            {
                case "System.String":
                    cell.SetCellValue(value.ToString());
                    cell.SetCellType(CellType.String);
                    break;
                case "System.Int16":
                case "System.Int32":
                case "System.Int64":
                case "System.Decimal":
                    if (!double.TryParse(value.ToString(), out var d))
                    {
                        d = 0;
                    }
                    cell.SetCellValue(d);
                    cell.SetCellType(CellType.Numeric);
                    break;
                case "System.Boolean":
                    if (!bool.TryParse(value.ToString(), out var b))
                    {
                        b = false;
                    }
                    cell.SetCellValue(b);
                    cell.SetCellType(CellType.Boolean);
                    break;
                case "System.DateTime":
                    if (!DateTime.TryParse(value.ToString(), out var time))
                    {
                        time = DateTime.MinValue;
                    }
                    string t = time.ToString("yyyy-MM-dd HH:mm:ss");
                    cell.SetCellValue(t);
                    cell.SetCellType(CellType.String);
                    break;
                default:
                    throw new ArgumentException($"不支持的数据格式：{type.FullName}，行：{cell.RowIndex}，列：{cell.ColumnIndex}");
            }
        }
    }
}
