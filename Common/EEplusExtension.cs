using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;

namespace Common
{
    public static class EPPlusExtension
    {
        // 获取工作簿（EPPlus）
        public static ExcelPackage GetExcelPackage(this string path, out string err, params string[] passwords)
        {
            err = "";
            try
            {
                var fileInfo = new FileInfo(path);
                var package = new ExcelPackage(fileInfo);
                return package;
            }
            catch (Exception e)
            {
                err = e.Message;
                // 如果有密码，尝试使用密码打开
                foreach (var password in passwords)
                {
                    try
                    {
                        var fileInfo = new FileInfo(path);
                        var package = new ExcelPackage(fileInfo, password);
                        return package;
                    }
                    catch (Exception e2)
                    {
                        err = e2.Message;
                    }
                }
            }
            if (err.Contains("process"))
            {
                err = "文件可能被占用，请先关闭文件再尝试操作";
            }
            return null;
        }

        // 设置单元格值（根据值类型）
        public static void SetCellValueIfNotEmpty(this ExcelRange cell, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                if (int.TryParse(value, out int numberValue))
                {
                    cell.Value = numberValue;
                }
                else
                {
                    cell.Value = value;
                }
            }
        }

        public static void SetCell(this ExcelRange cell, ExcelRange newCell)
        {
            if (newCell != null)
            {
                if (!string.IsNullOrEmpty(newCell.Text))
                {
                    if (int.TryParse(newCell.Text, out int numberValue))
                    {
                        cell.Value = numberValue;
                    }
                    else
                    {
                        cell.Value = newCell.Text;
                    }
                }

                if (newCell.Comment != null)
                {
                    cell.AddComment(newCell.Comment.Text);

                    if (newCell.Comment.LineColor != default && newCell.Comment.LineColor.ToArgb() != 0)
                    {
                        cell.Comment.LineColor = newCell.Comment.LineColor;
                    }
                    cell.Comment.LineStyle = newCell.Comment.LineStyle;
                    cell.Comment.LineWidth = newCell.Comment.LineWidth;

                    #region 错误日志

                    // var builder = new StringBuilder();
                    // builder.AppendLine($"cell.Text: {cell.Text}");
                    // builder.AppendLine($"cell.Comment.Text: {cell.Comment.Text}");
                    // builder.AppendLine($"cell.Comment.LineColor: {cell.Comment.LineColor.ToLogString()}");
                    // builder.AppendLine($"newCell.Text: {newCell.Text}");
                    // builder.AppendLine($"newCell.Comment.Text: {newCell.Comment.Text}");
                    // builder.AppendLine($"newCell.Comment.LineColor: {newCell.Comment.LineColor.ToLogString()}");
                    // 
                    // 
                    // try
                    // {
                    //     cell.Comment.LineColor = newCell.Comment.LineColor;
                    //     cell.Comment.LineStyle = newCell.Comment.LineStyle;
                    //     cell.Comment.LineWidth = newCell.Comment.LineWidth;
                    // }
                    // catch (Exception ex)
                    // {
                    //     builder.AppendLine();
                    //     builder.AppendLine(ex.ToLogString());
                    //     string folderPath = "DebugLog";
                    //     string filePath = Path.Combine(folderPath, $"EPPlusExtension.{cell.Start.Row}.{cell.Start.Column}.Error.txt");
                    //     Directory.CreateDirectory(folderPath);
                    //     File.WriteAllText(filePath, builder.ToString());
                    //     throw;
                    // }

                    #endregion
                }
            }

            // 复制纯色填充的背景色
            if (newCell.Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.Solid)
            {
                string backgroundRgbValue = newCell.Style.Fill.BackgroundColor.Rgb;
                if (backgroundRgbValue.Length >= 6)
                {
                    cell.Style.Fill.PatternType = newCell.Style.Fill.PatternType;
                    string backgroundHexColor = "#" + backgroundRgbValue.Substring(backgroundRgbValue.Length - 6);
                    var backgroundColor = ColorTranslator.FromHtml(backgroundHexColor);
                    cell.Style.Fill.SetBackground(backgroundColor);
                }
            }

            string fontRgbValue = newCell.Style.Font.Color.Rgb;
            if (fontRgbValue != null && fontRgbValue.Length >= 6)
            {
                string fontHexColor = "#" + fontRgbValue.Substring(fontRgbValue.Length - 6);
                var fontColor = ColorTranslator.FromHtml(fontHexColor);
                cell.Style.Font.Color.SetColor(fontColor);
            }
        }

        public static void MarkYellow(this ExcelRange cell)
        {
            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell.Style.Fill.SetBackground(ColorTranslator.FromHtml("#FFFF00"));
        }

        public static void MarkGreen(this ExcelRange cell)
        {
            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            cell.Style.Fill.SetBackground(ColorTranslator.FromHtml("#00FF00"));
        }

        public static void SetCell(this ExcelRange cell, string value, bool allowEmpty = false)
        {
            if (!string.IsNullOrEmpty(value) || allowEmpty)
            {
                // "0"要按照字符进行处理
                if (int.TryParse(value, out int numberValue) && value != "0")
                {
                    cell.Value = numberValue;
                }
                else
                {
                    cell.Value = value;
                }
            }
        }

        public static void SetCell(this ExcelRange cell, int value)
        {
            cell.Value = value;
        }

        // 获取单元格的字符串值
        public static string GetCellStringValue(this ExcelRange cell)
        {
            if (cell == null)
            {
                return "";
            }

            if (cell.Value == null)
                return "";

            // 根据单元格的类型返回不同的值
            switch (cell.Value)
            {
                case string strValue:
                    return strValue;
                case double numValue:
                    return numValue.ToString();
                case DateTime dateValue:
                    return dateValue.ToString("yyyy-MM-dd");
                case bool boolValue:
                    return boolValue.ToString();
                default:
                    return cell.Value.ToString();
            }
        }

        // 删除指定的多行并上移剩余行
        public static void DeleteRowsAndShiftUp(this ExcelWorksheet sheet, IEnumerable<int> rowIndices)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (rowIndices == null || !rowIndices.Any())
            {
                throw new ArgumentException("行索引集合不能为空", nameof(rowIndices));
            }

            // 获取有效行索引，去重并按从小到大排序
            var sortedRowIndices = rowIndices.Distinct().OrderBy(x => x).ToList();
            int lastRowNum = sheet.Dimension.End.Row;

            // 从最后一行开始删除，避免索引错乱
            for (int i = sortedRowIndices.Count - 1; i >= 0; i--)
            {
                int rowIndex = sortedRowIndices[i];
                if (rowIndex < 1 || rowIndex > lastRowNum)
                {
                    throw new ArgumentOutOfRangeException(nameof(rowIndices), $"行索引 {rowIndex} 超出范围");
                }

                // 删除最后一行
                sheet.DeleteRow(lastRowNum);
                lastRowNum--;
            }
        }

        // 删除指定的行范围并上移剩余行
        public static void DeleteRowRangeAndShiftUp(this ExcelWorksheet sheet, int startRow, int endRow)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (startRow < 1 || endRow < 1 || startRow > endRow || endRow > sheet.Dimension.End.Row)
            {
                throw new ArgumentOutOfRangeException(nameof(startRow), "行索引范围无效");
            }
            
            sheet.DeleteRow(startRow, endRow);
        }
    }
}
