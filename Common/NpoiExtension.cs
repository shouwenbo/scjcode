using NPOI.SS.UserModel;

namespace Common
{
    public static class NpoiExtension
    {
        public static IWorkbook GetNpoiWorkbook(this string path, out string err, params string[] passwords)
        {
            IWorkbook workbook = null;
            err = "";
            try
            {
                workbook = WorkbookFactory.Create(path);
            }
            catch (Exception e1)
            {
                err = e1.Message;
                foreach (var password in passwords)
                {
                    try
                    {
                        workbook = WorkbookFactory.Create(path, password);
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
            return workbook;
        }

        public static void SetCellValueIfNotEmpty(this ICell cell, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                if (int.TryParse(value, out int numberValue))
                {
                    cell.SetCellValue(numberValue);
                }
                else
                {
                    cell.SetCellValue(value);
                }
            }
        }

        public static string GetCellStringValue(this ICell cell)
        {
            if (cell == null)
            {
                return "";
            }

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue ?? "";

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        if (cell.DateCellValue.HasValue)
                        {
                            return cell.DateCellValue.Value.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            return "";
                        }
                    }
                    return cell.NumericCellValue.ToString();

                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();

                case CellType.Formula:
                    try
                    {
                        IFormulaEvaluator evaluator = cell.Sheet.Workbook.GetCreationHelper().CreateFormulaEvaluator();
                        var evaluatedValue = evaluator.Evaluate(cell);

                        if (evaluatedValue.CellType == CellType.Numeric)
                        {
                            return evaluatedValue.NumberValue.ToString();
                        }
                        else if (evaluatedValue.CellType == CellType.String)
                        {
                            return evaluatedValue.StringValue;
                        }
                        else if (evaluatedValue.CellType == CellType.Boolean)
                        {
                            return evaluatedValue.BooleanValue.ToString();
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                    catch
                    {
                        return cell.CellFormula ?? string.Empty;
                    }

                case CellType.Blank:
                    return "";

                default:
                    return cell.ToString() ?? "";
            }
        }

        /// <summary>
        /// 删除指定的多行并上移剩余行。
        /// </summary>
        /// <param name="sheet">工作表对象</param>
        /// <param name="rowIndices">要删除的行索引集合</param>
        public static void DeleteRowsAndShiftUp(this ISheet sheet, IEnumerable<int> rowIndices)
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
            int lastRowNum = sheet.LastRowNum;

            // 从最后一行开始删除，避免索引错乱
            for (int i = sortedRowIndices.Count - 1; i >= 0; i--)
            {
                int rowIndex = sortedRowIndices[i];
                if (rowIndex < 0 || rowIndex > lastRowNum)
                {
                    throw new ArgumentOutOfRangeException(nameof(rowIndices), $"行索引 {rowIndex} 超出范围");
                }

                // 上移行（从下往上处理，避免影响索引）
                if (rowIndex < lastRowNum)
                {
                    sheet.ShiftRows(rowIndex + 1, lastRowNum, -1);
                }

                // 删除最后一行
                IRow lastRow = sheet.GetRow(lastRowNum);
                if (lastRow != null)
                {
                    sheet.RemoveRow(lastRow);
                }

                lastRowNum--;
            }
        }

        /// <summary>
        /// 删除指定的多行并上移剩余行
        /// </summary>
        /// <param name="sheet">工作表对象</param>
        /// <param name="startRow">开始行</param>
        /// <param name="endRow">结束行</param>
        public static void DeleteRowRangeAndShiftUp(this ISheet sheet, int startRow, int endRow)
        {
            if (sheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }
            
            if (startRow < 0 || endRow < 0 || startRow > endRow || endRow > sheet.LastRowNum)
            {
                throw new ArgumentOutOfRangeException(nameof(startRow), "行索引范围无效");
            }
            
            // 获取最后一行的行号
            int lastRowNum = sheet.LastRowNum;
            
            // 先将需要删除的行范围内的行移除
            for (int i = startRow; i <= endRow; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row != null)
                {
                    sheet.RemoveRow(row);
                }
            }
            
            // 上移剩余行
            if (endRow < lastRowNum)
            {
                sheet.ShiftRows(endRow + 1, lastRowNum, -(endRow - startRow + 1));
            }
        }
    }
}
