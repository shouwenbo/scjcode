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
    }
}
