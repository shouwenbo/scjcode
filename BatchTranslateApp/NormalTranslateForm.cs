using ICSharpCode.SharpZipLib.Zip;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using RestSharp;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Common;
using Microsoft.Extensions.FileSystemGlobbing.Internal;

namespace BatchTranslateApp
{
    public partial class NormalTranslateForm : ExcelForm
    {
        private int allCount = 0;
        private int currentCount = 0;

        public NormalTranslateForm()
        {
            InitializeComponent();
            InitExcelSelector(this.gbx_file, this.btn_select_file, file =>
            {
                var list = LoadExcel(file);
                if (list != null && list.Count > 0)
                {
                    BatchModeCheck(list);
                }
                else
                {
                    this.txt_check_result.Text = "读取不到表格";
                }
            });
            CheckForIllegalCrossThreadCalls = false;
        }

        [Obsolete]
        [STAThread]
        public void Translate(string[] chineses, string join)
        {
            this.gbx_file.Visible = false;
            // this.cbx_ignore_ssl.Visible = false;
            currentCount = 0;
            var list = new List<string>();

            Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        allCount = chineses.Length;
                        var taskBuilder = new StringBuilder();

                        #region 错误检查

                        foreach (string chinese in chineses)
                        {

                            string pattern = @"[\u4e00-\u9fa5]+";
                            MatchCollection matches = Regex.Matches(chinese, pattern);
                            var errList = new List<string>();
                            foreach (System.Text.RegularExpressions.Match match in matches)
                            {
                                string chineseWords = match.Value;
                                if (chineseWords.Length > 10)
                                {
                                    errList.Add($"‘{chineseWords}’ 长度超过限制，每个词语最多10个汉字");
                                }
                            }

                            if (errList.Count > 0)
                            {
                                this.txt_check_result.Text = string.Join("\r\n", errList);
                                return;
                            }
                        }

                        #endregion

                        foreach (string chinese in chineses)
                        {
                            if (!string.IsNullOrWhiteSpace(chinese))
                            {
                                var korean = TranslateHelper.SentenceToKorean(chinese, true /*this.cbx_ignore_ssl.Checked*/);
                                if (string.IsNullOrEmpty(korean))
                                {
                                    this.txt_check_result.Text = $"‘{chinese}’翻译失败，出现这种情况可能是因为翻译太频繁或者一次性翻译的内容太多，请稍后再尝试或者一次少翻译一点";
                                    this.gbx_file.Visible = true;
                                    // this.cbx_ignore_ssl.Visible = true;
                                    return;
                                }
                                list.Add(korean ?? "");
                            }
                            else
                            {
                                list.Add("");
                            }
                            currentCount++;
                            this.lbl_task.Text = $"{currentCount}/{allCount}";
                        }

                        // ClipboardHelper.SetText(string.Join(join, list));
                        var ccc = string.Join("\n", list).Trim();
                        Clipboard.Clear();
                        Clipboard.SetDataObject(string.Join("\n", list).Trim());
                        // Clipboard.SetDataObject("啊啊啊啊33");
                        this.gbx_file.Visible = true;
                        // this.cbx_ignore_ssl.Visible = true;
                        this.txt_check_result.Text = "已复制到剪贴板，直接在表格中粘贴即可";
                        MessageBox.Show($"翻译完成，直接粘贴即可");
                    }
                    catch (Exception ex)
                    {
                        this.gbx_file.Visible = true;
                        // this.cbx_ignore_ssl.Visible = true;
                        this.txt_check_result.Text = ex.Message;
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });
        }

        [Obsolete]
        private void ParseData(string text)
        {
            if (text.Contains('\t')) // 横排
            {
                string[] lines = text.Split('\n');

                var list = new List<string>();

                foreach (string line in lines)
                {
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    string[] cells = line.Split('\t');

                    if (cells.Length != 1)
                    {
                        MessageBox.Show("格式有误，请重新复制！");
                    }

                    Translate(cells, "\t");
                }
            }
            else // 竖排
            {
                string[] lines = text.Split('\n');

                var list = new List<string>();

                Translate(lines, "\n");
            }
        }

        [Obsolete]
        private void btn_translatea_Click(object sender, EventArgs e)
        {
            string clipboardText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(clipboardText))
            {
                ParseData(clipboardText);
            }
            else
            {
                MessageBox.Show("剪贴板是空的！");
            }
        }

        public void BatchModeOutput(List<List<TranslateItem>> list)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // 创建样式
            ICellStyle styleBlue = workbook.CreateCellStyle();
            styleBlue.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            styleBlue.VerticalAlignment = VerticalAlignment.Center;
            IFont fontBold = workbook.CreateFont();
            fontBold.IsBold = true;
            styleBlue.SetFont(fontBold);
            styleBlue.FillForegroundColor = IndexedColors.LightBlue.Index;
            styleBlue.FillPattern = FillPattern.SolidForeground;

            ICellStyle styleGreen = workbook.CreateCellStyle();
            styleGreen.CloneStyleFrom(styleBlue);
            styleGreen.FillForegroundColor = IndexedColors.LightGreen.Index;

            // 设置第一行
            IRow row = sheet.CreateRow(0);
            var headers = new List<string>() { };
            for (int i = 0; i < list.Count; i++)
            {
                headers.Add("中文");
                headers.Add("韩文");
            }
            for (int i = 0; i < headers.Count; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = (i % 2 == 0) ? styleBlue : styleGreen;
            }

            // 插入数据
            for (int i = 0; i < list.Count; i++)
            {
                List<TranslateItem> translateItems = list[i];
                for (int j = 0; j < translateItems.Count; j++)
                {
                    IRow newRow = sheet.GetRow(j + 1) ?? sheet.CreateRow(j + 1);

                    ICell cellChinese = newRow.CreateCell(i * 2); // 中文列
                    cellChinese.SetCellValue(translateItems[j].Chinese);

                    ICell cellKorean = newRow.CreateCell(i * 2 + 1); // 韩文列
                    cellKorean.SetCellValue(translateItems[j].Korean);
                }
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "保存 Excel 文件";
                saveFileDialog.FileName = "翻译结果.xlsx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string targetFilePath = saveFileDialog.FileName;
                    using (FileStream stream = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(stream);
                    }
                    this.txt_check_result.Text = $"文件已成功保存到：{targetFilePath}";
                    Process.Start(new ProcessStartInfo(targetFilePath) { UseShellExecute = true });
                }
            }
        }

        public void BatchModeTranslate(List<List<string>> list)
        {
            this.gbx_file.Visible = false;
            // this.cbx_ignore_ssl.Visible = false;
            currentCount = 0;
            var resultList = new List<List<TranslateItem>>();

            Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        allCount = list.Sum(p => p.Count);
                        var taskBuilder = new StringBuilder();

                        foreach (var chineses in list)
                        {
                            var currentItemList = new List<TranslateItem>();

                            foreach (string chinese in chineses)
                            {
                                if (!string.IsNullOrWhiteSpace(chinese))
                                {
                                    var korean = TranslateHelper.SentenceToKorean(chinese, true /*this.cbx_ignore_ssl.Checked*/);
                                    if (string.IsNullOrEmpty(korean))
                                    {
                                        this.txt_check_result.Text = $"‘{chinese}’翻译失败，出现这种情况可能是因为翻译太频繁或者一次性翻译的内容太多，请稍后再尝试或者一次少翻译一点";
                                        this.gbx_file.Visible = true;
                                        // this.cbx_ignore_ssl.Visible = true;
                                        return;
                                    }
                                    currentItemList.Add(new TranslateItem()
                                    {
                                        Chinese = chinese,
                                        Korean = korean ?? ""
                                    });
                                }
                                else
                                {
                                    currentItemList.Add(new TranslateItem()
                                    {
                                        Chinese = chinese,
                                        Korean = ""
                                    });
                                }
                                currentCount++;
                                this.lbl_task.Text = $"{currentCount}/{allCount}";
                            }

                            resultList.Add(currentItemList);
                        }

                        this.gbx_file.Visible = true;
                        // this.cbx_ignore_ssl.Visible = true;
                        this.txt_check_result.Text = "翻译完成";
                        BatchModeOutput(resultList);
                    }
                    catch (Exception ex)
                    {
                        this.gbx_file.Visible = true;
                        // this.cbx_ignore_ssl.Visible = true;
                        this.txt_check_result.Text = ex.Message;
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });
        }

        public void BatchModeCheck(List<List<string>> list)
        {
            var errList = new List<string>();
            foreach (var chineseList in list)
            {
                foreach (string chinese in chineseList)
                {
                    //var newChinese = chinese
                    //    .Replace("技术产业开发区", " 技术产业开发区")
                    //    .Replace("自治州", " 自治州")
                    //    .Replace("产业开发区", " 产业开发区")
                    //    .Replace("自治旗", " 自治旗")
                    //    .Replace("经济开发区", " 经济开发区")
                    //    .Replace("技术开发区", " 技术开发区")
                    //    .Replace("改革示范区", " 改革示范区")
                    //    .Replace("综合实验区", " 综合实验区")
                    //    .Replace("一体化示范区", " 一体化示范区")
                    //    .Replace("产业聚集区", " 产业聚集区")
                    //    .Replace("海域", " 海域")
                    //    .Replace("傣族自治县", " 傣族自治县")
                    //    .Replace("撒拉族自治县", " 撒拉族自治县")
                    //    .Replace("现代产业园区", " 现代产业园区")
                    //    .Replace("技术产业园区", " 技术产业园区")
                    //    .Replace("自治县", " 自治县");
                    string pattern = @"[\u4e00-\u9fa5]+";
                    MatchCollection matches = Regex.Matches(chinese, pattern);
                    foreach (System.Text.RegularExpressions.Match match in matches)
                    {
                        string chineseWords = match.Value;
                        if (chineseWords.Length > 10)
                        {
                            errList.Add($"‘{chineseWords}’ 长度超过限制，每个词语最多10个汉字");
                        }
                    }
                }
            }

            if (errList.Count > 0)
            {
                this.txt_check_result.Text = string.Join("\r\n", errList);
                return;
            }
            else
            {
                BatchModeTranslate(list);
            }
        }

        public List<List<string>> LoadExcel(string path)
        {
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                #region 读取表格

                IWorkbook workbook = null;
                try
                {
                    workbook = new XSSFWorkbook(fileStream);
                }
                catch (ZipException)
                {
                    try
                    {
                        fileStream.Position = 0;
                        workbook = WorkbookFactory.Create(fileStream, "12000", true);
                    }
                    catch (Exception)
                    {
                        try
                        {
                            fileStream.Position = 0;
                            workbook = WorkbookFactory.Create(fileStream, "144000", true);
                        }
                        catch (Exception)
                        {
                            try
                            {
                                fileStream.Position = 0;
                                workbook = WorkbookFactory.Create(fileStream, "wh12000", true);
                            }
                            catch (Exception)
                            {
                                fileStream.Position = 0;
                                workbook = WorkbookFactory.Create(fileStream, "0314", true);
                            }
                        }
                    }
                }

                #endregion

                try
                {

                    List<List<string>> chineseColumns = new List<List<string>>();

                    ISheet sheet = workbook.GetSheetAt(0);

                    // 获取第一行，找到所有标题为“中文”的列索引
                    IRow headerRow = sheet.GetRow(0);
                    List<int> chineseColumnIndexes = new List<int>();

                    for (int i = 0; i < headerRow.LastCellNum; i++)
                    {
                        ICell cell = headerRow.GetCell(i);
                        if (cell != null && cell.StringCellValue == "中文")
                        {
                            chineseColumnIndexes.Add(i);
                        }
                    }

                    // 读取每一个中文列的数据
                    foreach (int colIndex in chineseColumnIndexes)
                    {
                        List<string> columnData = new List<string>();

                        for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                        {
                            IRow row = sheet.GetRow(rowIndex);
                            if (row != null)
                            {
                                ICell cell = row.GetCell(colIndex);
                                if (cell != null && cell.CellType == CellType.String)
                                {
                                    if (!string.IsNullOrEmpty(cell.StringCellValue))
                                    {
                                        columnData.Add(cell.StringCellValue);
                                    }
                                }
                                else
                                {
                                    // columnData.Add(string.Empty);  // 保持行对齐
                                }
                            }
                        }

                        if (columnData.Count > 0)
                        {
                            chineseColumns.Add(columnData);
                        }
                    }

                    return chineseColumns;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return new List<List<string>>();
                }
                finally
                {
                    workbook.Close();
                    fileStream.Close();
                }
            }
        }
    }
}
