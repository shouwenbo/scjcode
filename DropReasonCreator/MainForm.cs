using Common;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.VisualBasic;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;

namespace DropReasonCreator
{
    public partial class MainForm : Form
    {
        public List<string> MainReasonList = new List<string>
        {
            "깨달음부족",
            "신앙심부족",
            "실상믿음부족",
            "세상욕심",
            "침(언론)",
            "침(사람)",
            "수강환경안됨",
            "타인에의한수강",
            "건강악화",
            "건강약화",
            "행함부담",
            "정신질환"
        };

        public List<string> CategoryList = new List<string>
        {
            "중등",
            "고등"
        };

        public MainForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
        }

        public List<List<Student>> LoadExcel(string path, List<string> periods)
        {
            #region 读取表格

            IWorkbook workbook = path.GetNpoiWorkbook(out string openErr, "12000", "144000", "wh12000", "0314");
            if (workbook == null)
            {
                this.txt_check_result.Text = openErr;
                return new List<List<Student>>();
            }

            #endregion

            try
            {
                var result = new List<List<Student>>();

                var 初级List = new List<Student>();
                ISheet 初级sheet = workbook.GetSheetAt(0);
                for (int 初级rowIndex = 1; 初级rowIndex <= 初级sheet.LastRowNum; 初级rowIndex++)
                {
                    IRow 初级row = 初级sheet.GetRow(初级rowIndex);
                    var student = new Student()
                    {
                        KoreanName = 初级row.GetCell(1).GetCellStringValue(),
                        ChineseName = 初级row.GetCell(2).GetCellStringValue(),
                        School = 初级row.GetCell(4).GetCellStringValue(),
                        Period = 初级row.GetCell(5).GetCellStringValue(),
                        Category = 初级row.GetCell(9).GetCellStringValue(),
                        Step = 初级row.GetCell(10).GetCellStringValue().Trim('0'),
                        MainReason = MainReasonList.Any(p => p == 初级row.GetCell(7).GetCellStringValue()) ? 初级row.GetCell(7).GetCellStringValue() : "기타",
                        OtherReason = !MainReasonList.Any(p => p == 初级row.GetCell(7).GetCellStringValue()) ? $"기타({初级row.GetCell(7).GetCellStringValue()})" : "",
                    };
                    if (student.MainReason == "건강약화")
                    {
                        student.MainReason = "건강악화";
                    }
                    if (periods.Count > 0 && !periods.Any(p => p == student.Period))
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(student.Period.Trim()))
                    {
                        初级List.Add(student);
                    }
                }

                var 中高级List = new List<Student>();
                ISheet 中高级sheet = workbook.GetSheetAt(1);
                for (int 中高级rowIndex = 1; 中高级rowIndex <= 中高级sheet.LastRowNum; 中高级rowIndex++)
                {
                    IRow 中高级row = 中高级sheet.GetRow(中高级rowIndex);
                    var student = new Student()
                    {
                        KoreanName = 中高级row.GetCell(1).GetCellStringValue(),
                        ChineseName = 中高级row.GetCell(2).GetCellStringValue(),
                        School = 中高级row.GetCell(4).GetCellStringValue(),
                        Period = 中高级row.GetCell(5).GetCellStringValue(),
                        Category = 中高级row.GetCell(9).GetCellStringValue(),
                        Step = 中高级row.GetCell(10).GetCellStringValue().Trim('0'),
                        MainReason = MainReasonList.Any(p => p == 中高级row.GetCell(7).GetCellStringValue()) ? 中高级row.GetCell(7).GetCellStringValue() : "기타",
                        OtherReason = !MainReasonList.Any(p => p == 中高级row.GetCell(7).GetCellStringValue()) ? $"기타({中高级row.GetCell(7).GetCellStringValue()})" : "",
                    };
                    if (student.MainReason == "건강약화")
                    {
                        student.MainReason = "건강악화";
                    }
                    if (student.Category == "새신자")
                    {
                        student.Category = "새신자교육";
                    }
                    if (periods.Count > 0 && !periods.Any(p => p == student.Period))
                    {
                        continue;
                    }
                    if (!string.IsNullOrEmpty(student.Period.Trim()))
                    {
                        中高级List.Add(student);
                    }
                }

                if (初级List.Count == 0 && 中高级List.Count == 0)
                {
                    this.txt_check_result.Text = $"掉落表 找不到{string.Join("/", periods)}期次的任何数据";
                    return new List<List<Student>>();
                }

                result.Add(初级List);
                result.Add(中高级List);

                return result;
            }
            catch (Exception ex)
            {
                this.txt_check_result.Text = ex.Message;
                return new List<List<Student>>();
            }
            finally
            {
                workbook.Close();
            }
        }

        public void Output(List<List<Student>> dropData, string dateStr, List<string> periods)
        {
            if (dropData.Count != 2)
            {
                return;
            }
            IWorkbook workbook = new XSSFWorkbook();
            ISheet 初级Sheet = workbook.CreateSheet("初级掉落事由");
            ISheet 中高级Sheet = workbook.CreateSheet("中高级掉落事由");

            var columnList = new List<string>() { "学院", "期数", "韩文名", "中文名", "分类", "阶段", "事由", "后续", "不填", "其他" };
            IRow 初级首row = 初级Sheet.CreateRow(0);
            IRow 中高级首row = 中高级Sheet.CreateRow(0);
            for (int i = 0; i < columnList.Count; i++)
            {
                ICell 初级cell = 初级首row.CreateCell(i);
                初级cell.SetCellValue(columnList[i]);
                ICell 中高级cell = 中高级首row.CreateCell(i);
                中高级cell.SetCellValue(columnList[i]);
            }

            for (int i = 0; i < dropData[0].Count; i++)
            {
                IRow 初级newRow = 初级Sheet.CreateRow(i + 1);
                初级newRow.CreateCell(0).SetCellValue(dropData[0][i].School);
                初级newRow.CreateCell(1).SetCellValue(dropData[0][i].Period);
                初级newRow.CreateCell(2).SetCellValue(dropData[0][i].KoreanName);
                初级newRow.CreateCell(3).SetCellValue(dropData[0][i].ChineseName);
                初级newRow.CreateCell(4).SetCellValue(dropData[0][i].Category);
                初级newRow.CreateCell(5).SetCellValue(dropData[0][i].Step);
                初级newRow.CreateCell(6).SetCellValue(dropData[0][i].MainReason);
                初级newRow.CreateCell(7).SetCellValue("탈락 후 비관리");
                初级newRow.CreateCell(8).SetCellValue("");
                初级newRow.CreateCell(9).SetCellValue(dropData[0][i].OtherReason);
            }
            // 调整初级Sheet列宽
            for (int i = 0; i < columnList.Count; i++)
            {
                初级Sheet.AutoSizeColumn(i);
                初级Sheet.SetColumnWidth(i, 初级Sheet.GetColumnWidth(i) + 2048); // 适当增加宽度偏移量
            }

            for (int i = 0; i < dropData[1].Count; i++)
            {
                IRow 中高级newRow = 中高级Sheet.CreateRow(i + 1);
                中高级newRow.CreateCell(0).SetCellValue(dropData[1][i].School);
                中高级newRow.CreateCell(1).SetCellValue(dropData[1][i].Period);
                中高级newRow.CreateCell(2).SetCellValue(dropData[1][i].KoreanName);
                中高级newRow.CreateCell(3).SetCellValue(dropData[1][i].ChineseName);
                中高级newRow.CreateCell(4).SetCellValue(dropData[1][i].Category);
                中高级newRow.CreateCell(5).SetCellValue(dropData[1][i].Step);
                中高级newRow.CreateCell(6).SetCellValue(dropData[1][i].MainReason);
                中高级newRow.CreateCell(7).SetCellValue("탈락 후 비관리");
                中高级newRow.CreateCell(8).SetCellValue("");
                中高级newRow.CreateCell(9).SetCellValue(dropData[1][i].OtherReason);
            }
            // 调整中高级Sheet列宽
            for (int i = 0; i < columnList.Count; i++)
            {
                中高级Sheet.AutoSizeColumn(i);
                中高级Sheet.SetColumnWidth(i, 中高级Sheet.GetColumnWidth(i) + 2048); // 适当增加宽度偏移量
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                saveFileDialog.Title = "保存 Excel 文件";
                saveFileDialog.FileName = $"韩掉{(!string.IsNullOrEmpty(dateStr) ? $"({dateStr})" : "")} {string.Join(" ", periods)}.xlsx";

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

        private void gbx_file_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void gbx_file_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    var file = files[0];
                    InputPeriod(file);
                }
            }
        }

        private void btn_select_dir_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                InputPeriod(openFileDialog.FileName);
            }
        }

        private void InputPeriod(string file)
        {
            var dateMatch = Regex.Match(new FileInfo(file).Name, @"\b(\d{1,2})\.(\d{1,2})\b");
            var dateStr = dateMatch.Success ? dateMatch.Value : "";
            string userInput = Interaction.InputBox("请输入期数（多个期数请用逗号分隔，如不填写则默认为全部期数）：", "输入期数", "");
            var periods = userInput
                   .Split(new[] { ",", "，", " " }, StringSplitOptions.RemoveEmptyEntries)
                   .Where(s => !string.IsNullOrWhiteSpace(s))
                   .ToList();
            //if (periods.Count == 0)
            //{
            //    this.txt_check_result.Text = "未输入期数";
            //    return;
            //}
            this.txt_check_result.Text = $"正在整理{string.Join("/", periods)}";
            var list = LoadExcel(file, periods);
            Output(list, dateStr, periods);
        }
    }

    public class Student
    {
        /// <summary>
        /// 韩文名
        /// </summary>
        public string KoreanName { get; set; }
        /// <summary>
        /// 中文名
        /// </summary>
        public string ChineseName { get; set; }
        /// <summary>
        /// 学院
        /// </summary>
        public string School { get; set; }
        /// <summary>
        /// 期数
        /// </summary>
        public string Period { get; set; }
        /// <summary>
        /// 主要事由
        /// </summary>
        public string MainReason { get; set; }
        /// <summary>
        /// 其他事由
        /// </summary>
        public string OtherReason { get; set; }
        /// <summary>
        /// 分类
        /// </summary>
        public string Category { get; set; }
        /// <summary>
        /// 阶段
        /// </summary>
        public string Step { get; set; }
    }
}
