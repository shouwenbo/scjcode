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
using System.Reflection.Metadata;
using OfficeOpenXml;
using System.Globalization;
using System.Linq;

namespace BatchTranslateApp
{
    public partial class OnceOpenTranslateForm : ExcelForm
    {
        private int allCount = 0;
        private int currentCount = 0;

        public OnceOpenTranslateForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
            InitExcelSelector(this.gbx_file, this.btn_select_file, file =>
            {
                Run(file);
                //var list = LoadExcel(file);
                //if (list.Count > 0)
                //{
                //    Translate(list);
                //}
            });
            CheckForIllegalCrossThreadCalls = false; // 关闭跨线程调用检查
        }

        public void Output_NPOI(List<CdbStudent_NPOI> list)
        {
            using (FileStream file = new FileStream(@"template/一开韩文模板.xlsx", FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(file);
                ISheet sheet = workbook.GetSheetAt(2);
                sheet.ForceFormulaRecalculation = true;
                var periods = new List<string>();

                for (int i = 0; i < list.Count; i++)
                {
                    var member = list[i];
                    var rowIndex = i + 6;
                    if (!periods.Contains(member.Period))
                    {
                        periods.Add(member.Period);
                    }
                    sheet.GetRow(rowIndex).GetCell(0).SetCellValueIfNotEmpty(member.Period);
                    sheet.GetRow(rowIndex).GetCell(1).SetCellValueIfNotEmpty(member.Department);
                    sheet.GetRow(rowIndex).GetCell(2).SetCellValueIfNotEmpty(member.KoreanName);
                    sheet.GetRow(rowIndex).GetCell(3).SetCellValueIfNotEmpty(member.ChineseName);
                    sheet.GetRow(rowIndex).GetCell(4).SetCellValueIfNotEmpty(member.Gender);
                    sheet.GetRow(rowIndex).GetCell(5).SetCellValueIfNotEmpty(member.IDCardBirth);
                    sheet.GetRow(rowIndex).GetCell(8).SetCellValueIfNotEmpty("중화인민공화국"); // cbd数据缺失
                    sheet.GetRow(rowIndex).GetCell(9).SetCellValueIfNotEmpty(member.Phone);
                    sheet.GetRow(rowIndex).GetCell(10).SetCellValueIfNotEmpty("중화인민공화국"); // cbd数据缺失
                    sheet.GetRow(rowIndex).GetCell(11).SetCellValueIfNotEmpty(member.KoreanCity);
                    sheet.GetRow(rowIndex).GetCell(12).SetCellValueIfNotEmpty(member.CurrentAddress);
                    sheet.GetRow(rowIndex).GetCell(13).SetCellValueIfNotEmpty(member.IDCardAddress);
                    sheet.GetRow(rowIndex).GetCell(14).SetCellValueIfNotEmpty(member.Job);
                    sheet.GetRow(rowIndex).GetCell(15).SetCellValueIfNotEmpty(string.IsNullOrEmpty(member.ReStudy) ? "1" : "2");
                    //sheet.GetRow(rowIndex).GetCell(16).SetCellValue("노방");
                    sheet.GetRow(rowIndex).GetCell(17).SetCellValueIfNotEmpty(member.Times);
                    sheet.GetRow(rowIndex).GetCell(18).SetCellValueIfNotEmpty("0");
                    sheet.GetRow(rowIndex).GetCell(20).SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    sheet.GetRow(rowIndex).GetCell(21).SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    sheet.GetRow(rowIndex).GetCell(22).SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    sheet.GetRow(rowIndex).GetCell(23).SetCellValueIfNotEmpty("부산야고보지파");
                    sheet.GetRow(rowIndex).GetCell(24).SetCellValueIfNotEmpty("중국무한교회");
                    sheet.GetRow(rowIndex).GetCell(25).SetCellValueIfNotEmpty(member.GuiderDepartment);
                    sheet.GetRow(rowIndex).GetCell(26).SetCellValueIfNotEmpty(member.GuiderKoreanName);
                    sheet.GetRow(rowIndex).GetCell(27).SetCellValueIfNotEmpty(member.GuiderNumber);
                    sheet.GetRow(rowIndex).GetCell(28).SetCellValueIfNotEmpty(member.GuiderSchoolPeriod);
                    sheet.GetRow(rowIndex).GetCell(29).SetCellValueIfNotEmpty("부산야고보지파");
                    sheet.GetRow(rowIndex).GetCell(30).SetCellValueIfNotEmpty("중국무한교회");
                    sheet.GetRow(rowIndex).GetCell(31).SetCellValueIfNotEmpty(member.TeacherDepartment);
                    sheet.GetRow(rowIndex).GetCell(32).SetCellValueIfNotEmpty(member.TeacherKoreanName);
                    sheet.GetRow(rowIndex).GetCell(33).SetCellValueIfNotEmpty(member.TeacherNumber);
                    sheet.GetRow(rowIndex).GetCell(39).SetCellValueIfNotEmpty("1");
                    sheet.GetRow(rowIndex).GetCell(41).SetCellValueIfNotEmpty("1");
                    sheet.GetRow(rowIndex).GetCell(46).SetCellType(CellType.Formula);
                    sheet.GetRow(rowIndex).GetCell(46).CellFormula = $"AP{(rowIndex + 1)}+AQ{(rowIndex + 1)}-AR{(rowIndex + 1)}+AS{(rowIndex + 1)}-AT{(rowIndex + 1)}";
                    
                }

                sheet.DeleteRowRangeAndShiftUp(list.Count + 6, 2309); // 第二个参数：行数-1

                //sheet.DeleteRowsAndShiftUp(new List<int>() { 6 }); 
                //if (list.Count > 0)
                //{
                //    sheet.GetRow(6).GetCell(0).SetCellValueIfNotEmpty(list[0].Period);
                //}

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "保存 Excel 文件";
                    saveFileDialog.FileName = $"一开（{string.Join(" ", periods)}） 韩.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string targetFilePath = saveFileDialog.FileName;
                        using (FileStream stream = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(stream);
                        }
                        this.txt_check_result.Text = $"文件已成功保存到：{targetFilePath}";
                        var fileInfo = new FileInfo("yourFile.xlsx");

                        Process.Start(new ProcessStartInfo(targetFilePath) { UseShellExecute = true });
                    }
                }
            }
        }

        public void Output(List<CdbStudent> list)
        {
            using (var package = new ExcelPackage(new FileInfo(@"template/一开韩文模板.xlsx")))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[2];
                var periods = new List<string>();

                for (int i = 0; i < list.Count; i++)
                {
                    var member = list[i];
                    var rowIndex = i + 7;
                    if (!periods.Contains(member.Period.Text))
                    {
                        periods.Add(member.Period.Text);
                    }

                    sheet.Cells[rowIndex, 1].SetCell(member.Period);
                    sheet.Cells[rowIndex, 2].SetCell(member.Department);
                    sheet.Cells[rowIndex, 3].SetCell(member.KoreanName);
                    sheet.Cells[rowIndex, 4].SetCell(member.ChineseName);
                    sheet.Cells[rowIndex, 5].SetCell(member.Gender);
                    sheet.Cells[rowIndex, 6].SetCell(member.IDCardBirth);
                    if (!string.IsNullOrEmpty(member.ReStudy.Text))
                    {
                        sheet.Cells[rowIndex, 7].SetCell("재수강");
                    }
                    sheet.Cells[rowIndex, 8].SetCell(member.ReStudy);
                    sheet.Cells[rowIndex, 9].SetCell("중화인민공화국"); // cbd数据缺失
                    sheet.Cells[rowIndex, 10].SetCell(member.Phone);
                    sheet.Cells[rowIndex, 11].SetCell("중화인민공화국"); // cbd数据缺失
                    sheet.Cells[rowIndex, 12].SetCell(member.KoreanCity);
                    sheet.Cells[rowIndex, 13].SetCell(member.CurrentAddress);
                    sheet.Cells[rowIndex, 14].SetCell(member.IDCardAddress);
                    sheet.Cells[rowIndex, 15].SetCell(member.Job);
                    sheet.Cells[rowIndex, 16].SetCell(string.IsNullOrEmpty(member.ReStudy.Text) ? "1" : "2");
                    sheet.Cells[rowIndex, 17].SetCell(member.CDType);
                    sheet.Cells[rowIndex, 18].SetCell(member.Times);
                    sheet.Cells[rowIndex, 19].SetCell("0");
                    sheet.Cells[rowIndex, 21].SetCell(""); // cbd数据缺失，暂时默认为
                    sheet.Cells[rowIndex, 22].SetCell(""); // cbd数据缺失，暂时默认为
                    sheet.Cells[rowIndex, 23].SetCell(""); // cbd数据缺失，暂时默认为
                    sheet.Cells[rowIndex, 24].SetCell("부산야고보지파");
                    sheet.Cells[rowIndex, 25].SetCell("중국무한교회");
                    sheet.Cells[rowIndex, 26].SetCell(member.GuiderDepartment);
                    sheet.Cells[rowIndex, 27].SetCell(member.GuiderKoreanName);
                    sheet.Cells[rowIndex, 28].SetCell(member.GuiderNumber);
                    sheet.Cells[rowIndex, 29].SetCell(member.GuiderSchoolPeriod);
                    sheet.Cells[rowIndex, 30].SetCell("부산야고보지파");
                    sheet.Cells[rowIndex, 31].SetCell("중국무한교회");
                    sheet.Cells[rowIndex, 32].SetCell(member.TeacherDepartment);
                    sheet.Cells[rowIndex, 33].SetCell(member.TeacherKoreanName);
                    sheet.Cells[rowIndex, 34].SetCell(member.TeacherNumber);
                    sheet.Cells[rowIndex, 40].SetCell("1");
                    sheet.Cells[rowIndex, 42].SetCell("1");

                    // sheet.Cells[rowIndex, 47].Formula = $"AP{(rowIndex + 1)}+AQ{(rowIndex + 1)}-AR{(rowIndex + 1)}+AS{(rowIndex + 1)}-AT{(rowIndex + 1)}";

                    // sheet.Cells[rowIndex, 1].SetCellValueIfNotEmpty(member.Period);
                    // sheet.Cells[rowIndex, 2].SetCellValueIfNotEmpty(member.Department);
                    // sheet.Cells[rowIndex, 3].SetCellValueIfNotEmpty(member.KoreanName);
                    // sheet.Cells[rowIndex, 4].SetCellValueIfNotEmpty(member.ChineseName);
                    // sheet.Cells[rowIndex, 5].SetCellValueIfNotEmpty(member.Gender);
                    // sheet.Cells[rowIndex, 6].SetCellValueIfNotEmpty(member.IDCardBirth);
                    // sheet.Cells[rowIndex, 9].SetCellValueIfNotEmpty("중화인민공화국"); // cbd数据缺失
                    // sheet.Cells[rowIndex, 10].SetCellValueIfNotEmpty(member.Phone);
                    // sheet.Cells[rowIndex, 11].SetCellValueIfNotEmpty("중화인민공화국"); // cbd数据缺失
                    // sheet.Cells[rowIndex, 12].SetCellValueIfNotEmpty(member.KoreanCity);
                    // sheet.Cells[rowIndex, 13].SetCellValueIfNotEmpty(member.CurrentAddress);
                    // sheet.Cells[rowIndex, 14].SetCellValueIfNotEmpty(member.IDCardAddress);
                    // sheet.Cells[rowIndex, 15].SetCellValueIfNotEmpty(member.Job);
                    // sheet.Cells[rowIndex, 16].SetCellValueIfNotEmpty(string.IsNullOrEmpty(member.ReStudy) ? "1" : "2");
                    // sheet.Cells[rowIndex, 17].SetCellValueIfNotEmpty(member.CDType);
                    // sheet.Cells[rowIndex, 18].SetCellValueIfNotEmpty(member.Times);
                    // sheet.Cells[rowIndex, 19].SetCellValueIfNotEmpty("0");
                    // sheet.Cells[rowIndex, 21].SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    // sheet.Cells[rowIndex, 22].SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    // sheet.Cells[rowIndex, 23].SetCellValueIfNotEmpty(""); // cbd数据缺失，暂时默认为
                    // sheet.Cells[rowIndex, 24].SetCellValueIfNotEmpty("부산야고보지파");
                    // sheet.Cells[rowIndex, 25].SetCellValueIfNotEmpty("중국무한교회");
                    // sheet.Cells[rowIndex, 26].SetCellValueIfNotEmpty(member.GuiderDepartment);
                    // sheet.Cells[rowIndex, 27].SetCellValueIfNotEmpty(member.GuiderKoreanName);
                    // sheet.Cells[rowIndex, 28].SetCellValueIfNotEmpty(member.GuiderNumber);
                    // sheet.Cells[rowIndex, 29].SetCellValueIfNotEmpty(member.GuiderSchoolPeriod);
                    // sheet.Cells[rowIndex, 30].SetCellValueIfNotEmpty("부산야고보지파");
                    // sheet.Cells[rowIndex, 31].SetCellValueIfNotEmpty("중국무한교회");
                    // sheet.Cells[rowIndex, 32].SetCellValueIfNotEmpty(member.TeacherDepartment);
                    // sheet.Cells[rowIndex, 33].SetCellValueIfNotEmpty(member.TeacherKoreanName);
                    // sheet.Cells[rowIndex, 34].SetCellValueIfNotEmpty(member.TeacherNumber);
                    // sheet.Cells[rowIndex, 40].SetCellValueIfNotEmpty("1");
                    // sheet.Cells[rowIndex, 42].SetCellValueIfNotEmpty("1");
                    // 
                    // sheet.Cells[rowIndex, 47].Formula = $"AP{(rowIndex + 1)}+AQ{(rowIndex + 1)}-AR{(rowIndex + 1)}+AS{(rowIndex + 1)}-AT{(rowIndex + 1)}";

                }
                sheet.DeleteRow(list.Count + 7, 2304 - list.Count); // 第二个参数：行数-1 2308

                //sheet.DeleteRowsAndShiftUp(new List<int>() { 6 }); 
                //if (list.Count > 0)
                //{
                //    sheet.GetRow(6).GetCell(0).SetCellValueIfNotEmpty(list[0].Period);
                //}

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "保存 Excel 文件";
                    saveFileDialog.FileName = $"一开（{string.Join(" ", periods)}） 韩.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string targetFilePath = saveFileDialog.FileName;
                        package.SaveAs(new FileInfo(targetFilePath));
                        this.txt_check_result.Text = $"文件已成功保存到：{targetFilePath}";
                        Process.Start(new ProcessStartInfo(targetFilePath) { UseShellExecute = true });
                    }
                }
            }
        }

        public void Translate_NPOI(List<CdbStudent_NPOI> list)
        {
            this.gbx_file.Visible = false;
            currentCount = 0;
            Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        allCount = list.Count;
                        var taskBuilder = new StringBuilder();
                        foreach (var member in list)
                        {
                            if (!string.IsNullOrEmpty(member.Job))
                            {
                                member.Job = TranslateHelper.SentenceToKorean(member.Job);
                            }
                            if (!string.IsNullOrEmpty(member.CDType))
                            {
                                member.CDType = member.CDType
                                      .Replace("网络传道", "인터넷/통신")
                                      .Replace("网络", "인터넷/통신")
                                      .Replace("通信", "인터넷/통신")
                                      .Replace("家族", "가족/친척")
                                      .Replace("亲戚", "가족/친척")
                                      .Replace("路旁", "노방")
                                      .Replace("知人", "친구/선후배/직장동료")
                                      .Replace("朋友", "친구/선후배/직장동료")
                                      .Replace("前后辈", "친구/선후배/직장동료")
                                      .Replace("同事", "친구/선후배/직장동료")
                                      .Replace("路旁", "노방")
                                      .Replace("田地", "추수밭");
                            }
                            currentCount++;
                            this.lbl_task.Text = $"{currentCount}/{allCount}";
                        }
                        this.gbx_file.Visible = true;
                        this.txt_check_result.Text = "翻译完成";
                        Output_NPOI(list);
                    }
                    catch (Exception ex)
                    {
                        this.gbx_file.Visible = true;
                        this.txt_check_result.Text = ex.Message;
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });
        }

        public void Run(string path)
        {
            this.gbx_file.Visible = false;
            currentCount = 0;
            Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                        FileInfo fileInfo = new FileInfo(path);
                        if (!fileInfo.Exists)
                        {
                            this.txt_check_result.Text = "文件不存在";
                        }

                        var list = new List<CdbStudent>();
                        using (var package = new ExcelPackage(fileInfo, this.txt_cdb_password.Text))
                        {
                            #region 读取cdb周报

                            var sheet = package.Workbook.Worksheets[0];

                            // 遍历工作表的每一行（从第2行开始，跳过标题行）
                            for (int rowIndex = 2; rowIndex <= sheet.Dimension.End.Row; rowIndex++)  // EPPlus的行从1开始，索引为1表示第二行
                            {
                                var student = new CdbStudent()
                                {
                                    No = sheet.Cells[rowIndex, 1],  // 读取第1列（A列）
                                    Period = sheet.Cells[rowIndex, 2],  // 读取第2列（B列）
                                    ChineseName = sheet.Cells[rowIndex, 3],  // 读取第3列（C列）
                                    KoreanName = sheet.Cells[rowIndex, 4],  // 读取第4列（D列）
                                    Department = sheet.Cells[rowIndex, 5],  // 读取第5列（E列）
                                    Gender = sheet.Cells[rowIndex, 6],  // 读取第6列（F列）
                                    ReStudy = sheet.Cells[rowIndex, 7],  // 读取第7列（G列）
                                    IDCardBirth = sheet.Cells[rowIndex, 8],  // 读取第8列（H列）
                                    CDType = sheet.Cells[rowIndex, 9],  // 读取第9列（I列）
                                    Phone = sheet.Cells[rowIndex, 10],  // 读取第10列（J列）
                                    KoreanCity = sheet.Cells[rowIndex, 11],  // 读取第11列（K列）
                                    CurrentAddress = sheet.Cells[rowIndex, 12],  // 读取第12列（L列）
                                    IDCardAddress = sheet.Cells[rowIndex, 13],  // 读取第13列（M列）
                                    Job = sheet.Cells[rowIndex, 14],  // 读取第14列（N列）
                                    GuiderDepartment = sheet.Cells[rowIndex, 15],  // 读取第15列（O列）
                                    GuiderChineseName = sheet.Cells[rowIndex, 16],  // 读取第16列（P列）
                                    GuiderKoreanName = sheet.Cells[rowIndex, 17],  // 读取第17列（Q列）
                                    GuiderNumber = sheet.Cells[rowIndex, 18],  // 读取第18列（R列）
                                    GuiderSchoolPeriod = sheet.Cells[rowIndex, 19],  // 读取第19列（S列）
                                    TeacherDepartment = sheet.Cells[rowIndex, 20],  // 读取第20列（T列）
                                    TeacherChineseName = sheet.Cells[rowIndex, 21],  // 读取第21列（U列）
                                    TeacherKoreanName = sheet.Cells[rowIndex, 22],  // 读取第22列（V列）
                                    TeacherNumber = sheet.Cells[rowIndex, 23],  // 读取第23列（W列）
                                    Times = sheet.Cells[rowIndex, 24],  // 读取第24列（X列）
                                };

                                // 如果存在职位信息，转换为韩文
                                if (!string.IsNullOrEmpty(student.Job.Text))
                                {
                                    student.Job.Value = TranslateHelper.SentenceToKorean(student.Job.Text);
                                }

                                // 如果Period不为空则加入到列表
                                if (!string.IsNullOrEmpty(student.Period.Text.Trim()))
                                {
                                    list.Add(student);
                                }
                            }

                            #endregion

                            if (list.Count > 0)
                            {
                                #region 翻译

                                allCount = list.Count;
                                var taskBuilder = new StringBuilder();
                                foreach (var member in list)
                                {
                                    if (!string.IsNullOrEmpty(member.Job.Text))
                                    {
                                        member.Job.Value = TranslateHelper.SentenceToKorean(member.Job.Text);
                                    }
                                    if (!string.IsNullOrEmpty(member.Gender.Text))
                                    {
                                        member.Gender.Value = member.Gender.Text
                                              .Replace("男", "남")
                                              .Replace("女", "여")
                                              .Trim();
                                    }
                                    if (!string.IsNullOrEmpty(member.CDType.Text))
                                    {
                                        member.CDType.Value = member.CDType.Text
                                              .Replace("网络传道", "인터넷/통신")
                                              .Replace("网络", "인터넷/통신")
                                              .Replace("通信", "인터넷/통신")
                                              .Replace("家族", "가족/친척")
                                              .Replace("亲戚", "가족/친척")
                                              .Replace("路旁", "노방")
                                              .Replace("知人", "친구/선후배/직장동료")
                                              .Replace("朋友", "친구/선후배/직장동료")
                                              .Replace("前后辈", "친구/선후배/직장동료")
                                              .Replace("同事", "친구/선후배/직장동료")
                                              .Replace("路旁", "노방")
                                              .Replace("田地", "추수밭")
                                              .Trim();
                                    }
                                    currentCount++;
                                    this.lbl_task.Text = $"{currentCount}/{allCount}";
                                }
                                this.gbx_file.Visible = true;
                                this.txt_check_result.Text = "翻译完成";

                                #endregion

                                #region 排序

                                var departmentSortDic = new Dictionary<string, int>
                                {
                                    { "자문", 1 },
                                    { "장년", 2 },
                                    { "부녀", 3 },
                                    { "청년", 4 }
                                };

                                const int defaultSortValue = int.MaxValue;

                                var sortedList = list
                                    .OrderBy(student => student.Period.Text)
                                    .ThenBy(student =>
                                    {
                                        // 使用字典中的值进行排序，如果不在字典中则使用默认排序值
                                        return departmentSortDic.ContainsKey(student.Department.Text)
                                            ? departmentSortDic[student.Department.Text]
                                            : defaultSortValue;
                                    })
                                    .ThenBy(student => student.IDCardBirth.Text) // 按年龄从大到小排序
                                    .ToList();

                                #endregion

                                Output(sortedList);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.gbx_file.Visible = true;
                        this.txt_check_result.Text = ex.Message;
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });
        }

        public List<CdbStudent_NPOI> LoadExcel_NPOI(string path)
        {
            #region 读取表格

            IWorkbook workbook = path.GetNpoiWorkbook(out string openErr, "12000", "144000", "wh12000", "0314", this.txt_cdb_password.Text);
            if (workbook == null)
            {
                this.txt_check_result.Text = openErr;
                return new List<CdbStudent_NPOI>();
            }

            #endregion

            try
            {
                var list = new List<CdbStudent_NPOI>();
                ISheet sheet = workbook.GetSheetAt(0);
                for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    var student = new CdbStudent_NPOI()
                    {
                        No = row.GetCell(0).GetCellStringValue(),
                        Period = row.GetCell(1).GetCellStringValue(),
                        ChineseName = row.GetCell(2).GetCellStringValue(),
                        KoreanName = row.GetCell(3).GetCellStringValue(),
                        Department = row.GetCell(4).GetCellStringValue(),
                        Gender = row.GetCell(5).GetCellStringValue(),
                        ReStudy = row.GetCell(6).GetCellStringValue(),
                        IDCardBirth = row.GetCell(7).GetCellStringValue(),
                        CDType = row.GetCell(8).GetCellStringValue(),
                        Phone = row.GetCell(9).GetCellStringValue(),
                        KoreanCity = row.GetCell(10).GetCellStringValue(),
                        CurrentAddress = row.GetCell(11).GetCellStringValue(),
                        IDCardAddress = row.GetCell(12).GetCellStringValue(),
                        Job = row.GetCell(13).GetCellStringValue(),
                        GuiderDepartment = row.GetCell(14).GetCellStringValue(),
                        GuiderChineseName = row.GetCell(15).GetCellStringValue(),
                        GuiderKoreanName = row.GetCell(16).GetCellStringValue(),
                        GuiderNumber = row.GetCell(17).GetCellStringValue(),
                        GuiderSchoolPeriod = row.GetCell(18).GetCellStringValue(),
                        TeacherDepartment = row.GetCell(19).GetCellStringValue(),
                        TeacherChineseName = row.GetCell(20).GetCellStringValue(),
                        TeacherKoreanName = row.GetCell(21).GetCellStringValue(),
                        TeacherNumber = row.GetCell(22).GetCellStringValue(),
                        Times = row.GetCell(23).GetCellStringValue(),
                    };
                    if (!string.IsNullOrEmpty(student.Job))
                    {
                        student.Job = TranslateHelper.SentenceToKorean(student.Job);
                    }
                    if (!string.IsNullOrEmpty(student.Period.Trim()))
                    {
                        list.Add(student);
                    }
                }

                return list;
            }
            catch (Exception ex)
            {
                this.txt_check_result.Text = ex.Message;
                return new List<CdbStudent_NPOI>();
            }
            finally
            {
                workbook.Close();
            }
        }

        public DateTime? ParseBirthDate(string birthDate)
        {
            // 尝试将字符串转换为日期格式，若格式不对则返回 null
            if (DateTime.TryParseExact(birthDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
            {
                return date;
            }

            // 如果格式不正确，返回 null 或某个默认值（例如最早的日期）
            return null;
        }
    }

    /// <summary>
    /// Cdb学生
    /// </summary>
    public class CdbStudent_NPOI
    {
        public string No { get; set; }
        public string Period { get; set; }
        public string ChineseName { get; set; }
        public string KoreanName { get; set; }
        public string Department { get; set; }
        public string Gender { get; set; }
        public string ReStudy { get; set; }
        public string IDCardBirth { get; set; }
        public string CDType { get; set; }
        public string Phone { get; set; }
        public string KoreanCity { get; set; }
        public string CurrentAddress { get; set; }
        public string IDCardAddress { get; set; }
        public string Job { get; set; }
        public string GuiderDepartment { get; set; }
        public string GuiderChineseName { get; set; }
        public string GuiderKoreanName { get; set; }
        public string GuiderNumber { get; set; }
        public string GuiderSchoolPeriod { get; set; }
        public string TeacherDepartment { get; set; }
        public string TeacherChineseName { get; set; }
        public string TeacherKoreanName { get; set; }
        public string TeacherNumber { get; set; }
        public string Times { get; set; }
    }

    public class CdbStudent
    {
        public ExcelRange No { get; set; }
        public ExcelRange Period { get; set; }
        public ExcelRange ChineseName { get; set; }
        public ExcelRange KoreanName { get; set; }
        public ExcelRange Department { get; set; }
        public ExcelRange Gender { get; set; }
        public ExcelRange ReStudy { get; set; }
        public ExcelRange IDCardBirth { get; set; }
        public ExcelRange CDType { get; set; }
        public ExcelRange Phone { get; set; }
        public ExcelRange KoreanCity { get; set; }
        public ExcelRange CurrentAddress { get; set; }
        public ExcelRange IDCardAddress { get; set; }
        public ExcelRange Job { get; set; }
        public ExcelRange GuiderDepartment { get; set; }
        public ExcelRange GuiderChineseName { get; set; }
        public ExcelRange GuiderKoreanName { get; set; }
        public ExcelRange GuiderNumber { get; set; }
        public ExcelRange GuiderSchoolPeriod { get; set; }
        public ExcelRange TeacherDepartment { get; set; }
        public ExcelRange TeacherChineseName { get; set; }
        public ExcelRange TeacherKoreanName { get; set; }
        public ExcelRange TeacherNumber { get; set; }
        public ExcelRange Times { get; set; }
    }
}
