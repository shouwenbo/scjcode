using Common;
using OfficeOpenXml;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Xceed.Words.NET;

namespace WeeklyApp
{
    public partial class MainForm : ExcelForm
    {
        private int allCount = 0;
        private int currentCount = 0;

        private string info_file = string.Empty;
        private string drop_file = string.Empty;

        public MainForm()
        {
            InitializeComponent();
            InitExcelSelector(this.gbx_drop_file, this.btn_select_drop_file, file => { this.drop_file = file; this.lbl_drop_file_name.Text = $"已选择：{file}"; });
            InitExcelSelector(this.gbx_info_file, this.btn_select_info_file, file => { this.info_file = file; this.lbl_info_file_name.Text = $"已选择：{file}"; });

            // 测试
            // this.drop_file = @"W:\work\kzj\周报测试\template\测试 - 掉落.xlsx";
            // this.info_file = @"W:\work\kzj\周报测试\template\测试 - 信息.xlsx";
            // this.lbl_drop_file_name.Text = $"已选择：{drop_file}";
            // this.lbl_info_file_name.Text = $"已选择：{info_file}";
            // this.txt_drop_password.Text = "";
            // this.txt_info_password.Text = "";

            CheckForIllegalCrossThreadCalls = false; // 关闭跨线程调用检查
        }

        public void Output(List<DropDocument> list)
        {
            string inputFilePath = @"template\韩文掉落模板.docx";
            var sunday = DateTime.Now.GetSundayOfCurrentWeek();
            var year = sunday.Year;
            var month = sunday.Month;
            var day = sunday.Day;
            var scjyear = year - 1983;

            using (FolderBrowserDialog folderDialog = new FolderBrowserDialog())
            {
                folderDialog.Description = "请选择一个文件夹来保存 Excel 文件";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    string folderPath = folderDialog.SelectedPath;

                    foreach (var doc in list)
                    {
                        string outputFileName = $@"{folderPath}\부산야고보-탈락신청서중국무한교육관비){doc.Period}({doc.KoreanName})-신{scjyear}({year}).{month}.{day}..docx";

                        using (DocX document = DocX.Load(inputFilePath))
                        {
                            // document.ReplaceText("{{期数}}", 期数 + "期");
                            // document.ReplaceText("{{姓名}}", 姓名);
                            // document.ReplaceText("{{掉落课数}}", 掉落阶段 + 掉落课数);
                            // document.ReplaceText("{{新天纪年}}", scjyear.ToString());
                            // document.ReplaceText("{{新天纪月}}", month.ToString().PadLeft(2, '0'));
                            // document.ReplaceText("{{新天纪日}}", day.ToString().PadLeft(2, '0'));

                            document.ReplaceText("{{期数}}", $"{doc.Period}期");
                            document.ReplaceText("{{姓名}}", doc.ChineseName);
                            document.ReplaceText("{{电话}}", doc.Phone);
                            document.ReplaceText("{{掉落课数}}", doc.DropCourse);
                            document.ReplaceText("{{掉落日期}}", doc.DropDate);
                            document.ReplaceText("{{商谈者}}", doc.Visitor);
                            document.ReplaceText("{{3次}}", doc.VisitorCount);
                            document.ReplaceText("{{掉落事由}}", doc.DropReason);
                            document.ReplaceText("{{传道师}}", doc.Preacher);
                            // document.ReplaceText("{{传道师事由}}", 传道师意见);
                            document.ReplaceText("{{讲师}}", doc.Lecturer);
                            // document.ReplaceText("{{讲师事由}}", 讲师意见);
                            document.ReplaceText("{{部长}}", doc.Minister);
                            document.ReplaceText("{{新天纪年}}", scjyear.ToString());
                            document.ReplaceText("{{新天纪月}}", month.ToString());
                            document.ReplaceText("{{新天纪日}}", day.ToString());
                            document.ReplaceText("{{部长韩文名}}", doc.MinisterKoreanName);

                            // Save the modified document as a new file
                            document.SaveAs(outputFileName);
                        }
                    }

                    this.txt_check_result.Text = $"文件已成功保存到：{folderPath}";

                    // 打开文件夹而不是具体文件
                    Process.Start(new ProcessStartInfo("explorer", folderPath) { UseShellExecute = true });
                }
            }
        }

        public void Run()
        {
            #region 表格上传验证

            if (string.IsNullOrEmpty(info_file))
            {
                this.txt_check_result.Text = "请选择信息表格";
                return;
            }
            if (string.IsNullOrEmpty(drop_file))
            {
                this.txt_check_result.Text = "请选择韩掉表格";
                return;
            }

            #endregion

            this.gbx_info_file.Visible = false;
            this.gbx_drop_file.Visible = false;
            this.btn_run.Visible = false;
            currentCount = 0;
            Task.Run(() =>
            {
                var thread = new Thread(() =>
                {
                    try
                    {
                        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                        #region 表格文件验证

                        FileInfo dropFileInfo = new FileInfo(drop_file);
                        if (!dropFileInfo.Exists)
                        {
                            throw new Exception("韩掉表格文件不存在");
                        }

                        FileInfo infoFileInfo = new FileInfo(info_file);
                        if (!infoFileInfo.Exists)
                        {
                            throw new Exception("信息表格文件不存在");
                        }

                        #endregion

                        var dropInfoList = new List<DropInfo>();
                        var studentList = new List<Student>();
                        var workerList = new List<Worker>();

                        using (var dropPackage = new ExcelPackage(dropFileInfo, !string.IsNullOrEmpty(this.txt_drop_password.Text) ? this.txt_drop_password.Text : null))
                        {
                            using (var infoPackage = new ExcelPackage(infoFileInfo, !string.IsNullOrEmpty(this.txt_info_password.Text) ? this.txt_info_password.Text : null))
                            {
                                #region 读取DL表

                                var dropSheet = dropPackage.Workbook.Worksheets[0];
                                for (int rowIndex = 1; rowIndex <= dropSheet.Dimension.End.Row; rowIndex++)
                                {
                                    var dropInfo = new DropInfo()
                                    {
                                        ChineseName = dropSheet.Cells[rowIndex, 1],
                                        KoreanName = dropSheet.Cells[rowIndex, 2],
                                        IDCardBirth = dropSheet.Cells[rowIndex, 3],
                                        School = dropSheet.Cells[rowIndex, 4],
                                        Period = dropSheet.Cells[rowIndex, 5],
                                        DropDate = dropSheet.Cells[rowIndex, 6],
                                        MainReason = dropSheet.Cells[rowIndex, 7],
                                        DetailReason = dropSheet.Cells[rowIndex, 8],
                                        CourseStage = dropSheet.Cells[rowIndex, 9],
                                        CourseTimes = dropSheet.Cells[rowIndex, 10],
                                        BasicLecturer = dropSheet.Cells[rowIndex, 11],
                                        MidLecturer = dropSheet.Cells[rowIndex, 12],
                                        AdvancedLecturer = dropSheet.Cells[rowIndex, 13],
                                        Preacher = dropSheet.Cells[rowIndex, 14],
                                        Visitor = dropSheet.Cells[rowIndex, 15],
                                        VisitorCount = dropSheet.Cells[rowIndex, 16],
                                    };

                                    if (Regex.IsMatch(dropInfo.Period.Text.Trim(), @"^\d{3}(-\d{1})?$"))
                                    {
                                        dropInfoList.Add(dropInfo);
                                    }
                                }

                                #endregion

                                #region 读取信息表

                                var studentSheet = infoPackage.Workbook.Worksheets[0];
                                for (int rowIndex = 1; rowIndex <= studentSheet.Dimension.End.Row; rowIndex++)
                                {
                                    var student = new Student()
                                    {
                                        ChineseName = studentSheet.Cells[rowIndex, 1],
                                        KoreanName = studentSheet.Cells[rowIndex, 2],
                                        IDCardBirth = studentSheet.Cells[rowIndex, 3],
                                        Period = studentSheet.Cells[rowIndex, 4],
                                        BasicLecturer = studentSheet.Cells[rowIndex, 5],
                                        MidLecturer = studentSheet.Cells[rowIndex, 6],
                                        AdvancedLecturer = studentSheet.Cells[rowIndex, 7],
                                        Preacher = studentSheet.Cells[rowIndex, 8],
                                        Phone = studentSheet.Cells[rowIndex, 9],
                                    };

                                    if (Regex.IsMatch(student.Period.Text.Trim(), @"^\d{3}(-\d{1})?$"))
                                    {
                                        studentList.Add(student);
                                    }
                                }

                                var workerSheet = infoPackage.Workbook.Worksheets[1];
                                for (int rowIndex = 1; rowIndex <= workerSheet.Dimension.End.Row; rowIndex++)
                                {
                                    var worker = new Worker()
                                    {
                                        KoreanName = workerSheet.Cells[rowIndex, 1],
                                        NameAndPhone = workerSheet.Cells[rowIndex, 2],
                                    };

                                    if (!string.IsNullOrEmpty(worker.KoreanName.Text))
                                    {
                                        workerList.Add(worker);
                                    }
                                }

                                #endregion

                                if (dropInfoList.Count > 0)
                                {
                                    var documentList = new List<DropDocument>();

                                    allCount = dropInfoList.Count;

                                    foreach (var member in dropInfoList)
                                    {
                                        var period = $"{member.Period.Text.Trim()}";
                                        var chineseName = $"{member.ChineseName.Text.Trim()}";
                                        if (string.IsNullOrEmpty(chineseName))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期的中文姓名为空，请检查！");
                                        }
                                        var koreanName = $"{member.KoreanName.Text.Trim()}";
                                        if (string.IsNullOrEmpty(koreanName))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期的韩文姓名为空，请检查！");
                                        }
                                        var student = studentList.FirstOrDefault(s => s.Period.Text.Trim() == member.Period.Text.Trim() && s.ChineseName.Text.Trim() == chineseName);
                                        if (student == null)
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”在信息表中未找到，请检查！");
                                        }
                                        var phone = student.Phone.Text.Trim();
                                        if (string.IsNullOrEmpty(phone))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”在信息表中电话为空，请检查！");
                                        }
                                        var stage = member.CourseStage.Text.Trim().Replace("중등", "中级").Replace("고등", "高级").Replace("새신자", "新家族");
                                        var times = member.CourseTimes.Text.Trim().Replace("과", "课").Replace("회", "回");
                                        if (string.IsNullOrEmpty(stage))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的课程阶段为空，请检查！");
                                        }
                                        if (string.IsNullOrEmpty(times))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的课数为空，请检查！");
                                        }

                                        var lecturer = (stage == "高级" || stage == "新家族") ? member.AdvancedLecturer.Text.Trim() : member.MidLecturer.Text.Trim();
                                        if (string.IsNullOrEmpty(lecturer))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的JS错误，请检查！");
                                        }
                                        var dropCourse = $"{stage}{times}";
                                        var dropDate = member.DropDate.Text.Trim();
                                        if (string.IsNullOrEmpty(dropDate))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的掉落日期为空，请检查！");
                                        }
                                        var visitor = member.Visitor.Text.Trim();
                                        if (string.IsNullOrEmpty(visitor))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的探访者为空，请检查！");
                                        }
                                        var visitorCount = member.VisitorCount.Text.Trim();
                                        if (string.IsNullOrEmpty(visitorCount))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的探访次数为空，请检查！");
                                        }
                                        var dropReason = member.DetailReason.Text.Trim();
                                        if (string.IsNullOrEmpty(dropReason))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的DL事由为空，请检查！");
                                        }
                                        var preacher = member.Preacher.Text.Trim();
                                        if (string.IsNullOrEmpty(preacher))
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的CDS为空，请检查！");
                                        }
                                        // var b = new StringBuilder();
                                        // foreach (var item in workerList)
                                        // {
                                        //     b.AppendLine($"{item.KoreanName.Text} - {item.NameAndPhone.Text}");
                                        // }
                                        // File.WriteAllText(@"W:\log.txt", b.ToString());
                                        var precherInfo = workerList.FirstOrDefault(w => w.KoreanName.Text.Trim() == preacher);
                                        if (precherInfo == null)
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的CDS“{preacher}”在信息表中未找到，请检查！");
                                        }
                                        preacher = precherInfo.NameAndPhone.Text.Trim();
                                        var lecturerInfo = workerList.FirstOrDefault(w => w.KoreanName.Text.Contains(lecturer));
                                        if (lecturerInfo == null)
                                        {
                                            throw new Exception($"掉落信息中{member.Period.Text.Trim()}期“{chineseName}”的JS“{lecturer}”在信息表中未找到，请检查！");
                                        }
                                        lecturer = lecturerInfo.NameAndPhone.Text.Trim();
                                        var ministerKoreanName = this.txt_minister.Text;
                                        if (string.IsNullOrEmpty(ministerKoreanName))
                                        {
                                            throw new Exception($"BZ为空，请检查！");
                                        }
                                        var ministerInfo = workerList.FirstOrDefault(w => w.KoreanName.Text.Contains(ministerKoreanName));
                                        if (ministerInfo == null)
                                        {
                                            throw new Exception($"BZ“{ministerKoreanName}”在信息表中未找到，请检查！");
                                        }
                                        var minister = ministerInfo.NameAndPhone.Text.Trim();

                                        documentList.Add(new DropDocument()
                                        {
                                            Period = period,
                                            ChineseName = chineseName,
                                            KoreanName = koreanName,
                                            Phone = phone,
                                            DropCourse = dropCourse,
                                            DropDate = dropDate,
                                            Visitor = visitor,
                                            VisitorCount = visitorCount,
                                            DropReason = dropReason,
                                            Preacher = preacher,
                                            Lecturer = lecturer,
                                            Minister = minister,
                                            MinisterKoreanName = ministerKoreanName
                                        });

                                        currentCount++;
                                        this.lbl_task.Text = $"{currentCount}/{allCount}";
                                    }

                                    this.gbx_info_file.Visible = true;
                                    this.gbx_drop_file.Visible = true;
                                    this.btn_run.Visible = true;
                                    this.txt_check_result.Text = "处理完成";

                                    Output(documentList);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.gbx_info_file.Visible = true;
                        this.gbx_drop_file.Visible = true;
                        this.btn_run.Visible = true;
                        this.txt_check_result.Text = ex.ToLogString();
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
                thread.Join();
            });
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

        private void btn_run_Click(object sender, EventArgs e)
        {
            Run();
        }
    }

    public class DropInfo
    {
        public ExcelRange ChineseName { get; set; } // A
        public ExcelRange KoreanName { get; set; } // B
        public ExcelRange IDCardBirth { get; set; } // C
        public ExcelRange School { get; set; } // D
        public ExcelRange Period { get; set; } // E
        public ExcelRange DropDate { get; set; } // F
        public ExcelRange MainReason { get; set; } // G
        public ExcelRange DetailReason { get; set; } // H
        public ExcelRange CourseStage { get; set; } // I
        public ExcelRange CourseTimes { get; set; } // J
        public ExcelRange BasicLecturer { get; set; } // K
        public ExcelRange MidLecturer { get; set; } // L
        public ExcelRange AdvancedLecturer { get; set; } // M
        public ExcelRange Preacher { get; set; } // N
        public ExcelRange Visitor { get; set; } // O
        public ExcelRange VisitorCount { get; set; } // P
    }

    public class Student
    {
        public ExcelRange ChineseName { get; set; } // A
        public ExcelRange KoreanName { get; set; } // B
        public ExcelRange IDCardBirth { get; set; } // C
        public ExcelRange Period { get; set; } // D
        public ExcelRange BasicLecturer { get; set; } // E
        public ExcelRange MidLecturer { get; set; } // F
        public ExcelRange AdvancedLecturer { get; set; } // G
        public ExcelRange Preacher { get; set; } // H
        public ExcelRange Phone { get; set; } // I
    }

    public class Worker
    {
        public ExcelRange KoreanName { get; set; } // B
        public ExcelRange NameAndPhone { get; set; } // B
    }

    public class DropDocument
    {
        public string Period { get; set; } // 期数
        public string ChineseName { get;set; } // 中文姓名
        public string KoreanName { get; set; } // 韩文姓名
        public string Phone { get; set; } // 电话
        public string DropCourse { get; set; } // DL课程
        public string DropDate { get; set; } // 掉落日期
        public string Visitor { get; set; } // 探访者
        public string VisitorCount { get; set; } // 探访次数
        public string DropReason { get; set; } // DL事由
        public string Preacher { get; set; } // CDS
        public string Lecturer { get; set; } // JS
        public string Minister { get; set; } // BZ
        public string MinisterKoreanName { get; set; } // BZ
    }
}