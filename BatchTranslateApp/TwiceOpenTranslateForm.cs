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
    public partial class TwiceOpenTranslateForm : ExcelForm
    {
        private int allCount = 0;
        private int currentCount = 0;

        private string once_file = string.Empty;
        private string drop_file = string.Empty;
        private string change_file = string.Empty;

        public TwiceOpenTranslateForm()
        {
            InitializeComponent();
            InitExcelSelector(this.gbx_once_file, this.btn_select_once_file, file => { this.once_file = file; this.lbl_once_file_name.Text = $"已选择：{file}"; });
            InitExcelSelector(this.gbx_drop_file, this.btn_select_drop_file, file => { this.drop_file = file; this.lbl_drop_file_name.Text = $"已选择：{file}"; });
            InitExcelSelector(this.gbx_change_file, this.btn_select_change_file, file => { this.change_file = file; this.lbl_change_file_name.Text = $"已选择：{file}"; });

            // 测试
            this.once_file = @"W:\work\kzj\二开韩文版测试\159，-1，-2一开 杨 测试版.xlsx";
            this.drop_file = @"W:\work\kzj\二开韩文版测试\韩掉 测试版.xlsx";
            this.change_file = @"W:\work\kzj\二开韩文版测试\二开信息变更&地址汇总 测试版.xlsx";
            this.lbl_once_file_name.Text = $"已选择：{once_file}";
            this.lbl_drop_file_name.Text = $"已选择：{drop_file}";
            this.lbl_change_file_name.Text = $"已选择：{change_file}";

            CheckForIllegalCrossThreadCalls = false; // 关闭跨线程调用检查
        }

        public void Output(List<OnceStudent> onceStudentList)
        {
            using (var package = new ExcelPackage(new FileInfo(@"template/一开韩文模板.xlsx")))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[2];
                var periods = new List<string>();

                for (int i = 0; i < onceStudentList.Count; i++)
                {
                    var member = onceStudentList[i];
                    var rowIndex = i + 7;
                    if (!periods.Contains(member.Period.Text))
                    {
                        periods.Add(member.Period.Text);
                    }

                    //sheet.Cells[rowIndex, 1].SetCell(member.Period);
                    //sheet.Cells[rowIndex, 2].SetCell(member.Department);
                    //sheet.Cells[rowIndex, 3].SetCell(member.KoreanName);
                    //sheet.Cells[rowIndex, 4].SetCell(member.ChineseName);
                    //sheet.Cells[rowIndex, 5].SetCell(member.Gender);
                    //sheet.Cells[rowIndex, 6].SetCell(member.IDCardBirth);
                    //if (!string.IsNullOrEmpty(member.ReStudy.Text))
                    //{
                    //    sheet.Cells[rowIndex, 7].SetCell("재수강");
                    //}
                    //sheet.Cells[rowIndex, 8].SetCell(member.ReStudy);
                    //sheet.Cells[rowIndex, 9].SetCell("중화인민공화국"); // cbd数据缺失
                    //sheet.Cells[rowIndex, 10].SetCell(member.Phone);
                    //sheet.Cells[rowIndex, 11].SetCell("중화인민공화국"); // cbd数据缺失
                    //sheet.Cells[rowIndex, 12].SetCell(member.KoreanCity);
                    //sheet.Cells[rowIndex, 13].SetCell(member.CurrentAddress);
                    //sheet.Cells[rowIndex, 14].SetCell(member.IDCardAddress);
                    //sheet.Cells[rowIndex, 15].SetCell(member.Job);
                    //sheet.Cells[rowIndex, 16].SetCell(string.IsNullOrEmpty(member.ReStudy.Text) ? "1" : "2");
                    //sheet.Cells[rowIndex, 17].SetCell(member.CDType);
                    //sheet.Cells[rowIndex, 18].SetCell(member.Times);
                    //sheet.Cells[rowIndex, 19].SetCell("0");
                    //sheet.Cells[rowIndex, 21].SetCell(""); // cbd数据缺失，暂时默认为
                    //sheet.Cells[rowIndex, 22].SetCell(""); // cbd数据缺失，暂时默认为
                    //sheet.Cells[rowIndex, 23].SetCell(""); // cbd数据缺失，暂时默认为
                    //sheet.Cells[rowIndex, 24].SetCell("부산야고보지파");
                    //sheet.Cells[rowIndex, 25].SetCell("중국무한교회");
                    //sheet.Cells[rowIndex, 26].SetCell(member.GuiderDepartment);
                    //sheet.Cells[rowIndex, 27].SetCell(member.GuiderKoreanName);
                    //sheet.Cells[rowIndex, 28].SetCell(member.GuiderNumber);
                    //sheet.Cells[rowIndex, 29].SetCell(member.GuiderSchoolPeriod);
                    //sheet.Cells[rowIndex, 30].SetCell("부산야고보지파");
                    //sheet.Cells[rowIndex, 31].SetCell("중국무한교회");
                    //sheet.Cells[rowIndex, 32].SetCell(member.TeacherDepartment);
                    //sheet.Cells[rowIndex, 33].SetCell(member.TeacherKoreanName);
                    //sheet.Cells[rowIndex, 34].SetCell(member.TeacherNumber);
                    //sheet.Cells[rowIndex, 40].SetCell("1");
                    //sheet.Cells[rowIndex, 42].SetCell("1");

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
                sheet.DeleteRow(onceStudentList.Count + 7, 2304 - onceStudentList.Count); // 第二个参数：行数-1 2308

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

        public void Run()
        {
            #region 表格上传验证
            if (string.IsNullOrEmpty(once_file))
            {
                this.txt_check_result.Text = "请选择一开表格";
                return;
            }
            if (string.IsNullOrEmpty(drop_file))
            {
                this.txt_check_result.Text = "请选择韩掉表格";
                return;
            }
            if (string.IsNullOrEmpty(change_file))
            {
                this.txt_check_result.Text = "请选择韩掉表格";
                return;
            }
            #endregion

            this.gbx_once_file.Visible = false;
            this.gbx_drop_file.Visible = false;
            this.gbx_change_file.Visible = false;
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

                        FileInfo onceFileInfo = new FileInfo(once_file);
                        if (!onceFileInfo.Exists)
                        {
                            throw new Exception("一开表格文件不存在");
                        }

                        FileInfo dropFileInfo = new FileInfo(drop_file);
                        if (!dropFileInfo.Exists)
                        {
                            throw new Exception("韩掉表格文件不存在");
                        }

                        FileInfo changeFileInfo = new FileInfo(change_file);
                        if (!changeFileInfo.Exists)
                        {
                            throw new Exception("变更表格文件不存在");
                        }

                        #endregion

                        var onceStudentList = new List<OnceStudent>();

                        using (var package = new ExcelPackage(onceFileInfo, this.txt_cdb_password.Text))
                        {
                            #region 读取一开表格

                            var sheet = package.Workbook.Worksheets[2];

                            // 遍历工作表的每一行（从第2行开始，跳过标题行）
                            for (int rowIndex = 7; rowIndex <= sheet.Dimension.End.Row; rowIndex++)  // EPPlus的行从1开始，索引为1表示第二行
                            {
                                var onceStudent = new OnceStudent()
                                {
                                    Period = sheet.Cells[rowIndex, 1],  // 读取第1列（A列）
                                    Department = sheet.Cells[rowIndex, 2],  // 读取第2列（B列）
                                    KoreanName = sheet.Cells[rowIndex, 3],  // 读取第3列（C列）
                                    ChineseName = sheet.Cells[rowIndex, 4],  // 读取第4列（D列）
                                    Gender = sheet.Cells[rowIndex, 5],  // 读取第5列（E列）
                                    IDCardBirth = sheet.Cells[rowIndex, 6],  // 读取第6列（F列）
                                    ReStudy = sheet.Cells[rowIndex, 7],  // 读取第7列（G列）
                                    ReStudyDesc = sheet.Cells[rowIndex, 8],  // 读取第8列（H列）
                                    Country = sheet.Cells[rowIndex, 9],  // 读取第9列（I列）
                                    Phone = sheet.Cells[rowIndex, 10],  // 读取第10列（J列）
                                    CurrentAddressCountry = sheet.Cells[rowIndex, 11],  // 读取第11列（K列）
                                    CurrentAddressCity = sheet.Cells[rowIndex, 12],  // 读取第12列（L列）
                                    CurrentAddress = sheet.Cells[rowIndex, 13],  // 读取第13列（M列）
                                    IDCardAddress = sheet.Cells[rowIndex, 14],  // 读取第14列（N列）
                                    Job = sheet.Cells[rowIndex, 15],  // 读取第15列（O列）
                                    IsFirstTime = sheet.Cells[rowIndex, 16],  // 读取第16列（P列）
                                    CDType = sheet.Cells[rowIndex, 17],  // 读取第17列（Q列）
                                    RoomTimes = sheet.Cells[rowIndex, 18],  // 读取第18列（R列）
                                    RoomOnce = sheet.Cells[rowIndex, 19],  // 读取第19列（S列）
                                    RoomEnd = sheet.Cells[rowIndex, 20],  // 读取第20列（T列）
                                    PreacherName = sheet.Cells[rowIndex, 21],  // 读取第21列（U列）
                                    PreacherNameOnce = sheet.Cells[rowIndex, 22],  // 读取第22列（V列）
                                    PreacherNameEnd = sheet.Cells[rowIndex, 23],  // 读取第23列（W列）
                                    GuiderBranch = sheet.Cells[rowIndex, 24],  // 读取第24列（X列）
                                    GuiderChurch = sheet.Cells[rowIndex, 25],  // 读取第25列（Y列）
                                    GuiderDepartment = sheet.Cells[rowIndex, 26],  // 读取第26列（Z列）
                                    GuiderName = sheet.Cells[rowIndex, 27],  // 读取第27列（AA列）
                                    GuiderNumber = sheet.Cells[rowIndex, 28],  // 读取第28列（AB列）
                                    GuiderPeriod = sheet.Cells[rowIndex, 29],  // 读取第29列（AC列）
                                    TeacherBranch = sheet.Cells[rowIndex, 30],  // 读取第30列（AD列）
                                    TeacherChurch = sheet.Cells[rowIndex, 31],  // 读取第31列（AE列）
                                    TeacherDepartment = sheet.Cells[rowIndex, 32],  // 读取第32列（AF列）
                                    TeacherName = sheet.Cells[rowIndex, 33],  // 读取第33列（AG列）
                                    TeacherNumber = sheet.Cells[rowIndex, 34],  // 读取第34列（AH列）
                                    LeafBranch = sheet.Cells[rowIndex, 35],  // 读取第35列（AI列）
                                    LeafChurch = sheet.Cells[rowIndex, 36],  // 读取第36列（AJ列）
                                    LeafDepartment = sheet.Cells[rowIndex, 37],  // 读取第37列（AK列）
                                    LeafName = sheet.Cells[rowIndex, 38],  // 读取第38列（AL列）
                                    LeafNumber = sheet.Cells[rowIndex, 39],  // 读取第39列（AM列）
                                    OnceOpened = sheet.Cells[rowIndex, 40],  // 读取第40列（AN列）
                                    IsOffline = sheet.Cells[rowIndex, 41],  // 读取第41列（AO列）
                                    IsEnding = sheet.Cells[rowIndex, 42],  // 读取第42列（AP列）
                                    OpenMoveAdd = sheet.Cells[rowIndex, 43],  // 读取第43列（AQ列）
                                    OpenMoveSubtract = sheet.Cells[rowIndex, 44],  // 读取第44列（AR列）
                                    OpenDropOutAdd = sheet.Cells[rowIndex, 45],  // 读取第45列（AS列）
                                    OpenDropOutSubtract = sheet.Cells[rowIndex, 46],  // 读取第46列（AT列）
                                    SignInNumber = sheet.Cells[rowIndex, 47],  // 读取第47列（AU列）
                                    EndingNumber = sheet.Cells[rowIndex, 48],  // 读取第48列（AV列）
                                    MediumDropNumber = sheet.Cells[rowIndex, 49],  // 读取第49列（AW列）
                                    DropType = sheet.Cells[rowIndex, 50],  // 读取第50列（AX列）
                                    DropStep = sheet.Cells[rowIndex, 51],  // 读取第51列（AY列）
                                    DropReason = sheet.Cells[rowIndex, 52],  // 读取第52列（AZ列）
                                    DropManage = sheet.Cells[rowIndex, 53],  // 读取第53列（BA列）
                                    DropOutMoveDetail = sheet.Cells[rowIndex, 54],  // 读取第54列（BB列）
                                    DropReasonDesc = sheet.Cells[rowIndex, 55],  // 读取第55列（BC列）
                                    AbsencesNumber = sheet.Cells[rowIndex, 56],  // 读取第56列（BD列）
                                    AverageGrade = sheet.Cells[rowIndex, 57],  // 读取第57列（BE列）
                                    StudyFruitDetail = sheet.Cells[rowIndex, 58],  // 读取第58列（BF列）
                                    StudyFruitNumber = sheet.Cells[rowIndex, 59],  // 读取第59列（BG列）
                                    StudyRoomDetail = sheet.Cells[rowIndex, 60],  // 读取第60列（BH列）
                                    StudyRoomNumber = sheet.Cells[rowIndex, 61],  // 读取第61列（BI列）
                                };

                                // 如果存在职位信息，转换为韩文
                                if (!string.IsNullOrEmpty(onceStudent.Job.Text))
                                {
                                    // student.Job.Value = TranslateHelper.SentenceToKorean(student.Job.Text);
                                }

                                // 如果Period不为空则加入到列表
                                if (!string.IsNullOrEmpty(onceStudent.Period.Text.Trim()))
                                {
                                    onceStudentList.Add(onceStudent);
                                }
                            }

                            #endregion

                            if (onceStudentList.Count > 0)
                            {
                                #region 翻译

                                allCount = onceStudentList.Count;
                                var taskBuilder = new StringBuilder();
                                foreach (var member in onceStudentList)
                                {

                                    currentCount++;
                                    this.lbl_task.Text = $"{currentCount}/{allCount}";
                                }
                                this.gbx_once_file.Visible = true;
                                this.gbx_drop_file.Visible = true;
                                this.gbx_change_file.Visible = true;
                                this.btn_run.Visible = true;
                                this.txt_check_result.Text = "翻译完成";

                                #endregion

                                #region 排序

                                //var departmentSortDic = new Dictionary<string, int>
                                //{
                                //    { "자문", 1 },
                                //    { "장년", 2 },
                                //    { "부녀", 3 },
                                //    { "청년", 4 }
                                //};

                                //const int defaultSortValue = int.MaxValue;

                                //var sortedList = list
                                //    .OrderBy(student => student.Period.Text)
                                //    .ThenBy(student =>
                                //    {
                                //        // 使用字典中的值进行排序，如果不在字典中则使用默认排序值
                                //        return departmentSortDic.ContainsKey(student.Department.Text)
                                //            ? departmentSortDic[student.Department.Text]
                                //            : defaultSortValue;
                                //    })
                                //    .ThenBy(student => student.IDCardBirth.Text) // 按年龄从大到小排序
                                //    .ToList();

                                #endregion

                                Output(onceStudentList);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.gbx_once_file.Visible = true;
                        this.gbx_drop_file.Visible = true;
                        this.gbx_change_file.Visible = true;
                        this.btn_run.Visible = true;
                        this.txt_check_result.Text = ex.Message;
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

    public class OnceStudent
    {
        public ExcelRange Period { get; set; } // A
        public ExcelRange Department { get; set; } // B
        public ExcelRange KoreanName { get; set; } // C
        public ExcelRange ChineseName { get; set; } // D
        public ExcelRange Gender { get; set; } // E
        public ExcelRange IDCardBirth { get; set; } // F
        public ExcelRange ReStudy { get; set; } // G
        public ExcelRange ReStudyDesc { get; set; } // H
        public ExcelRange Country { get; set; } // I
        public ExcelRange Phone { get; set; } // J
        public ExcelRange CurrentAddressCountry { get; set; } // K
        public ExcelRange CurrentAddressCity { get; set; } // L
        public ExcelRange CurrentAddress { get; set; } // M
        public ExcelRange IDCardAddress { get; set; } // N
        public ExcelRange Job { get; set; } // O
        public ExcelRange IsFirstTime { get; set; } // P
        public ExcelRange CDType { get; set; } // Q
        public ExcelRange RoomTimes { get; set; } // R
        public ExcelRange RoomOnce { get; set; } // S
        public ExcelRange RoomEnd { get; set; } // T
        public ExcelRange PreacherName { get; set; } // U
        public ExcelRange PreacherNameOnce { get; set; } // V
        public ExcelRange PreacherNameEnd { get; set; } // W
        public ExcelRange GuiderBranch { get; set; } // X
        public ExcelRange GuiderChurch { get; set; } // Y
        public ExcelRange GuiderDepartment { get; set; } // Z
        public ExcelRange GuiderName { get; set; } // AA
        public ExcelRange GuiderNumber { get; set; } // AB
        public ExcelRange GuiderPeriod { get; set; } // AC
        public ExcelRange TeacherBranch { get; set; } // AD
        public ExcelRange TeacherChurch { get; set; } // AE
        public ExcelRange TeacherDepartment { get; set; } // AF
        public ExcelRange TeacherName { get; set; } // AG
        public ExcelRange TeacherNumber { get; set; } // AH
        public ExcelRange LeafBranch { get; set; } // AI
        public ExcelRange LeafChurch { get; set; } // AJ
        public ExcelRange LeafDepartment { get; set; } // AK
        public ExcelRange LeafName { get; set; } // AL
        public ExcelRange LeafNumber { get; set; } // AM
        public ExcelRange OnceOpened { get; set; } // AN
        public ExcelRange IsOffline { get; set; } // AO
        public ExcelRange IsEnding { get; set; } // AP
        public ExcelRange OpenMoveAdd { get; set; } // AQ
        public ExcelRange OpenMoveSubtract { get; set; } // AR
        public ExcelRange OpenDropOutAdd { get; set; } // AS
        public ExcelRange OpenDropOutSubtract { get; set; } // AT
        public ExcelRange SignInNumber { get; set; } // AU
        public ExcelRange EndingNumber { get; set; } // AV
        public ExcelRange MediumDropNumber { get; set; } // AW
        public ExcelRange DropType { get; set; } // AX
        public ExcelRange DropStep { get; set; } // AY
        public ExcelRange DropReason { get; set; } // AZ
        public ExcelRange DropManage { get; set; } // BA
        public ExcelRange DropOutMoveDetail { get; set; } // BB
        public ExcelRange DropReasonDesc { get; set; } // BC
        public ExcelRange AbsencesNumber { get; set; } // BD
        public ExcelRange AverageGrade { get; set; } // BE
        public ExcelRange StudyFruitDetail { get; set; } // BF
        public ExcelRange StudyFruitNumber { get; set; } // BG
        public ExcelRange StudyRoomDetail { get; set; } // BH
        public ExcelRange StudyRoomNumber { get; set; } // BI
    }

    public class DropStudent
    {
        public ExcelRange School { get; set; } // A
        public ExcelRange Period { get; set; } // B
        public ExcelRange KoreanName { get; set; } // C
        public ExcelRange ChineseName { get; set; } // D
        public ExcelRange Category { get; set; } // E
    }
}
