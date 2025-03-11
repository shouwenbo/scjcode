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
using System.Reflection;

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
            // this.once_file = @"W:\work\kzj\二开韩文版测试\159，-1，-2一开 杨 测试版.xlsx";
            // this.drop_file = @"W:\work\kzj\二开韩文版测试\韩掉 测试版.xlsx";
            // this.change_file = @"W:\work\kzj\二开韩文版测试\二开信息变更&地址汇总 测试版.xlsx";
            // this.lbl_once_file_name.Text = $"已选择：{once_file}";
            // this.lbl_drop_file_name.Text = $"已选择：{drop_file}";
            // this.lbl_change_file_name.Text = $"已选择：{change_file}";
            // this.txt_once_password.Text = "wh12000";
            // this.txt_drop_password.Text = "wh0217";
            // this.txt_change_password.Text = "wh0217";

            CheckForIllegalCrossThreadCalls = false; // 关闭跨线程调用检查
        }

        public void Output(List<OnceStudent> list)
        {
            using (var package = new ExcelPackage(new FileInfo(@"template/二开韩文模板（未加密）.xlsx")))
            //using (var package = new ExcelPackage(new FileInfo(@"template/二开韩文模板.xlsx"), "wh12000"))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[2];
                var periods = new List<string>();

                var grbg = new List<string>();

                for (int i = 0; i < list.Count; i++)
                {
                    var member = list[i];
                    var rowIndex = i + 7;
                    if (!periods.Contains(member.Period.Text))
                    {
                        periods.Add(member.Period.Text);
                    }

                    sheet.Cells[rowIndex, 1].SetCell(member.Period);  // 设置第1列（A列）
                    sheet.Cells[rowIndex, 2].SetCell(member.Department);  // 设置第2列（B列）
                    sheet.Cells[rowIndex, 3].SetCell(member.KoreanName);  // 设置第3列（C列）
                    sheet.Cells[rowIndex, 4].SetCell(member.ChineseName);  // 设置第4列（D列）
                    sheet.Cells[rowIndex, 5].SetCell(member.Gender);  // 设置第5列（E列）
                    sheet.Cells[rowIndex, 6].SetCell(member.IDCardBirth);  // 设置第6列（F列）
                    sheet.Cells[rowIndex, 7].SetCell(member.ReStudy);  // 设置第7列（G列）
                    sheet.Cells[rowIndex, 8].SetCell(member.ReStudyDesc);  // 设置第8列（H列）
                    sheet.Cells[rowIndex, 9].SetCell(member.Country);  // 设置第9列（I列）
                    sheet.Cells[rowIndex, 10].SetCell(member.Phone);  // 设置第10列（J列）
                    sheet.Cells[rowIndex, 11].SetCell(member.CurrentAddressCountry);  // 设置第11列（K列）
                    sheet.Cells[rowIndex, 12].SetCell(member.CurrentAddressCity);  // 设置第12列（L列）
                    sheet.Cells[rowIndex, 13].SetCell(member.CurrentAddress);  // 设置第13列（M列）
                    sheet.Cells[rowIndex, 14].SetCell(member.IDCardAddress);  // 设置第14列（N列）
                    sheet.Cells[rowIndex, 15].SetCell(member.Job);  // 设置第15列（O列）
                    sheet.Cells[rowIndex, 16].SetCell(member.IsFirstTime);  // 设置第16列（P列）
                    sheet.Cells[rowIndex, 17].SetCell(member.CDType);  // 设置第17列（Q列）
                    sheet.Cells[rowIndex, 18].SetCell(member.RoomTimes);  // 设置第18列（R列）
                    sheet.Cells[rowIndex, 19].SetCell(member.RoomOnce);  // 设置第19列（S列）
                    sheet.Cells[rowIndex, 20].SetCell(member.RoomEnd);  // 设置第20列（T列）
                    sheet.Cells[rowIndex, 21].SetCell(member.PreacherName);  // 设置第21列（U列）
                    sheet.Cells[rowIndex, 22].SetCell(member.PreacherNameOnce);  // 设置第22列（V列）
                    sheet.Cells[rowIndex, 23].SetCell(member.PreacherNameEnd);  // 设置第23列（W列）
                    sheet.Cells[rowIndex, 24].SetCell(member.GuiderBranch);  // 设置第24列（X列）
                    sheet.Cells[rowIndex, 25].SetCell(member.GuiderChurch);  // 设置第25列（Y列）
                    sheet.Cells[rowIndex, 26].SetCell(member.GuiderDepartment);  // 设置第26列（Z列）
                    sheet.Cells[rowIndex, 27].SetCell(member.GuiderName);  // 设置第27列（AA列）
                    sheet.Cells[rowIndex, 28].SetCell(member.GuiderNumber);  // 设置第28列（AB列）

                    grbg.Add(member.GuiderNumber.Style.Fill.BackgroundColor.Rgb);

                    sheet.Cells[rowIndex, 29].SetCell(member.GuiderPeriod);  // 设置第29列（AC列）
                    sheet.Cells[rowIndex, 30].SetCell(member.TeacherBranch);  // 设置第30列（AD列）
                    sheet.Cells[rowIndex, 31].SetCell(member.TeacherChurch);  // 设置第31列（AE列）
                    sheet.Cells[rowIndex, 32].SetCell(member.TeacherDepartment);  // 设置第32列（AF列）
                    sheet.Cells[rowIndex, 33].SetCell(member.TeacherName);  // 设置第33列（AG列）
                    sheet.Cells[rowIndex, 34].SetCell(member.TeacherNumber);  // 设置第34列（AH列）
                    sheet.Cells[rowIndex, 35].SetCell(member.LeafBranch);  // 设置第35列（AI列）
                    sheet.Cells[rowIndex, 36].SetCell(member.LeafChurch);  // 设置第36列（AJ列）
                    sheet.Cells[rowIndex, 37].SetCell(member.LeafDepartment);  // 设置第37列（AK列）
                    sheet.Cells[rowIndex, 38].SetCell(member.LeafName);  // 设置第38列（AL列）
                    sheet.Cells[rowIndex, 39].SetCell(member.LeafNumber);  // 设置第39列（AM列）
                    sheet.Cells[rowIndex, 40].SetCell(member.OnceOpened);  // 设置第40列（AN列）
                    sheet.Cells[rowIndex, 41].SetCell(member.IsOffline);  // 设置第41列（AO列）
                    sheet.Cells[rowIndex, 42].SetCell(member.IsEnding);  // 设置第42列（AP列）
                    sheet.Cells[rowIndex, 43].SetCell(member.OpenMoveAdd);  // 设置第43列（AQ列）
                    sheet.Cells[rowIndex, 44].SetCell(member.OpenMoveSubtract);  // 设置第44列（AR列）
                    sheet.Cells[rowIndex, 45].SetCell(member.OpenDropOutAdd);  // 设置第45列（AS列）
                    sheet.Cells[rowIndex, 46].SetCell(member.OpenDropOutSubtract);  // 设置第46列（AT列）
                    // sheet.Cells[rowIndex, 47].SetCell(member.SignInNumber);  // 设置第47列（AU列）
                    sheet.Cells[rowIndex, 48].SetCell(member.EndingNumber);  // 设置第48列（AV列）
                    sheet.Cells[rowIndex, 49].SetCell(member.MediumDropNumber);  // 设置第49列（AW列）
                    sheet.Cells[rowIndex, 50].SetCell(member.DropType);  // 设置第50列（AX列）
                    sheet.Cells[rowIndex, 51].SetCell(member.DropStep);  // 设置第51列（AY列）
                    sheet.Cells[rowIndex, 52].SetCell(member.DropReason);  // 设置第52列（AZ列）
                    sheet.Cells[rowIndex, 53].SetCell(member.DropManage);  // 设置第53列（BA列）
                    sheet.Cells[rowIndex, 54].SetCell(member.DropOutMoveDetail);  // 设置第54列（BB列）
                    sheet.Cells[rowIndex, 55].SetCell(member.DropReasonDesc);  // 设置第55列（BC列）
                    sheet.Cells[rowIndex, 56].SetCell(member.AbsencesNumber);  // 设置第56列（BD列）
                    sheet.Cells[rowIndex, 57].SetCell(member.AverageGrade);  // 设置第57列（BE列）
                    sheet.Cells[rowIndex, 58].SetCell(member.StudyFruitDetail);  // 设置第58列（BF列）
                    sheet.Cells[rowIndex, 59].SetCell(member.StudyFruitNumber);  // 设置第59列（BG列）
                    sheet.Cells[rowIndex, 60].SetCell(member.StudyRoomDetail);  // 设置第60列（BH列）
                    sheet.Cells[rowIndex, 61].SetCell(member.StudyRoomNumber);  // 设置第61列（BI列）

                }
                sheet.DeleteRow(list.Count + 7, 3014 - list.Count); // 第二个参数：行数-1

                //Console.WriteLine(grbg);
                //File.WriteAllText(@"W:\work\kzj\二开韩文版测试\log_grbg.txt", string.Join("\r\n", grbg));
                //Console.WriteLine();

                //sheet.DeleteRowsAndShiftUp(new List<int>() { 6 }); 
                //if (list.Count > 0)
                //{
                //    sheet.GetRow(6).GetCell(0).SetCellValueIfNotEmpty(list[0].Period);
                //}

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "保存 Excel 文件";
                    saveFileDialog.FileName = $"二开（{string.Join(" ", periods)}） 韩.xlsx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string targetFilePath = saveFileDialog.FileName;
                        //package.Workbook.Worksheets[2].Cells[10, 10].Style.Fill.SetBackground(System.Drawing.Color.LightGreen);
                        //package.Workbook.Worksheets[2].Cells[10, 10].Style.Font.Color.SetColor(System.Drawing.Color.LightGreen);
                        //package.Workbook.Worksheets[2].Cells[10, 10].Style.Font.SetFromFont("微软雅黑", 44, true, true, true, true);

                        package.SaveAs(new FileInfo(targetFilePath));
                        this.txt_check_result.Text = $"文件已成功保存到：{targetFilePath}";
























                        //// 打开 Excel 文件
                        //using (var fileStream = new FileStream(targetFilePath, FileMode.Open, FileAccess.ReadWrite))
                        //{
                        //    XSSFWorkbook workbook = new XSSFWorkbook(fileStream);  // 读取 Excel 工作簿
                        //    ISheet npoisheet = workbook.GetSheetAt(2);  // 获取第二个工作表（索引从 0 开始）

                        //    // 创建字体
                        //    IFont font = workbook.CreateFont();
                        //    font.Color = IndexedColors.Red.Index;  // 设置字体颜色为红色
                        //    font.FontHeightInPoints = 12;  // 设置字体大小为12

                        //    // 获取 ChineseName 列（假设是第 4 列，索引为 3）
                        //    int chineseNameColumnIndex = 3;
                        //    int rowIndex = 6; // 从第7行开始（索引从0开始，6代表第7行）

                        //    // 遍历所有行
                        //    for (int i = rowIndex; i <= npoisheet.LastRowNum; i++)
                        //    {
                        //        IRow row = npoisheet.GetRow(i);
                        //        if (row != null)
                        //        {
                        //            ICell cell = row.GetCell(chineseNameColumnIndex);  // 获取ChineseName列的单元格
                        //            if (cell != null)
                        //            {
                        //                // 设置字体样式
                        //                var cellStyle = workbook.CreateCellStyle();
                        //                cellStyle.SetFont(font);
                        //                cell.CellStyle = cellStyle;
                        //            }
                        //        }
                        //    }

                        //    // 保存修改后的文件
                        //    using (var outputStream = new FileStream(targetFilePath, FileMode.Create, FileAccess.Write))
                        //    {
                        //        workbook.Write(outputStream);
                        //    }
                        //}























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
                        var dropStudentList = new List<DropStudent>();
                        var changeStudentList = new List<ChangeStudent>();

                        using (var oncePackage = new ExcelPackage(onceFileInfo, this.txt_once_password.Text))
                        {
                            using (var dropPackage = new ExcelPackage(dropFileInfo, this.txt_drop_password.Text))
                            {
                                using (var changePackage = new ExcelPackage(changeFileInfo, this.txt_change_password.Text))
                                {
                                    #region 读取韩掉表

                                    var dropSheet = dropPackage.Workbook.Worksheets[0];
                                    for (int rowIndex = 1; rowIndex <= dropSheet.Dimension.End.Row; rowIndex++)
                                    {
                                        var dropStudent = new DropStudent()
                                        {
                                            School = dropSheet.Cells[rowIndex, 1],
                                            Period = dropSheet.Cells[rowIndex, 2],
                                            KoreanName = dropSheet.Cells[rowIndex, 3],
                                            ChineseName = dropSheet.Cells[rowIndex, 4],
                                            DropType = dropSheet.Cells[rowIndex, 5],
                                            DropStep = dropSheet.Cells[rowIndex, 6],
                                            DropReason = dropSheet.Cells[rowIndex, 7],
                                            DropManage = dropSheet.Cells[rowIndex, 8],
                                            DropOutMoveDetail = dropSheet.Cells[rowIndex, 9],
                                            DropReasonDesc = dropSheet.Cells[rowIndex, 10]
                                        };

                                        if (Regex.IsMatch(dropStudent.Period.Text.Trim(), @"^\d{3}(-\d{1})?$"))
                                        {
                                            dropStudentList.Add(dropStudent);
                                        }
                                    }

                                    #endregion

                                    #region 读取变更表

                                    var changeSheet = changePackage.Workbook.Worksheets[0];
                                    for (int rowIndex = 1; rowIndex <= changeSheet.Dimension.End.Row; rowIndex++)
                                    {
                                        var changeStudent = new ChangeStudent()
                                        {
                                            Period = changeSheet.Cells[rowIndex, 1],
                                            ChineseName = changeSheet.Cells[rowIndex, 2],
                                            NewChineseName = changeSheet.Cells[rowIndex, 3],
                                            KoreanName = changeSheet.Cells[rowIndex, 4],
                                            NewKoreanName = changeSheet.Cells[rowIndex, 5],
                                            IDCardBirth = changeSheet.Cells[rowIndex, 6],
                                            NewIDCardBirth = changeSheet.Cells[rowIndex, 7],
                                            Phone = changeSheet.Cells[rowIndex, 8],
                                            NewPhone = changeSheet.Cells[rowIndex, 9],
                                            Country = changeSheet.Cells[rowIndex, 10],
                                            CurrentAddressCountry = changeSheet.Cells[rowIndex, 11],
                                            CurrentAddressCity = changeSheet.Cells[rowIndex, 12],
                                            CurrentAddressCityKorean = changeSheet.Cells[rowIndex, 13],
                                            CurrentAddress = changeSheet.Cells[rowIndex, 14],
                                            IDCardAddress = changeSheet.Cells[rowIndex, 15],
                                        };

                                        if (Regex.IsMatch(changeStudent.Period.Text.Trim(), @"^\d{3}(-\d{1})?$"))
                                        {
                                            changeStudentList.Add(changeStudent);
                                        }
                                    }

                                    #endregion

                                    #region 读取一开表格

                                    var onceSheet = oncePackage.Workbook.Worksheets[2];
                                    // 遍历工作表的每一行（从第2行开始，跳过标题行）
                                    for (int rowIndex = 7; rowIndex <= onceSheet.Dimension.End.Row; rowIndex++)  // EPPlus的行从1开始，索引为1表示第二行
                                    {
                                        var onceStudent = new OnceStudent()
                                        {
                                            Period = onceSheet.Cells[rowIndex, 1],  // 读取第1列（A列）
                                            Department = onceSheet.Cells[rowIndex, 2],  // 读取第2列（B列）
                                            KoreanName = onceSheet.Cells[rowIndex, 3],  // 读取第3列（C列）
                                            ChineseName = onceSheet.Cells[rowIndex, 4],  // 读取第4列（D列）
                                            Gender = onceSheet.Cells[rowIndex, 5],  // 读取第5列（E列）
                                            IDCardBirth = onceSheet.Cells[rowIndex, 6],  // 读取第6列（F列）
                                            ReStudy = onceSheet.Cells[rowIndex, 7],  // 读取第7列（G列）
                                            ReStudyDesc = onceSheet.Cells[rowIndex, 8],  // 读取第8列（H列）
                                            Country = onceSheet.Cells[rowIndex, 9],  // 读取第9列（I列）
                                            Phone = onceSheet.Cells[rowIndex, 10],  // 读取第10列（J列）
                                            CurrentAddressCountry = onceSheet.Cells[rowIndex, 11],  // 读取第11列（K列）
                                            CurrentAddressCity = onceSheet.Cells[rowIndex, 12],  // 读取第12列（L列）
                                            CurrentAddress = onceSheet.Cells[rowIndex, 13],  // 读取第13列（M列）
                                            IDCardAddress = onceSheet.Cells[rowIndex, 14],  // 读取第14列（N列）
                                            Job = onceSheet.Cells[rowIndex, 15],  // 读取第15列（O列）
                                            IsFirstTime = onceSheet.Cells[rowIndex, 16],  // 读取第16列（P列）
                                            CDType = onceSheet.Cells[rowIndex, 17],  // 读取第17列（Q列）
                                            RoomTimes = onceSheet.Cells[rowIndex, 18],  // 读取第18列（R列）
                                            RoomOnce = onceSheet.Cells[rowIndex, 19],  // 读取第19列（S列）
                                            RoomEnd = onceSheet.Cells[rowIndex, 20],  // 读取第20列（T列）
                                            PreacherName = onceSheet.Cells[rowIndex, 21],  // 读取第21列（U列）
                                            PreacherNameOnce = onceSheet.Cells[rowIndex, 22],  // 读取第22列（V列）
                                            PreacherNameEnd = onceSheet.Cells[rowIndex, 23],  // 读取第23列（W列）
                                            GuiderBranch = onceSheet.Cells[rowIndex, 24],  // 读取第24列（X列）
                                            GuiderChurch = onceSheet.Cells[rowIndex, 25],  // 读取第25列（Y列）
                                            GuiderDepartment = onceSheet.Cells[rowIndex, 26],  // 读取第26列（Z列）
                                            GuiderName = onceSheet.Cells[rowIndex, 27],  // 读取第27列（AA列）
                                            GuiderNumber = onceSheet.Cells[rowIndex, 28],  // 读取第28列（AB列）
                                            GuiderPeriod = onceSheet.Cells[rowIndex, 29],  // 读取第29列（AC列）
                                            TeacherBranch = onceSheet.Cells[rowIndex, 30],  // 读取第30列（AD列）
                                            TeacherChurch = onceSheet.Cells[rowIndex, 31],  // 读取第31列（AE列）
                                            TeacherDepartment = onceSheet.Cells[rowIndex, 32],  // 读取第32列（AF列）
                                            TeacherName = onceSheet.Cells[rowIndex, 33],  // 读取第33列（AG列）
                                            TeacherNumber = onceSheet.Cells[rowIndex, 34],  // 读取第34列（AH列）
                                            LeafBranch = onceSheet.Cells[rowIndex, 35],  // 读取第35列（AI列）
                                            LeafChurch = onceSheet.Cells[rowIndex, 36],  // 读取第36列（AJ列）
                                            LeafDepartment = onceSheet.Cells[rowIndex, 37],  // 读取第37列（AK列）
                                            LeafName = onceSheet.Cells[rowIndex, 38],  // 读取第38列（AL列）
                                            LeafNumber = onceSheet.Cells[rowIndex, 39],  // 读取第39列（AM列）
                                            OnceOpened = onceSheet.Cells[rowIndex, 40],  // 读取第40列（AN列）
                                            IsOffline = onceSheet.Cells[rowIndex, 41],  // 读取第41列（AO列）
                                            IsEnding = onceSheet.Cells[rowIndex, 42],  // 读取第42列（AP列）
                                            OpenMoveAdd = onceSheet.Cells[rowIndex, 43],  // 读取第43列（AQ列）
                                            OpenMoveSubtract = onceSheet.Cells[rowIndex, 44],  // 读取第44列（AR列）
                                            OpenDropOutAdd = onceSheet.Cells[rowIndex, 45],  // 读取第45列（AS列）
                                            OpenDropOutSubtract = onceSheet.Cells[rowIndex, 46],  // 读取第46列（AT列）
                                            SignInNumber = onceSheet.Cells[rowIndex, 47],  // 读取第47列（AU列）
                                            EndingNumber = onceSheet.Cells[rowIndex, 48],  // 读取第48列（AV列）
                                            MediumDropNumber = onceSheet.Cells[rowIndex, 49],  // 读取第49列（AW列）
                                            DropType = onceSheet.Cells[rowIndex, 50],  // 读取第50列（AX列）
                                            DropStep = onceSheet.Cells[rowIndex, 51],  // 读取第51列（AY列）
                                            DropReason = onceSheet.Cells[rowIndex, 52],  // 读取第52列（AZ列）
                                            DropManage = onceSheet.Cells[rowIndex, 53],  // 读取第53列（BA列）
                                            DropOutMoveDetail = onceSheet.Cells[rowIndex, 54],  // 读取第54列（BB列）
                                            DropReasonDesc = onceSheet.Cells[rowIndex, 55],  // 读取第55列（BC列）
                                            AbsencesNumber = onceSheet.Cells[rowIndex, 56],  // 读取第56列（BD列）
                                            AverageGrade = onceSheet.Cells[rowIndex, 57],  // 读取第57列（BE列）
                                            StudyFruitDetail = onceSheet.Cells[rowIndex, 58],  // 读取第58列（BF列）
                                            StudyFruitNumber = onceSheet.Cells[rowIndex, 59],  // 读取第59列（BG列）
                                            StudyRoomDetail = onceSheet.Cells[rowIndex, 60],  // 读取第60列（BH列）
                                            StudyRoomNumber = onceSheet.Cells[rowIndex, 61],  // 读取第61列（BI列）
                                        };

                                        // 如果Period不为空则加入到列表
                                        if (Regex.IsMatch(onceStudent.Period.Text.Trim(), @"^\d{3}(-\d{1})?$"))
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

                                        #region 掉落信息设置

                                        foreach (var member in onceStudentList)
                                        {
                                            var dropMember = dropStudentList.FirstOrDefault(p => p.Period.Text == member.Period.Text && p.ChineseName.Text == member.ChineseName.Text);
                                            if (dropMember != null)
                                            {
                                                if (dropMember.DropType.Text != "초등")
                                                {
                                                    member.DropType.SetCell(dropMember.DropType);
                                                }
                                                member.DropStep.SetCell(dropMember.DropStep);
                                                member.DropReason.SetCell(dropMember.DropReason);
                                                member.DropManage.SetCell(dropMember.DropManage);
                                                member.DropReasonDesc.SetCell(dropMember.DropReasonDesc);
                                                member.IsEnding.SetCell("", true);
                                            }

                                            /*
                                             
                                             SignInNumber = onceSheet.Cells[rowIndex, 47],  // 读取第47列（AU列）
                                            EndingNumber = onceSheet.Cells[rowIndex, 48],  // 读取第48列（AV列）
                                            MediumDropNumber = onceSheet.Cells[rowIndex, 49],  // 读取第49列（AW列）
                                            DropType = onceSheet.Cells[rowIndex, 50],  // 读取第50列（AX列）
                                            DropStep = onceSheet.Cells[rowIndex, 51],  // 读取第51列（AY列）
                                            DropReason = onceSheet.Cells[rowIndex, 52],  // 读取第52列（AZ列）
                                            DropManage = onceSheet.Cells[rowIndex, 53],  // 读取第53列（BA列）
                                            DropOutMoveDetail = onceSheet.Cells[rowIndex, 54],  // 读取第54列（BB列）
                                            DropReasonDesc = onceSheet.Cells[rowIndex, 55],  // 读取第55列（BC列）
                                             
                                             */
                                        }

                                        #endregion

                                        #region 重新计算部署

                                        foreach (var member in onceStudentList)
                                        {
                                            if (string.IsNullOrEmpty(member.DropStep.Text))
                                            {
                                                if (int.TryParse(member.IDCardBirth.Text.Substring(0, 4), out int birthYear))
                                                {
                                                    var age = DateTime.Now.Year - birthYear;

                                                    // 壮年/妇女 → 老年
                                                    if ((member.Department.Text == "장년" || member.Department.Text == "부녀") && age >= 70)
                                                    {
                                                        member.Department.SetCell("자문");
                                                        member.Department.MarkYellow();
                                                    }

                                                    // 壮年/妇女/老年 → 青年
                                                    if (member.OriAge >= 40 && age < 40)
                                                    {
                                                        member.Department.MarkGreen();
                                                        member.Department.AddComment("部署计算失败，因为不知道是否有子女，请手动调整");
                                                    }

                                                    // 老年 → 壮年/妇女
                                                    if ((member.Department.Text == "자문") && (age >= 40 && age < 70))
                                                    {
                                                        member.Department.SetCell(member.Gender.Text == "남" ? "장년" : "부녀");
                                                        member.Department.MarkYellow();
                                                    }

                                                    // 青年 → 壮年/妇女
                                                    if ((member.Department.Text == "청년") && (age >= 40 && age < 70))
                                                    {
                                                        member.Department.SetCell(member.Gender.Text == "남" ? "장년" : "부녀");
                                                        member.Department.MarkYellow();
                                                    }

                                                    // 青年 → 老年
                                                    if ((member.Department.Text == "청년") && age >= 70)
                                                    {
                                                        member.Department.SetCell("자문");
                                                        member.Department.MarkYellow();
                                                    }
                                                }
                                            }
                                        }

                                        #endregion

                                        #region 掉落的学生整行改变字体颜色

                                        var targetColor = Color.FromArgb(0, 112, 192); // 初级Drop颜色
                                        foreach (var member in onceStudentList)
                                        {
                                            if (dropStudentList.Exists(p => (p.Period.Text == member.Period.Text && p.ChineseName.Text == member.ChineseName.Text)))
                                            {
                                                var properties = member.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
                                                foreach (var property in properties)
                                                {
                                                    if (property.PropertyType == typeof(ExcelRange))
                                                    {
                                                        var excelRange = property.GetValue(member) as ExcelRange;
                                                        excelRange?.Style.Font.Color.SetColor(targetColor);
                                                    }
                                                }
                                            }
                                        }

                                        #endregion

                                        // File.WriteAllText(@"W:\work\kzj\二开韩文版测试\log_grbg.txt", string.Join("\r\n", resultList.Select(p => $"{p.Period} {p.ChineseName.Text}").ToList()));

                                        #region 数据变更

                                        foreach (var member in onceStudentList)
                                        {
                                            var changeMember = changeStudentList
                                            .FirstOrDefault(p => p.Period.Text == member.Period.Text && 
                                                                 p.ChineseName.Text == member.ChineseName.Text && 
                                                                 p.IDCardBirth.Text == member.IDCardBirth.Text);
                                            if (changeMember != null)
                                            {
                                                if (member.ChineseName.Text != changeMember.NewChineseName.Text)
                                                {
                                                    changeMember.NewChineseName.MarkYellow();
                                                }
                                                member.ChineseName.SetCell(changeMember.NewChineseName);

                                                if (member.KoreanName.Text != changeMember.NewKoreanName.Text)
                                                {
                                                    changeMember.NewKoreanName.MarkYellow();
                                                }
                                                member.KoreanName.SetCell(changeMember.NewKoreanName);

                                                if (member.IDCardBirth.Text != changeMember.NewIDCardBirth.Text)
                                                {
                                                    changeMember.NewIDCardBirth.MarkYellow();
                                                }
                                                if (int.TryParse(member.IDCardBirth.Text.Substring(0, 4), out int oriBirthYear))
                                                {
                                                    member.OriAge = DateTime.Now.Year - oriBirthYear; // 记录变更前的年龄
                                                }
                                                member.IDCardBirth.SetCell(changeMember.NewIDCardBirth);

                                                if (member.Phone.Text != changeMember.NewPhone.Text)
                                                {
                                                    changeMember.NewPhone.MarkYellow();
                                                }
                                                member.Phone.SetCell(changeMember.NewPhone);

                                                if (member.Country.Text != changeMember.Country.Text)
                                                {
                                                    changeMember.Country.MarkYellow();
                                                }
                                                member.Country.SetCell(changeMember.Country);

                                                if (member.CurrentAddressCountry.Text != changeMember.CurrentAddressCountry.Text)
                                                {
                                                    changeMember.CurrentAddressCountry.MarkYellow();
                                                }
                                                member.CurrentAddressCountry.SetCell(changeMember.CurrentAddressCountry);

                                                if (member.CurrentAddressCity.Text != changeMember.CurrentAddressCityKorean.Text)
                                                {
                                                    changeMember.CurrentAddressCityKorean.MarkYellow();
                                                }
                                                member.CurrentAddressCity.SetCell(changeMember.CurrentAddressCityKorean);

                                                if (member.CurrentAddress.Text != changeMember.CurrentAddress.Text)
                                                {
                                                    changeMember.CurrentAddress.MarkYellow();
                                                }
                                                member.CurrentAddress.SetCell(changeMember.CurrentAddress);

                                                if (member.IDCardAddress.Text != changeMember.IDCardAddress.Text)
                                                {
                                                    changeMember.IDCardAddress.MarkYellow();
                                                }
                                                member.IDCardAddress.SetCell(changeMember.IDCardAddress);
                                            }
                                        }

                                        #endregion                 

                                        #region 排序整理

                                        var departmentSortDic = new Dictionary<string, int>
                                        {
                                            { "자문", 1 },
                                            { "장년", 2 },
                                            { "부녀", 3 },
                                            { "청년", 4 }
                                        };

                                        const int defaultSortValue = int.MaxValue;

                                        var resultList = onceStudentList
                                        // 按照期数进行排序
                                        .OrderBy(member => member.Period.Text)
                                        // 按照是否掉落进行排序
                                        .ThenBy(member => !string.IsNullOrEmpty(member.DropStep.Text) ? 1 : 0)
                                        // 对掉落的学生进行课程排序
                                        .ThenByDescending(member => 
                                        {
                                            if (!string.IsNullOrEmpty(member.DropStep.Text))
                                            {
                                                var courseNumber = Regex.Match(member.DropStep.Text, @"\d+").Value;
                                                return int.TryParse(courseNumber, out int courseNumberMatchResult) ? courseNumberMatchResult : 0;
                                            }
                                            return 0;
                                        })
                                        // 对未掉落的学生进行老壮妇青排序
                                        .ThenBy(student => string.IsNullOrEmpty(student.DropStep.Text)
                                                           ? departmentSortDic.GetValueOrDefault(student.Department.Text, defaultSortValue)
                                                           : defaultSortValue)
                                        // 对未掉落的学生进行老壮妇青排序后再进行生日排序
                                        .ThenBy(student => string.IsNullOrEmpty(student.DropStep.Text)
                                                           ? student.IDCardBirth.Text
                                                           : "")
                                        .ToList();

                                        #endregion             

                                        #region 重新计算部署后，再按照老壮妇青排序

                                        // var departmentSortDic = new Dictionary<string, int>
                                        // {
                                        //     { "자문", 1 },
                                        //     { "장년", 2 },
                                        //     { "부녀", 3 },
                                        //     { "청년", 4 }
                                        // };
                                        // 
                                        // const int defaultSortValue = int.MaxValue;
                                        // 
                                        // resultList = resultList
                                        //     .OrderBy(student => student.Period.Text)
                                        //     .ThenBy(student =>
                                        //     {
                                        //         return departmentSortDic.ContainsKey(student.Department.Text)
                                        //             ? departmentSortDic[student.Department.Text]
                                        //             : defaultSortValue;
                                        //     })
                                        //     .ThenBy(student => student.IDCardBirth.Text) // 按年龄从大到小排序
                                        //     .ToList();

                                        #endregion

                                        Output(resultList);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        this.gbx_once_file.Visible = true;
                        this.gbx_drop_file.Visible = true;
                        this.gbx_change_file.Visible = true;
                        this.btn_run.Visible = true;
                        this.txt_check_result.Text = $"{ex.Message}\r\n{ex.InnerException}\r\n{ex.Source}\r\n{ex.StackTrace}\r\n{ex.TargetSite}";
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

        public int OriAge { get; set; } // 原始年龄
    }

    public class PeriodOnceStudentList
    {
        public string Period { get; set; }
        public List<OnceStudent> List { get; set; }
    }

    public class DropStudent
    {
        public ExcelRange School { get; set; } // A
        public ExcelRange Period { get; set; } // B
        public ExcelRange KoreanName { get; set; } // C
        public ExcelRange ChineseName { get; set; } // D
        public ExcelRange DropType { get; set; } // E
        public ExcelRange DropStep { get; set; } // F
        public ExcelRange DropReason { get; set; } // G
        public ExcelRange DropManage { get; set; } // H
        public ExcelRange DropOutMoveDetail { get; set; } // I
        public ExcelRange DropReasonDesc { get; set; } // J
    }

    public class ChangeStudent
    {
        public ExcelRange Period { get; set; } // A
        public ExcelRange ChineseName { get; set; } // B
        public ExcelRange NewChineseName { get; set; } // C
        public ExcelRange KoreanName { get; set; } // D
        public ExcelRange NewKoreanName { get; set; } // E
        public ExcelRange IDCardBirth { get; set; } // F
        public ExcelRange NewIDCardBirth { get; set; } // G
        public ExcelRange Phone { get; set; } // H
        public ExcelRange NewPhone { get; set; } // I
        public ExcelRange Country { get; set; } // J
        public ExcelRange CurrentAddressCountry { get; set; } // K
        public ExcelRange CurrentAddressCity { get; set; } // L
        public ExcelRange CurrentAddressCityKorean { get; set; } // M
        public ExcelRange CurrentAddress { get; set; } // N
        public ExcelRange IDCardAddress { get; set; } // O
    }
}
