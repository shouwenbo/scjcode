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
    public partial class JJBTranslateForm : ExcelForm
    {
        private int allCount = 0;
        private int currentCount = 0;

        public JJBTranslateForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
            InitExcelSelector(this.gbx_file, this.btn_select_file, file =>
            {
                var list = new JJBTableHelper(new MessageBoxDialogService()).Load(file);
                var periodMatch = Regex.Match(new FileInfo(file).Name, @"\b\d{3}(-\d+)?\b");
                var sheetName = periodMatch.Success ? periodMatch.Value : "请输入期数";
                if (list != null && list.Count > 0)
                {
                    JJBModeTranslate(list, sheetName);
                }
                else
                {
                    this.txt_check_result.Text = "读取不到表格";
                }
            });
            CheckForIllegalCrossThreadCalls = false; // 关闭跨线程调用检查
        }

        #region JJB翻译

        public void JJBModeOutput(List<JJBMember> list, string sheetName)
        {
            using (FileStream file = new FileStream(@"template/JJB韩文模板.xlsx", FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(file);

                // 获取第一个工作表
                ISheet sheet = workbook.GetSheetAt(0);
                workbook.SetSheetName(0, sheetName);

                for (int i = 0; i < list.Count; i++)
                {
                    var member = list[i];
                    sheet.GetRow(0).GetCell(i + 2).SetCellValueIfNotEmpty(member.Index);
                    sheet.GetRow(1).GetCell(i + 2).SetCellValueIfNotEmpty(member.KoreanName);
                    sheet.GetRow(3).GetCell(i + 2).SetCellValueIfNotEmpty(member.ChineseName);
                    sheet.GetRow(4).GetCell(i + 2).SetCellValueIfNotEmpty(member.Country);
                    sheet.GetRow(5).GetCell(i + 2).SetCellValueIfNotEmpty(member.IDCardBirthday);
                    sheet.GetRow(6).GetCell(i + 2).SetCellValueIfNotEmpty(member.Gender);
                    sheet.GetRow(7).GetCell(i + 2).SetCellValueIfNotEmpty(member.CurrentAddressChinese);
                    // sheet.GetRow(8).GetCell(i + 2).SetCellValue(member.IDCardAddressChinese);
                    sheet.GetRow(8).GetCell(i + 2).SetCellValueIfNotEmpty(member.Phone);
                    sheet.GetRow(9).GetCell(i + 2).SetCellValueIfNotEmpty(member.Height);
                    sheet.GetRow(10).GetCell(i + 2).SetCellValueIfNotEmpty(member.BloodType);
                    sheet.GetRow(11).GetCell(i + 2).SetCellValueIfNotEmpty(member.IsMarried);

                    sheet.GetRow(12).GetCell(i + 2).SetCellValueIfNotEmpty(member.CDType);

                    IDrawing drawing = sheet.CreateDrawingPatriarch();
                    IClientAnchor anchor = drawing.CreateAnchor(0, 0, 0, 0, i + 2, 11, i + 2 + 1, 11 + 1);
                    IComment comment = drawing.CreateCellComment(anchor);
                    comment.String = new XSSFRichTextString(member.CDTypeCommentString);
                    sheet.GetRow(12).GetCell(i + 2).CellComment = comment;
                    // sheet.GetRow(12).GetCell(i + 2).CellComment = member.CDTypeComment;

                    sheet.GetRow(13).GetCell(i + 2).SetCellValueIfNotEmpty(member.BirthReligion);
                    sheet.GetRow(14).GetCell(i + 2).SetCellValueIfNotEmpty(member.OtherReligion);
                    sheet.GetRow(15).GetCell(i + 2).SetCellValueIfNotEmpty(member.ChurchName);
                    sheet.GetRow(17).GetCell(i + 2).SetCellValueIfNotEmpty(member.YearsOfFaith);
                    sheet.GetRow(18).GetCell(i + 2).SetCellValueIfNotEmpty(member.MaternalFaith);
                    sheet.GetRow(19).GetCell(i + 2).SetCellValueIfNotEmpty(member.Email);
                    sheet.GetRow(21).GetCell(i + 2).SetCellValueIfNotEmpty(member.FinalEducation);
                    if (member.EducationList.Count > 0)
                    {
                        sheet.GetRow(22).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[0].Category);
                        sheet.GetRow(23).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[0].SchoolName);
                        sheet.GetRow(24).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[0].Major);
                        sheet.GetRow(25).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[0].Status);
                        sheet.GetRow(26).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[0].Period);
                    }
                    if (member.EducationList.Count > 1)
                    {
                        sheet.GetRow(27).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[1].Category);
                        sheet.GetRow(28).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[1].SchoolName);
                        sheet.GetRow(29).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[1].Major);
                        sheet.GetRow(30).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[1].Status);
                        sheet.GetRow(31).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[1].Period);
                        Console.WriteLine();
                    }
                    if (member.EducationList.Count > 2)
                    {
                        sheet.GetRow(32).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[2].Category);
                        sheet.GetRow(33).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[2].SchoolName);
                        sheet.GetRow(34).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[2].Major);
                        sheet.GetRow(35).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[2].Status);
                        sheet.GetRow(36).GetCell(i + 2).SetCellValueIfNotEmpty(member.EducationList[2].Period);
                    }
                    sheet.GetRow(46).GetCell(i + 2).SetCellValueIfNotEmpty(member.SiblingRelationship);
                    if (member.FamilyMemberList.Count > 0)
                    {
                        sheet.GetRow(48).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].Name);
                        sheet.GetRow(49).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].Gender);
                        sheet.GetRow(50).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].Birthday);
                        sheet.GetRow(51).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].Relationship);
                        sheet.GetRow(52).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].BirthChurch);
                        sheet.GetRow(53).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[0].SCJNumber);
                    }
                    if (member.FamilyMemberList.Count > 1)
                    {
                        sheet.GetRow(54).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].Name);
                        sheet.GetRow(55).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].Gender);
                        sheet.GetRow(56).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].Birthday);
                        sheet.GetRow(57).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].Relationship);
                        sheet.GetRow(58).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].BirthChurch);
                        sheet.GetRow(59).GetCell(i + 2).SetCellValueIfNotEmpty(member.FamilyMemberList[1].SCJNumber);
                    }
                }

                int maxEducationCount = list.Max(member => member.EducationList?.Count ?? 0);
                if (maxEducationCount <= 1)
                {
                    sheet.DeleteRowsAndShiftUp(new int[] { 27, 28, 29, 30, 31, 32, 33, 34, 35, 36 });
                }
                if (maxEducationCount == 2)
                {
                    sheet.DeleteRowsAndShiftUp(new int[] { 27, 28, 29, 30, 31 });
                }

                //sheet.GetRow(4).GetCell(7).SetCellValue("一二三四五");
                //sheet.GetRow(5).GetCell(7).SetCellValue("烦烦烦方法");
                //sheet.GetRow(6).GetCell(7).SetCellValue("日日日日日");
                //sheet.GetRow(7).GetCell(7).SetCellValue("他吞吞吐吐");
                //sheet.GetRow(8).GetCell(7).SetCellValue("哈哈哈哈哈");
                //sheet.GetRow(9).GetCell(7).SetCellValue("范德萨富士");
                //sheet.GetRow(10).GetCell(7).SetCellValue("犯得上反对");

                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel 工作簿 (*.xlsx)|*.xlsx";
                    saveFileDialog.Title = "保存 Excel 文件";
                    saveFileDialog.FileName = sheetName == "请输入期数" ? "翻译结果.xlsx" : $"{sheetName} jjb（{list.Count}人） 韩";

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
        }

        public void JJBModeTranslate(List<JJBMember> list, string sheetName)
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
                            #region 基本信息翻译
                            if (!string.IsNullOrEmpty(member.CurrentAddressChinese))
                            {
                                member.CurrentAddressChinese = TranslateHelper.SentenceToKorean(member.CurrentAddressChinese);
                            }
                            if (!string.IsNullOrEmpty(member.Country))
                            {
                                member.Country = TranslateHelper.SentenceToKorean(member.Country);
                            }
                            if (!string.IsNullOrEmpty(member.CDType))
                            {
                                member.CDType = member.CDType
                                      .Replace("家族/亲戚", "가족/친척")
                                      .Replace("知人（朋友/前后辈/同事）", "친구/선후배/직장동료")
                                      .Replace("路旁", "노방")
                                      .Replace("网络/通信", "인터넷/통신")
                                      .Replace("田地", "추수밭");
                            }
                            if (!string.IsNullOrEmpty(member.CDTypeCommentString))
                            {
                                member.CDTypeCommentString = member.CDTypeCommentString
                                      .Replace("家族", "가족")
                                      .Replace("亲戚", "친척")
                                      .Replace("朋友", "친구")
                                      .Replace("前后辈", "선후배")
                                      .Replace("同事", "직장동료")
                                      .Replace("网络", "인터넷")
                                      .Replace("通信", "통신");
                            }
                            if (!string.IsNullOrEmpty(member.BirthReligion))
                            {
                                member.BirthReligion = member.BirthReligion
                                      .Replace("长老教", "장로교")
                                      .Replace("监理教", "감리교")
                                      .Replace("圣洁教", "성결교")
                                      .Replace("浸礼桥", "침례교")
                                      .Replace("纯福音", "순복음")
                                      .Replace("天主教", "천주교")
                                      .Replace("佛教", "불교")
                                      .Replace("无教", "무교")
                                      .Replace("scj(宣教教会)", "신천지(선교교회)")
                                      .Replace("摩门教", "몰몬교")
                                      .Replace("伊斯兰教", "이슬람교")
                                      .Replace("其他宗教", "기타종교")
                                      .Replace("三自教", "삼자교");
                            }
                            if (!string.IsNullOrEmpty(member.OtherReligion))
                            {
                                member.OtherReligion = TranslateHelper.SentenceToKorean(member.OtherReligion.Replace("三自教", "삼자교"));
                            }
                            if (!string.IsNullOrEmpty(member.ChurchName))
                            {
                                member.ChurchName = TranslateHelper.SentenceToKorean(member.ChurchName);
                            }
                            if (!string.IsNullOrEmpty(member.YearsOfFaith))
                            {
                                member.YearsOfFaith = member.YearsOfFaith.Replace("年", "년");
                            }
                            if (!string.IsNullOrEmpty(member.FinalEducation))
                            {
                                member.FinalEducation = member.FinalEducation
                                      .Replace("大学在读", "대학교재학")
                                      .Replace("大学毕业", "대학교졸업")
                                      .Replace("大学辍学", "대학교중퇴")
                                      .Replace("研究生以上", "대학원이상")
                                      .Replace("神学院在读", "신학교재학")
                                      .Replace("神学院毕业", "신학교졸업")
                                      .Replace("神学院辍学", "신학교중퇴")
                                      .Replace("专科在读", "전문대재학")
                                      .Replace("专科毕业", "전문대졸업")
                                      .Replace("专科辍学", "전문대중퇴")
                                      .Replace("神学研究生在读", "신학대학원재학")
                                      .Replace("神学研究生毕业", "신학대학원졸업")
                                      .Replace("神学研究生辍学", "신학대학원중퇴");
                            }
                            #endregion
                            #region 学历翻译
                            foreach (var education in member.EducationList)
                            {
                                if (!string.IsNullOrEmpty(education.Category))
                                {
                                    education.Category = education.Category
                                             .Replace("大学", "대학교")
                                             .Replace("专科", "전문대")
                                             .Replace("研究生", "대학원");
                                }
                                if (!string.IsNullOrEmpty(education.SchoolName))
                                {
                                    education.SchoolName = TranslateHelper.SentenceToKorean(education.SchoolName);
                                }
                                if (!string.IsNullOrEmpty(education.Major))
                                {
                                    education.Major = TranslateHelper.SentenceToKorean(education.Major);
                                }
                                if (!string.IsNullOrEmpty(education.Status))
                                {
                                    education.Status = education.Status
                                             .Replace("毕业", "졸업")
                                             .Replace("辍学", "중퇴")
                                             .Replace("在读", "재학");
                                }
                            }
                            #endregion
                            #region 家属翻译
                            foreach (var family in member.FamilyMemberList)
                            {
                                if (!string.IsNullOrEmpty(family.Name))
                                {
                                    family.Name = TranslateHelper.SentenceToKorean(family.Name);
                                }
                                if (!string.IsNullOrEmpty(family.Relationship))
                                {
                                    family.Relationship = family.Relationship
                                          .Replace("父", "부")
                                          .Replace("母", "모")
                                          .Replace("丈夫", "남편")
                                          .Replace("夫人", "부인")
                                          .Replace("儿子", "아들")
                                          .Replace("女儿", "딸")
                                          .Replace("哥哥", member.Gender == "남" ? "형" : "오빠")
                                          .Replace("弟弟", "동생")
                                          .Replace("妹妹", "동생")
                                          .Replace("姐姐", member.Gender == "남" ? "누나" : "언니")
                                          .Replace("岳父", "장인")
                                          .Replace("岳母", "장모")
                                          .Replace("祖父", "조부")
                                          .Replace("祖母", "조모")
                                          .Replace("婆婆", "시어머니")
                                          .Replace("公公", "시아버지")
                                          .Replace("女婿", "사위")
                                          .Replace("儿媳", "며느리")
                                          .Replace("孙子", "손자");
                                }
                                if (!string.IsNullOrEmpty(family.BirthChurch))
                                {
                                    family.BirthChurch = TranslateHelper.SentenceToKorean(family.BirthChurch);
                                }
                                if (!string.IsNullOrEmpty(family.SCJNumber))
                                {
                                    family.SCJNumber = family.SCJNumber
                                          .Replace("期", "기");
                                }
                            }
                            #endregion
                            currentCount++;
                            this.lbl_task.Text = $"{currentCount}/{allCount}";
                        }
                        this.gbx_file.Visible = true;
                        this.txt_check_result.Text = "翻译完成";
                        JJBModeOutput(list, sheetName);
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

        #endregion
    }
}
