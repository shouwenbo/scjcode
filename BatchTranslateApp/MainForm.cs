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
    public partial class MainForm : Form
    {
        private int allCount = 0;
        private int currentCount = 0;

        public MainForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
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

        #region 批量翻译

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

        #endregion

        #region

        public void BatchModeCheck(List<List<string>> list)
        {
            var errList = new List<string>();
            foreach (var chineseList in list)
            {
                foreach (string chinese in chineseList)
                {
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

        #endregion

        #region JJB翻译

        public void JJBModeOutput(List<Member> list, string sheetName)
        {
            using (FileStream file = new FileStream(@"template/韩文模板.xlsx", FileMode.Open, FileAccess.Read))
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
        }

        public void JJBModeTranslate(List<Member> list, string sheetName)
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

        #region 读取Excel

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

        #endregion

        #region

        private void btn_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var file = openFileDialog.FileName;
                var list = LoadExcel(file);
                if (list != null && list.Count > 0)
                {
                    BatchModeCheck(list);
                }
                else
                {
                    this.txt_check_result.Text = "读取不到表格";
                }
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
                    if (radio_batch_mode.Checked)
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
                    }
                    else
                    {
                        var list = new TableHelper(new MessageBoxDialogService()).Load(file);
                        var periodMatch = Regex.Match(new FileInfo(file).Name, @"\d{3}");
                        var sheetName = periodMatch.Success ? periodMatch.Value : "请输入期数";
                        if (list != null && list.Count > 0)
                        {
                            JJBModeTranslate(list, sheetName);
                        }
                        else
                        {
                            this.txt_check_result.Text = "读取不到表格";
                        }
                    }
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

        #endregion

        private void lbl_task_Click(object sender, EventArgs e)
        {

        }

        private void lbl_version_Click(object sender, EventArgs e)
        {

        }

        private void lbl_remark_Click(object sender, EventArgs e)
        {

        }
    }
}
