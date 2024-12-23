using Common;
using Microsoft.Extensions.Configuration;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace JJBWinformApp
{
    public partial class MainForm : Form
    {
        private readonly TableService tableService;

        public List<Member>? MemberList { get; set; }

        public List<EducationInfo>? EducationList { get; set; }
        public int EducationIndex { get; set; } = -1;
        public List<FamilyMemberInfo>? FamilyList { get; set; }
        public int FamilyIndex { get; set; } = -1;

        public MainForm(TableService tableService)
        {
            this.tableService = tableService;
            InitializeComponent();
            foreach (var control in gpx_basic_info.Controls)
            {
                if (control is Label)
                {
                    var label = control as Label;
                    if (!label.Name.Contains("label"))
                    {
                        label.Text = "";
                        label.ForeColor = Color.Blue;
                    }
                }
            }
            foreach (var control in gpx_education.Controls)
            {
                if (control is Label)
                {
                    var label = control as Label;
                    if (!label.Name.Contains("label"))
                    {
                        label.Text = "";
                        label.ForeColor = Color.Blue;
                    }
                }
            }
            foreach (var control in gpx_family.Controls)
            {
                if (control is Label)
                {
                    var label = control as Label;
                    if (!label.Name.Contains("label"))
                    {
                        label.Text = "";
                        label.ForeColor = Color.Blue;
                    }
                }
            }
            //LoadTable(@"D:\Users\寿文博\Documents\Jesus\新天地\协力 & 教籍簿\jjb 培训（65人）.xlsx");
            //LoadTable(@"D:\Users\寿文博\Documents\Jesus\新天地\协力 & 教籍簿\153-4期\更正后的版本\153-4期 jjb表格.xlsx");
        }

        private void btn_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadTable(openFileDialog.FileName);
            }
        }

        private void LoadTable(string path)
        {
            string selectedFileName = path;
            this.lbl_file_path.Text = selectedFileName;
            MemberList = tableService.Load(this.lbl_file_path.Text, int.Parse(this.txt_config_first_column.Text), true, this.cbx_backcolor.Checked, this.cbx_phone.Checked);
            LoadList();
        }

        private void LoadList()
        {
            if (MemberList != null && MemberList.Count > 0)
            {
                this.lbx_members.DataSource = MemberList;
                this.lbx_members.DisplayMember = "ListName";
                this.lbx_members.Refresh();
                var member = MemberList[0];
                SelectMember(member);
            }
            else
            {
                this.lbx_members.Items.Clear();
                this.EducationList.Clear();
                this.FamilyList.Clear();
                ReloadEducation();
                ReloadFamily();
            }
        }

        private void SelectMember(Member member)
        {
            // this.lbl_column_index.Text = member.ColumnIndex.ToString();
            this.lbl_column_name.Text = member.ColumnName;
            this.lbl_index.Text = member.Index;
            this.lbl_korean_name.Text = member.KoreanName;
            this.lbl_chinese_name.Text = member.ChineseName;
            this.lbl_country.Text = member.Country;
            this.lbl_idcard_birthday.Text = member.IDCardBirthday;
            this.lbl_gender.Text = member.Gender;
            this.lbl_actual_birthday.Text = member.ActualBirthday;
            this.lbl_idcard_address_chinese.Text = member.IDCardAddressChinese;
            this.lbl_idcard_address_korean.Text = member.IDCardAddressKorean;
            this.lbl_current_address_chinese.Text = member.CurrentAddressChinese;
            this.lbl_current_address_korean.Text = member.CurrentAddressKorean;
            this.lbl_phone.Text = member.Phone;
            this.lbl_email.Text = member.Email;
            this.lbl_height.Text = member.Height;
            this.lbl_blood_type.Text = member.BloodType;
            this.lbl_is_married.Text = member.IsMarried;
            this.lbl_cd_type.Text = member.CDType;
            this.lbl_birth_religion.Text = member.BirthReligion;
            this.lbl_other_religion.Text = member.OtherReligion;
            this.lbl_church_name.Text = member.ChurchName;
            this.lbl_position.Text = member.Position;
            this.lbl_years_of_faith.Text = member.YearsOfFaith;
            this.lbl_maternal_faith.Text = member.MaternalFaith;
            this.lbl_final_education.Text = member.FinalEducation;
            this.lbl_sibling_relationship.Text = member.SiblingRelationship;
            this.EducationList = member.EducationList;
            this.FamilyList = member.FamilyMemberList;
            ReloadEducation();
            ReloadFamily();
        }

        private void ReloadEducation()
        {
            this.btn_education_prev.Enabled = false;
            this.btn_education_next.Enabled = false;
            var education = EducationList.FirstOrDefault();
            
            if (education != null)
            {
                EducationIndex = 0;
                SelectEducation();

            }
            else
            {
                EducationIndex = -1;
                SelectEducation();
                this.EducationList.Clear();
            }
        }

        private void SelectEducation()
        {
            if (EducationIndex > -1)
            {
                var education = EducationList[EducationIndex];
                if (EducationIndex == 0)
                {
                    this.btn_education_prev.Enabled = false;
                }
                if (EducationIndex > 0)
                {
                    this.btn_education_prev.Enabled = true;
                }
                if (EducationIndex < EducationList.Count)
                {
                    this.btn_education_next.Enabled = true;
                }
                if ((EducationIndex + 1) >= EducationList.Count)
                {
                    this.btn_education_next.Enabled = false;
                }
                this.lbl_education_category.Text = education.Category;
                this.lbl_education_school_name.Text = education.SchoolName;
                this.lbl_education_major.Text = education.Major;
                this.lbl_education_status.Text = education.Status;
                this.lbl_education_period.Text = education.Period;
            }
            else
            {
                this.btn_education_prev.Enabled = false;
                this.btn_education_next.Enabled = false;
                this.lbl_education_category.Text = "";
                this.lbl_education_school_name.Text = "";
                this.lbl_education_major.Text = "";
                this.lbl_education_status.Text = "";
                this.lbl_education_period.Text = "";
            }
        }

        private void btn_education_prev_Click(object sender, EventArgs e)
        {
            EducationIndex--;
            SelectEducation();
        }

        private void btn_education_next_Click(object sender, EventArgs e)
        {
            EducationIndex++;
            SelectEducation();
        }

        private void ReloadFamily()
        {
            this.btn_family_prev.Enabled = false;
            this.btn_family_next.Enabled = false;
            var family = FamilyList.FirstOrDefault();

            if (family != null)
            {
                FamilyIndex = 0;
                SelectFamily();

            }
            else
            {
                FamilyIndex = -1;
                SelectFamily();
                this.FamilyList.Clear();
            }
        }

        private void SelectFamily()
        {
            if (FamilyIndex > -1)
            {
                var family = FamilyList[FamilyIndex];
                if (FamilyIndex == 0)
                {
                    this.btn_family_prev.Enabled = false;
                }
                if (FamilyIndex > 0)
                {
                    this.btn_family_prev.Enabled = true;
                }
                if (FamilyIndex < FamilyList.Count)
                {
                    this.btn_family_next.Enabled = true;
                }
                if ((FamilyIndex + 1) >= FamilyList.Count)
                {
                    this.btn_family_next.Enabled = false;
                }
                this.lbl_family_name.Text = family.Name;
                this.lbl_family_gender.Text = family.Gender;
                this.lbl_family_birthday.Text = family.Birthday;
                this.lbl_family_relationship.Text = family.Relationship;
                this.lbl_family_religion.Text = family.Religion;
                this.lbl_family_other_religion.Text = family.OtherReligion;
                this.lbl_family_birth_church.Text = family.BirthChurch;
                this.lbl_family_scj_number.Text = family.SCJNumber;
            }
            else
            {
                this.btn_family_prev.Enabled = false;
                this.btn_family_next.Enabled = false;
                this.lbl_family_name.Text = "";
                this.lbl_family_gender.Text = "";
                this.lbl_family_birthday.Text = "";
                this.lbl_family_relationship.Text = "";
                this.lbl_family_religion.Text = "";
                this.lbl_family_other_religion.Text = "";
                this.lbl_family_birth_church.Text = "";
                this.lbl_family_scj_number.Text = "";
            }
        }

        private void btn_family_prev_Click(object sender, EventArgs e)
        {
            FamilyIndex--;
            SelectFamily();
        }

        private void btn_family_next_Click(object sender, EventArgs e)
        {
            FamilyIndex++;
            SelectFamily();
        }

        private void lbx_members_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
                return;

            var member = MemberList[e.Index];

            Color textColor = member.ValidateError ? Color.Red : e.ForeColor;

            e.DrawBackground();

            // 创建用于绘制文本的Brush
            using (Brush brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(member.ListName, e.Font, brush, e.Bounds, StringFormat.GenericDefault);
            }
            e.DrawFocusRectangle();
        }

        private void lbx_members_SelectedIndexChanged(object sender, EventArgs e)
        {
            var member = MemberList[lbx_members.SelectedIndex];
            SelectMember(member);
        }

        /// <summary>
        /// 检查
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btn_check_all_Click(object sender, EventArgs e)
        {
            var gatherPath = File.ReadAllText("GatherPath.txt");
            if (string.IsNullOrEmpty(gatherPath))
            {
                MessageBox.Show("请先配置汇总表格");
                return;
            }
            var gatherList = tableService.LoadGatherList(gatherPath);

            if (MemberList == null || MemberList.Count == 0)
            {
                return;
            }
            var list = new List<string>();
            foreach (var member in MemberList)
            {
                var checkList = member.Check();
                var builder = new StringBuilder();
                if (checkList.Count > 0)
                {
                    builder.AppendLine($"{member.ListName.Trim()}：");
                    for (int i = 0; i < checkList.Count; i++)
                    {
                        builder.AppendLine($"{i + 1}：{checkList[i]}");
                    }
                    list.Add(builder.ToString());
                }

                #region 二次开讲/汇总表 信息核对

                var gatherMemberArray = gatherList.Where(p => p.ChineseName.Trim() == member.ChineseName.Trim()).ToList();
                if (gatherMemberArray.Count != 1)
                {
                    builder.AppendLine($"汇总表匹配到了多条数据");
                    MessageBox.Show($"{member.ChineseName} 在汇总表匹配到了 {gatherMemberArray.Count} 条数据");
                    return;
                }
                var gatherMember = gatherMemberArray[0];
                var gatherIdCardBirthDay = gatherMember.IDCardBirthday.Replace("-", "").Replace("(+)", "");
                var gatherActualBirthDay = gatherMember.ActualBirthday.Replace("-", "").Replace("(+)", "");
                if (member.IDCardBirthday != gatherIdCardBirthDay)
                {
                    builder.AppendLine($"身份证生日: {member.IDCardBirthday} 与汇总表的身份证生日‘{gatherMember.IDCardBirthday}’不一致");
                }
                if (member.ActualBirthday != gatherMember.ActualBirthday)
                {
                    builder.AppendLine($"实际生日: {member.ActualBirthday} 与汇总表的实际生日‘{gatherMember.ActualBirthday}’不一致");
                }

                #endregion
            }
            new ResultForm(string.Join("\r\n", list)).Show();
        }

        private void btn_create_bak_file_Click(object sender, EventArgs e)
        {
            if (MemberList == null || MemberList.Count == 0)
            {
                return;
            }
            var list = new List<string>();
            File.WriteAllText("更正备注.txt", string.Join("\r\n\r\n", MemberList.Select(p => $"{p.Index} {p.ChineseName} ({p.ColumnName}列):")));
        }

        private void btn_open_compare_form_Click(object sender, EventArgs e)
        {
            new TableCompareForm(tableService, this.lbl_file_path.Text).Show();
        }

        private void btn_select_2_kj_jjb_hz_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                File.WriteAllText("GatherPath.txt", openFileDialog.FileName);
                MessageBox.Show("配置成功");
            }
        }

        private void btn_create_hzgl_table_Click(object sender, EventArgs e)
        {
            if (MemberList == null || MemberList.Count == 0)
            {
                return;
            }

            var path = File.ReadAllText("GatherPath.txt");
            if (string.IsNullOrEmpty(path))
            {
                MessageBox.Show("请先配置汇总表格");
                return;
            }

            var resultList = new List<Member>();
            var gatherList = tableService.LoadGatherList(path);

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("Sheet1");

            for (int i = 0; i < MemberList.Count; i++)
            {
                var chineseName = MemberList[i].ChineseName;
                var gatherMemberArray = gatherList.Where(p => p.ChineseName == chineseName).ToList();
                if (gatherMemberArray.Count != 1)
                {
                    MessageBox.Show($"姓名:{chineseName} 出现重名");
                    return;
                }
                var gatherMember = gatherMemberArray[0];
                resultList.Add(new Member()
                {
                    ChineseName = MemberList[i].ChineseName,
                    KoreanName = gatherMember.KoreanName,
                    ActualBirthday = gatherMember.ActualBirthday,
                    IDCardBirthday = gatherMember.IDCardBirthday,
                    Gender = gatherMember.Gender
                });
            }

            IRow row1 = sheet.CreateRow(0);
            IRow row2 = sheet.CreateRow(1);
            IRow row3 = sheet.CreateRow(2);
            IRow row4 = sheet.CreateRow(3);
            IRow row5 = sheet.CreateRow(4);
            for (int i = 0; i < resultList.Count; i++)
            {
                row1.CreateCell(i).SetCellValue(resultList[i].ChineseName);
                row2.CreateCell(i).SetCellValue(resultList[i].KoreanName);
                row3.CreateCell(i).SetCellValue(resultList[i].ActualBirthday);
                row4.CreateCell(i).SetCellValue(resultList[i].IDCardBirthday);
                row5.CreateCell(i).SetCellValue(resultList[i].Gender);
            }

            // 保存文件
            string filePath = $"{DateTime.Now:yyyyMMdd} 汇总表格 {DateTime.Now:HHmmss}.xlsx";
            using (FileStream stream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }

            MessageBox.Show("表格已生成完毕");
        }
    }
}