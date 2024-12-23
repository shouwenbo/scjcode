using Common;
using System.Reflection;
using System.Text;

namespace JJBWinformApp
{
    public partial class TableCompareForm : Form
    {
        private readonly TableService tableService;

        public List<Member>? OriMemberList { get; set; }
        public List<Member>? NewMemberList { get; set; }

        public TableCompareForm(TableService tableService, string new_path)
        {
            this.tableService = tableService;
            InitializeComponent();
            this.lbl_new_table_path.Text = new_path;
        }

        private void LoadOriTable(string path = "")
        {
            if (path != "")
            {
                this.lbl_ori_table_path.Text = path;
            }
            OriMemberList = tableService.Load(this.lbl_ori_table_path.Text, 4, false);
            if (OriMemberList == null)
            {
                this.lbl_ori_table_path.Text = "";
            }
        }

        private void LoadNewTable(string path = "")
        {
            if (path != "")
            {
                this.lbl_new_table_path.Text = path;
            }
            NewMemberList = tableService.Load(this.lbl_new_table_path.Text, 4, false);
            if (NewMemberList == null)
            {
                this.lbl_new_table_path.Text = "";
            }
        }

        private void btn_start_compare_Click(object sender, EventArgs e)
        {
            LoadOriTable();
            LoadNewTable();
            if (OriMemberList != null && NewMemberList != null)
            {
                var list = new List<string>();
                foreach (var oriMember in OriMemberList)
                {
                    var checkList = oriMember.Check();
                    if (checkList.Count > 0)
                    {
                        var builder = new StringBuilder();
                        builder.AppendLine($"{oriMember.Index} {oriMember.ChineseName}：");
                        var newMember = NewMemberList.SingleOrDefault(p => p.Index == oriMember.Index);
                        Console.WriteLine();
                        //var newMember = NewMemberList.SingleOrDefault(p => p.KoreanName == oriMember.KoreanName了);
                        if (newMember != null)
                        {
                            CompareObjects(oriMember, newMember).ForEach(p => builder.AppendLine(p));
                            foreach (var oriEducation in oriMember.EducationList)
                            {
                                var newEducation = newMember.EducationList.SingleOrDefault(p => p.Index == oriEducation.Index);
                                if (newEducation != null)
                                {
                                    CompareObjects(oriEducation, newEducation).ForEach(p => builder.AppendLine($"学历信息{oriEducation.Index}{p}"));
                                }
                                else
                                {
                                    builder.AppendLine($"学历信息{oriEducation.Index}: 已删除此学历");
                                }
                            }
                            foreach (var oriFamily in oriMember.FamilyMemberList)
                            {
                                var newFamily = newMember.FamilyMemberList.SingleOrDefault(p => p.Index == oriFamily.Index);
                                if (newFamily != null)
                                {
                                    CompareObjects(oriFamily, newFamily).ForEach(p => builder.AppendLine($"家族事项{oriFamily.Index}{p}"));
                                }
                                else
                                {
                                    builder.AppendLine($"家族事项{oriFamily.Index}: 已删除此家族成员");
                                }
                            }
                        }
                        else
                        {
                            builder.AppendLine("已删除此学员");
                        }
                        list.Add(builder.ToString());
                    }
                }
                new ResultForm(string.Join("\r\n", list)).Show();
            }
            else
            {
                MessageBox.Show("无法读取数据");
            }
        }

        private void btn_ori_table_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadOriTable(openFileDialog.FileName);
            }
        }

        private void btn_new_table_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                LoadNewTable(openFileDialog.FileName);
            }
        }

        public List<string> CompareObjects<T>(T obj1, T obj2)
        {
            var result = new List<string>();
            Type type = typeof(T);
            PropertyInfo[] properties = type.GetProperties();

            foreach (PropertyInfo property in properties)
            {
                if (property.CanRead && property.CanWrite && property.PropertyType == typeof(string))
                {
                    string propertyName = property.Name;
                    string propertyComment = GetPropertyComment(property);

                    if (!string.IsNullOrEmpty(propertyComment))
                    {
                        string value1 = (string)property.GetValue(obj1);
                        string value2 = (string)property.GetValue(obj2);

                        if (value1 != value2)
                        {
                            result.Add($"{propertyComment}: 将\"{value1}\" 改为了 \"{value2}\"");
                        }
                    }
                }
            }

            return result;
        }

        private string? GetPropertyComment(PropertyInfo property)
        {
            var attribute = (PropertyNameAttribute)Attribute.GetCustomAttribute(property, typeof(PropertyNameAttribute));

            return attribute?.Name ?? null;
        }
    }
}
