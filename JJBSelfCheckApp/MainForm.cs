using Common;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace JJBSelfCheckApp
{
    public partial class MainForm : Form
    {
        [DllImport("kernel32.dll")]
        public static extern IntPtr _lopen(string lpPathName, int iReadWrite);

        [DllImport("kernel32.dll")]
        public static extern bool CloseHandle(IntPtr hObject);

        public const int OF_READWRITE = 2;
        public const int OF_SHARE_DENY_NONE = 0x40;
        public readonly IntPtr HFILE_ERROR = new IntPtr(-1);

        public List<Member>? MemberList { get; set; }

        public MainForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
        }

        private void LoadTable(string path)
        {
            try
            {
                string selectedFileName = path;
                // this.lbl_file_path.Text = selectedFileName;
                MemberList = new TableHelper(new MessageBoxDialogService()).Load(selectedFileName);
                LoadList();

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
                }
                new ResultForm(string.Join("\r\n", list)).Show();
            }
            catch (Exception ex)
            {
                IntPtr vHandle = _lopen(path, OF_READWRITE | OF_SHARE_DENY_NONE);

                if (vHandle == HFILE_ERROR)
                {
                    MessageBox.Show("文件被占用，请先关闭文件再尝试操作");
                }
                else
                {
                    MessageBox.Show(ex.ToString());
                }

                CloseHandle(vHandle);
            }
        }

        private void LoadList()
        {
            if (MemberList != null && MemberList.Count > 0)
            {
                //this.lbx_members.DataSource = MemberList;
                //this.lbx_members.DisplayMember = "ListName";
                //this.lbx_members.Refresh();
                var member = MemberList[0];
                //SelectMember(member);
            }
            else
            {
                //this.lbx_members.Items.Clear();
                //this.EducationList.Clear();
                //this.FamilyList.Clear();
                //ReloadEducation();
                //ReloadFamily();
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
                // 获取拖入的文件路径
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    // this.lbl_file_path.Text = files[0]; // 只处理第一个文件
                    LoadTable(files[0]);
                }
            }
        }

        private void btn_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // this.lbl_file_path.Text = openFileDialog.FileName;
                LoadTable(openFileDialog.FileName);
            }
        }
    }
}
