using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection.Metadata;
using System.Text;

namespace BatchPhotoRenameApp
{
    public partial class MainForm : Form
    {
        public string[] files;

        public MainForm()
        {
            InitializeComponent();
            this.gbx_file.AllowDrop = true;
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
                files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    this.lbl_file_count.Text = $"已载入{files.Length}个文件";
                    this.btn_batch_rename.Enabled = true;
                }
            }
        }

        private void btn_select_dir_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "所有文件 (*.*)|*.*";
            openFileDialog.Multiselect = true;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                files = openFileDialog.FileNames;
                if (files.Length > 0)
                {
                    this.lbl_file_count.Text = $"已载入{files.Length}个文件";
                    this.btn_batch_rename.Enabled = true;
                }
            }
        }
        private void ParseData(string text)
        {
            string[] lines = text.Split('\n');

            var list = new List<Student>();

            foreach (string line in lines)
            {
                if (string.IsNullOrWhiteSpace(line)) continue;

                string[] cells = line.Split('\t');

                if (cells.Length == 2)
                {
                    list.Add(new Student()
                    {
                        Name2 = cells[0].Trim(),
                        Name1 = cells[1].Trim()
                    });
                }
            }

            if (list.Count > 0)
            {
                using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
                {
                    folderBrowserDialog.Description = "请选择一个文件夹";
                    folderBrowserDialog.ShowNewFolderButton = true;
                    string defaultFolderPath = AppDomain.CurrentDomain.BaseDirectory;
                    folderBrowserDialog.SelectedPath = defaultFolderPath;

                    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                    {
                        string selectedPath = folderBrowserDialog.SelectedPath;
                        var matchOk = true;
                        var errBuilder = new StringBuilder();

                        var fileNameDic = new Dictionary<string, string>();

                        foreach (var file in files)
                        {
                            var fileInfo = new FileInfo(file);
                            string fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                            var studentList1 = list.Where(p => p.Name1 == fileName).ToList();
                            var studentList2 = list.Where(p => p.Name2 == fileName).ToList();
                            if (studentList1.Count == 1)
                            {
                                fileNameDic[fileName] = studentList1[0].Name2;
                            }
                            else if (studentList2.Count == 1)
                            {
                                fileNameDic[fileName] = studentList2[0].Name1;
                            }
                            else
                            {
                                errBuilder.AppendLine($"复制的内容中没有‘{fileName}’或存在多个");
                                matchOk = false;
                            }
                        }

                        if (matchOk)
                        {
                            foreach (var file in files)
                            {
                                var fileInfo = new FileInfo(file);
                                string fileName = Path.GetFileNameWithoutExtension(fileInfo.Name);
                                var toName = fileNameDic[fileName];
                                string fileExtension = fileInfo.Extension;
                                string targetFilePath = Path.Combine(selectedPath, toName + fileExtension);
                                File.Copy(fileInfo.FullName, targetFilePath, true);
                            }

                            txt_check_result.Text = "导出成功";

                            Process.Start("explorer.exe", selectedPath);
                        }
                        else
                        {
                            txt_check_result.Text = errBuilder.ToString();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("复制内容不符合规范！");
            }
        }

        private void btn_batch_rename_Click(object sender, EventArgs e)
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
    }
}
