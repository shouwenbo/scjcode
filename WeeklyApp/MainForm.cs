using Common;
using ICSharpCode.SharpZipLib.Zip;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.Util;
using NPOI.XSSF.UserModel;
using Org.BouncyCastle.Bcpg.Sig;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.IO;
using System.Reflection.Metadata;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace WeeklyApp
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

        public MainForm()
        {
            InitializeComponent();

            var prevSunday = DateTime.Now.GetSundayOfPreviousWeek();
            lbl_hwzb_file_path.Text = $@"请选择上周提交的韩文周报“중국무한-신학부주간보고서-신{prevSunday.Year - 1983}({prevSunday.Year}).{prevSunday.Month}.{prevSunday.Day}..xlsx”";

            // 测试

            lbl_zbjl_file_path.Text = @"C:\Users\寿文博\Desktop\周报助手\周报接龙.xlsx";
            lbl_zbjl_file_path.ForeColor = Color.Blue;
            lbl_dlbg_file_path.Text = @"C:\Users\寿文博\Desktop\周报助手\掉落报告.xlsx";
            lbl_dlbg_file_path.ForeColor = Color.Blue;
            lbl_hwzb_file_path.Text = @"D:\Users\寿文博\Documents\Jesus\釜山雅各 & 学院 & 周报\周报学习\41.5.31周报\중국무한-신학부주간보고서-신41(2024).6.2..xlsx";
        }

        public void 一键生成()
        {

            var basePath = AppDomain.CurrentDomain.BaseDirectory;
            string tablePath = $@"{basePath}\template\掉落报告模板.xlsx";

            var sunday = DateTime.Now.GetSundayOfCurrentWeek();
            var month = sunday.Month;
            var day = sunday.Day;

            string outputDir = $"output/掉落报告/";

            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            var outputPath = $@"{basePath}\{outputDir}\{month}.{day}掉落报告.xlsx";

            try
            {
                using FileStream fileStream = new FileStream(tablePath, FileMode.Open, FileAccess.Read);
                IWorkbook? workbook = WorkbookFactory.Create(fileStream, "wh12000", true);
                if (workbook == null)
                {
                    MessageBox.Show("无法读取表格文件");
                    return;
                }

                try
                {
                    ISheet 各期数汇总表 = workbook.GetSheetAt(0);
                    //各期数汇总表.ShiftRows(3, 3, 3, true, false);
                    var 各期数汇总表_参照行 = 各期数汇总表.GetRow(2);
                    var 各期数汇总表_参照行_样式 = 各期数汇总表_参照行.RowStyle;
                    for (int i = 3; i < 20; i++)
                    {
                        var 各期数汇总表_新增一行 = 各期数汇总表.CreateRow(i);
                        if (各期数汇总表_参照行_样式 != null)
                        {
                            各期数汇总表_新增一行.RowStyle = 各期数汇总表_参照行_样式;
                        }
                        各期数汇总表_新增一行.Height = 各期数汇总表_参照行.Height;

                        for (int col = 0; col < 各期数汇总表_参照行.LastCellNum; col++)
                        {
                            var cellsource = 各期数汇总表_参照行.GetCell(col);
                            var cellInsert = 各期数汇总表_新增一行.CreateCell(col);
                            var cellStyle = cellsource.CellStyle;
                            if (cellStyle != null)
                            {
                                cellInsert.CellStyle = cellsource.CellStyle;
                            }
                        }
                    }

                    //int lastRowNum = sheet.LastRowNum;
                    //IRow newRow = sheet.CreateRow(lastRowNum + 1);
                    //ICell newCell = newRow.CreateCell(0);
                    //newCell.SetCellValue("新值");

                    using FileStream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                    workbook.Write(outputStream);
                    outputStream.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return;
                }
                finally
                {
                    workbook.Close();
                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {
                IntPtr vHandle = _lopen(tablePath, OF_READWRITE | OF_SHARE_DENY_NONE);

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

        /*
         
        {{期数}} 4
        {{姓名}} 1
        !!{{韩文姓名}} 2
        {{电话}} 21
        {{掉落课数}} 18
        {{掉落日期}} 20
        !!{{期数}} 19
        {{商谈者}} 
        {{商谈次数}}
        {{掉落事由}} 26
        {{传道师}} 25
        {{传道师事由}}
        {{讲师}} 24/23
        {{讲师事由}} 
        {{部长}} 费增明/159-7200-5442
        {{新天纪年}}
        {{新天纪月}}
        {{新天纪日}}
         
         */

        #region 选择文件

        private void btn_zbjl_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lbl_zbjl_file_path.Text = openFileDialog.FileName;
                lbl_zbjl_file_path.ForeColor = Color.Blue;
            }
        }

        private void btn_dlbg_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lbl_dlbg_file_path.Text = openFileDialog.FileName;
                lbl_dlbg_file_path.ForeColor = Color.Blue;
            }
        }

        private void btn_hwzb_select_file_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                lbl_hwzb_file_path.Text = openFileDialog.FileName;
                lbl_hwzb_file_path.ForeColor = Color.Blue;
            }
        }

        #endregion

        #region 旧版 <41060103> 生成方式

        [Obsolete]
        private void btn_output_Click(object sender, EventArgs e)
        {
            string clipboardText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(clipboardText))
            {
                new FallManager_410920().ParseData(clipboardText);
            }
            else
            {
                MessageBox.Show("剪贴板是空的！");
            }
        }

        private void btn_yjsc_Click(object sender, EventArgs e)
        {

        }

        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            string clipboardText = Clipboard.GetText();
            if (!string.IsNullOrEmpty(clipboardText))
            {
                new FallManager_410920().ParseData(clipboardText);
            }
            else
            {
                MessageBox.Show("剪贴板是空的！");
            }
        }
    }
}


/*
{{期数}}
{{姓名}}
{{电话}}
{{掉落课数}}
{{掉落日期}}
{{商谈者}}
{{商谈次数}}
{{详细事由}}
{{传道师}}
{{传道师事由}}
{{讲师}}
{{讲师事由}}
{{部长}}
{{新天纪年}}
{{新天纪月}}
{{新天纪日}}
*/