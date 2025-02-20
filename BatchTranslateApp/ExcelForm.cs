using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BatchTranslateApp
{
    public class ExcelForm: Form
    {
        /// <summary>
        /// 初始化Excel文件选择器
        /// </summary>
        public void InitExcelSelector(GroupBox gbx, Button selectFileBtn, Action<string> onFileSelected)
        {
            gbx.AllowDrop = true;
            gbx.DragEnter += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    e.Effect = DragDropEffects.Copy;
                }
                else
                {
                    e.Effect = DragDropEffects.None;
                }
            };
            gbx.DragDrop += (s, e) =>
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = (string[])e.Data.GetData(DataFormats.FileDrop);
                    if (files.Length > 0)
                    {
                        var file = files[0];
                        
                        onFileSelected(file);
                    }
                }
            };
            selectFileBtn.Click += (s, e) =>
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var file = openFileDialog.FileName;
                    onFileSelected(file);
                }
            };
        }

        public void InitExcelSelector(GroupBox gbx, Button selectFileBtn)
        {
            InitExcelSelector(gbx, selectFileBtn, str => { });
        }
    }
}
