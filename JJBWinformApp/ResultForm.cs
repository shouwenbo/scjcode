using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JJBWinformApp
{
    public partial class ResultForm : Form
    {
        public ResultForm(string checkResult)
        {
            InitializeComponent();
            this.txt_check_result.Text = string.IsNullOrEmpty(checkResult) ? "表格验证全部合格！" : checkResult;
        }
    }
}
