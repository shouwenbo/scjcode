using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BatchTranslateApp
{
    public partial class Mainform : Form
    {
        public Mainform()
        {
            InitializeComponent();
            this.panel_content.BackColor = System.Drawing.Color.LightBlue;
            this.panel_content.BorderStyle = BorderStyle.FixedSingle;
            this.btn_normal_translate.Click += (s, e) => SelectForm(this.btn_normal_translate, new NormalTranslateForm());
            this.btn_jjb_translate.Click += (s, e) => SelectForm(this.btn_jjb_translate, new JJBTranslateForm());
            this.btn_once_open_translate.Click += (s, e) => SelectForm(this.btn_once_open_translate, new OnceOpenTranslateForm());
            this.btn_twice_open_translate.Click += (s, e) => SelectForm(this.btn_twice_open_translate, new TwiceOpenTranslateForm());
            SelectForm(this.btn_normal_translate, new NormalTranslateForm());
        }

        /// <summary>
        /// 选择窗口
        /// </summary>
        public void SelectForm(Button button, Form form)
        {
            foreach (var tab in this.panel_tab_list.Controls)
            {
                if (tab is Button btn)
                {
                    btn.Enabled = true;
                }
            }
            this.lbl_title.Text = button.Text;
            button.Enabled = false;

            var isExistingForm = false;
            foreach (Control control in this.panel_content.Controls)
            {
                if (control is Form existingForm)
                {
                    if (existingForm.GetType() == form.GetType())
                    {
                        existingForm.BringToFront();
                        existingForm.Visible = true;
                        isExistingForm = true;
                    }
                    else
                    {
                        existingForm.Visible = false;
                    }
                }
            }
            if (!isExistingForm)
            {
                form.FormBorderStyle = FormBorderStyle.None;
                form.TopLevel = false;
                this.panel_content.Controls.Add(form);
                int x = (this.panel_content.Width - form.Width) / 2;
                int y = (this.panel_content.Height - form.Height) / 2;
                form.Location = new Point(x, y);
                form.Show();
            }
        }

        #region 资源释放
        private void Mainform_FormClosing(object sender, FormClosingEventArgs e)
        {
            foreach (Control control in this.panel_content.Controls)
            {
                if (control is Form form)
                {
                    form.Close();
                    form.Dispose();
                    this.panel_content.Controls.Remove(form);
                }
            }
        }
        #endregion
    }
}
