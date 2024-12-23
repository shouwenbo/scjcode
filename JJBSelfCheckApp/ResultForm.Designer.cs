namespace JJBSelfCheckApp
{
    partial class ResultForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ResultForm));
            txt_check_result = new TextBox();
            label1 = new Label();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(15, 16);
            txt_check_result.Margin = new Padding(4, 4, 4, 4);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(987, 549);
            txt_check_result.TabIndex = 1;
            // 
            // label1
            // 
            label1.ForeColor = Color.Red;
            label1.Location = new Point(17, 579);
            label1.Margin = new Padding(4, 0, 4, 0);
            label1.Name = "label1";
            label1.Size = new Size(986, 75);
            label1.TabIndex = 2;
            label1.Text = "温馨提示：本工具无法100%排查错误，仅仅起到辅助作用，所以依然需要人工检查，如果发现有错误工具没检查出来可以反馈，下个版本会进行优化";
            // 
            // ResultForm
            // 
            AutoScaleDimensions = new SizeF(14F, 31F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ClientSize = new Size(1018, 654);
            Controls.Add(label1);
            Controls.Add(txt_check_result);
            Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4, 4, 4, 4);
            MaximizeBox = false;
            Name = "ResultForm";
            Text = "检查结果";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txt_check_result;
        private Label label1;
    }
}