namespace BatchTranslateApp
{
    partial class OnceOpenTranslateForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(OnceOpenTranslateForm));
            txt_check_result = new TextBox();
            lbl_task = new Label();
            lbl_task_title = new Label();
            lbl_remark = new Label();
            gbx_file = new GroupBox();
            btn_select_file = new Button();
            lbl_file_path = new Label();
            lbl_cdb_password = new Label();
            txt_cdb_password = new TextBox();
            gbx_file.SuspendLayout();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(409, 152);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(294, 133);
            txt_check_result.TabIndex = 6;
            // 
            // lbl_task
            // 
            lbl_task.AutoSize = true;
            lbl_task.Location = new Point(246, 222);
            lbl_task.Margin = new Padding(4, 0, 4, 0);
            lbl_task.Name = "lbl_task";
            lbl_task.Size = new Size(56, 17);
            lbl_task.TabIndex = 9;
            lbl_task.Text = "暂未开始";
            // 
            // lbl_task_title
            // 
            lbl_task_title.AutoSize = true;
            lbl_task_title.Location = new Point(118, 222);
            lbl_task_title.Margin = new Padding(4, 0, 4, 0);
            lbl_task_title.Name = "lbl_task_title";
            lbl_task_title.Size = new Size(68, 17);
            lbl_task_title.TabIndex = 8;
            lbl_task_title.Text = "翻译进度：";
            // 
            // lbl_remark
            // 
            lbl_remark.AutoSize = true;
            lbl_remark.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_remark.Location = new Point(13, 268);
            lbl_remark.Margin = new Padding(2, 0, 2, 0);
            lbl_remark.Name = "lbl_remark";
            lbl_remark.Size = new Size(193, 17);
            lbl_remark.TabIndex = 10;
            lbl_remark.Text = "翻译来源：hanwenxingming.com";
            // 
            // gbx_file
            // 
            gbx_file.Controls.Add(btn_select_file);
            gbx_file.Controls.Add(lbl_file_path);
            gbx_file.Location = new Point(13, 13);
            gbx_file.Margin = new Padding(4);
            gbx_file.Name = "gbx_file";
            gbx_file.Padding = new Padding(4);
            gbx_file.Size = new Size(692, 132);
            gbx_file.TabIndex = 12;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择表格（cdb周报）";
            // 
            // btn_select_file
            // 
            btn_select_file.Location = new Point(216, 50);
            btn_select_file.Name = "btn_select_file";
            btn_select_file.Size = new Size(110, 34);
            btn_select_file.TabIndex = 4;
            btn_select_file.Text = "点此选择";
            btn_select_file.UseVisualStyleBackColor = true;
            // 
            // lbl_file_path
            // 
            lbl_file_path.AutoSize = true;
            lbl_file_path.ImeMode = ImeMode.NoControl;
            lbl_file_path.Location = new Point(334, 56);
            lbl_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_file_path.MaximumSize = new Size(509, 0);
            lbl_file_path.Name = "lbl_file_path";
            lbl_file_path.Size = new Size(194, 17);
            lbl_file_path.TabIndex = 3;
            lbl_file_path.Text = "或将 表格（cdb周报） 拖拽到此处";
            // 
            // lbl_cdb_password
            // 
            lbl_cdb_password.AutoSize = true;
            lbl_cdb_password.Location = new Point(13, 152);
            lbl_cdb_password.Name = "lbl_cdb_password";
            lbl_cdb_password.Size = new Size(90, 17);
            lbl_cdb_password.TabIndex = 13;
            lbl_cdb_password.Text = "cdb周报密码：";
            // 
            // txt_cdb_password
            // 
            txt_cdb_password.Location = new Point(96, 149);
            txt_cdb_password.Name = "txt_cdb_password";
            txt_cdb_password.Size = new Size(100, 23);
            txt_cdb_password.TabIndex = 14;
            // 
            // OnceOpenTranslateForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(718, 294);
            Controls.Add(txt_cdb_password);
            Controls.Add(lbl_cdb_password);
            Controls.Add(gbx_file);
            Controls.Add(txt_check_result);
            Controls.Add(lbl_task_title);
            Controls.Add(lbl_task);
            Controls.Add(lbl_remark);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "OnceOpenTranslateForm";
            Text = "韩文翻译工具";
            gbx_file.ResumeLayout(false);
            gbx_file.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox txt_check_result;
        private Label lbl_task;
        private Label lbl_task_title;
        private Label lbl_remark;
        private GroupBox gbx_file;
        private Button btn_select_file;
        private Label lbl_file_path;
        private Label lbl_cdb_password;
        private TextBox txt_cdb_password;
    }
}
