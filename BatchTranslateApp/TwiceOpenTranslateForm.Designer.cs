namespace BatchTranslateApp
{
    partial class TwiceOpenTranslateForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TwiceOpenTranslateForm));
            txt_check_result = new TextBox();
            lbl_task = new Label();
            lbl_task_title = new Label();
            lbl_remark = new Label();
            gbx_once_file = new GroupBox();
            lbl_once_file_name = new Label();
            btn_select_once_file = new Button();
            lbl_once_file_path = new Label();
            lbl_once_password = new Label();
            txt_once_password = new TextBox();
            btn_select_drop_file = new Button();
            lbl_drop_file_path = new Label();
            gbx_drop_file = new GroupBox();
            lbl_drop_file_name = new Label();
            btn_select_change_file = new Button();
            lbl_change_file_path = new Label();
            gbx_change_file = new GroupBox();
            lbl_change_file_name = new Label();
            btn_run = new Button();
            txt_drop_password = new TextBox();
            lbl_drop_password = new Label();
            txt_change_password = new TextBox();
            lbl_change_password = new Label();
            gbx_once_file.SuspendLayout();
            gbx_drop_file.SuspendLayout();
            gbx_change_file.SuspendLayout();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(409, 173);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(296, 112);
            txt_check_result.TabIndex = 6;
            // 
            // lbl_task
            // 
            lbl_task.AutoSize = true;
            lbl_task.Location = new Point(243, 213);
            lbl_task.Margin = new Padding(4, 0, 4, 0);
            lbl_task.Name = "lbl_task";
            lbl_task.Size = new Size(56, 17);
            lbl_task.TabIndex = 9;
            lbl_task.Text = "暂未开始";
            // 
            // lbl_task_title
            // 
            lbl_task_title.AutoSize = true;
            lbl_task_title.Location = new Point(115, 213);
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
            // gbx_once_file
            // 
            gbx_once_file.Controls.Add(lbl_once_file_name);
            gbx_once_file.Controls.Add(btn_select_once_file);
            gbx_once_file.Controls.Add(lbl_once_file_path);
            gbx_once_file.Location = new Point(13, 13);
            gbx_once_file.Margin = new Padding(4);
            gbx_once_file.Name = "gbx_once_file";
            gbx_once_file.Padding = new Padding(4);
            gbx_once_file.Size = new Size(224, 129);
            gbx_once_file.TabIndex = 12;
            gbx_once_file.TabStop = false;
            gbx_once_file.Text = "请选择表格（一开）";
            // 
            // lbl_once_file_name
            // 
            lbl_once_file_name.Font = new Font("Microsoft YaHei UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_once_file_name.ForeColor = Color.Green;
            lbl_once_file_name.Location = new Point(3, 75);
            lbl_once_file_name.Name = "lbl_once_file_name";
            lbl_once_file_name.Size = new Size(214, 50);
            lbl_once_file_name.TabIndex = 5;
            lbl_once_file_name.Text = "未选择文件";
            // 
            // btn_select_once_file
            // 
            btn_select_once_file.Location = new Point(53, 21);
            btn_select_once_file.Name = "btn_select_once_file";
            btn_select_once_file.Size = new Size(110, 34);
            btn_select_once_file.TabIndex = 4;
            btn_select_once_file.Text = "点此选择";
            btn_select_once_file.UseVisualStyleBackColor = true;
            // 
            // lbl_once_file_path
            // 
            lbl_once_file_path.AutoSize = true;
            lbl_once_file_path.ImeMode = ImeMode.NoControl;
            lbl_once_file_path.Location = new Point(18, 58);
            lbl_once_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_once_file_path.MaximumSize = new Size(509, 0);
            lbl_once_file_path.Name = "lbl_once_file_path";
            lbl_once_file_path.Size = new Size(172, 17);
            lbl_once_file_path.TabIndex = 3;
            lbl_once_file_path.Text = "或将 表格（一开） 拖拽到此处";
            // 
            // lbl_once_password
            // 
            lbl_once_password.AutoSize = true;
            lbl_once_password.Location = new Point(13, 145);
            lbl_once_password.Name = "lbl_once_password";
            lbl_once_password.Size = new Size(116, 17);
            lbl_once_password.TabIndex = 13;
            lbl_once_password.Text = "表格（一开）密码：";
            // 
            // txt_once_password
            // 
            txt_once_password.Location = new Point(137, 142);
            txt_once_password.Name = "txt_once_password";
            txt_once_password.Size = new Size(100, 23);
            txt_once_password.TabIndex = 14;
            // 
            // btn_select_drop_file
            // 
            btn_select_drop_file.Location = new Point(59, 21);
            btn_select_drop_file.Name = "btn_select_drop_file";
            btn_select_drop_file.Size = new Size(110, 34);
            btn_select_drop_file.TabIndex = 4;
            btn_select_drop_file.Text = "点此选择";
            btn_select_drop_file.UseVisualStyleBackColor = true;
            // 
            // lbl_drop_file_path
            // 
            lbl_drop_file_path.AutoSize = true;
            lbl_drop_file_path.ImeMode = ImeMode.NoControl;
            lbl_drop_file_path.Location = new Point(29, 58);
            lbl_drop_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_drop_file_path.MaximumSize = new Size(509, 0);
            lbl_drop_file_path.Name = "lbl_drop_file_path";
            lbl_drop_file_path.Size = new Size(172, 17);
            lbl_drop_file_path.TabIndex = 3;
            lbl_drop_file_path.Text = "或将 表格（韩掉） 拖拽到此处";
            // 
            // gbx_drop_file
            // 
            gbx_drop_file.Controls.Add(lbl_drop_file_name);
            gbx_drop_file.Controls.Add(btn_select_drop_file);
            gbx_drop_file.Controls.Add(lbl_drop_file_path);
            gbx_drop_file.Location = new Point(247, 13);
            gbx_drop_file.Margin = new Padding(4);
            gbx_drop_file.Name = "gbx_drop_file";
            gbx_drop_file.Padding = new Padding(4);
            gbx_drop_file.Size = new Size(224, 129);
            gbx_drop_file.TabIndex = 13;
            gbx_drop_file.TabStop = false;
            gbx_drop_file.Text = "请选择表格（韩掉）";
            // 
            // lbl_drop_file_name
            // 
            lbl_drop_file_name.Font = new Font("Microsoft YaHei UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_drop_file_name.ForeColor = Color.Green;
            lbl_drop_file_name.Location = new Point(3, 75);
            lbl_drop_file_name.Name = "lbl_drop_file_name";
            lbl_drop_file_name.Size = new Size(214, 50);
            lbl_drop_file_name.TabIndex = 6;
            lbl_drop_file_name.Text = "未选择文件";
            // 
            // btn_select_change_file
            // 
            btn_select_change_file.Location = new Point(56, 21);
            btn_select_change_file.Name = "btn_select_change_file";
            btn_select_change_file.Size = new Size(110, 34);
            btn_select_change_file.TabIndex = 4;
            btn_select_change_file.Text = "点此选择";
            btn_select_change_file.UseVisualStyleBackColor = true;
            // 
            // lbl_change_file_path
            // 
            lbl_change_file_path.AutoSize = true;
            lbl_change_file_path.ImeMode = ImeMode.NoControl;
            lbl_change_file_path.Location = new Point(26, 58);
            lbl_change_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_change_file_path.MaximumSize = new Size(509, 0);
            lbl_change_file_path.Name = "lbl_change_file_path";
            lbl_change_file_path.Size = new Size(172, 17);
            lbl_change_file_path.TabIndex = 3;
            lbl_change_file_path.Text = "或将 表格（变更） 拖拽到此处";
            // 
            // gbx_change_file
            // 
            gbx_change_file.Controls.Add(lbl_change_file_name);
            gbx_change_file.Controls.Add(btn_select_change_file);
            gbx_change_file.Controls.Add(lbl_change_file_path);
            gbx_change_file.Location = new Point(481, 13);
            gbx_change_file.Margin = new Padding(4);
            gbx_change_file.Name = "gbx_change_file";
            gbx_change_file.Padding = new Padding(4);
            gbx_change_file.Size = new Size(224, 129);
            gbx_change_file.TabIndex = 14;
            gbx_change_file.TabStop = false;
            gbx_change_file.Text = "请选择表格（变更）";
            // 
            // lbl_change_file_name
            // 
            lbl_change_file_name.Font = new Font("Microsoft YaHei UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_change_file_name.ForeColor = Color.Green;
            lbl_change_file_name.Location = new Point(3, 75);
            lbl_change_file_name.Name = "lbl_change_file_name";
            lbl_change_file_name.Size = new Size(214, 50);
            lbl_change_file_name.TabIndex = 7;
            lbl_change_file_name.Text = "未选择文件";
            // 
            // btn_run
            // 
            btn_run.Location = new Point(247, 173);
            btn_run.Name = "btn_run";
            btn_run.Size = new Size(134, 26);
            btn_run.TabIndex = 15;
            btn_run.Text = "点此运行";
            btn_run.UseVisualStyleBackColor = true;
            btn_run.Click += btn_run_Click;
            // 
            // txt_drop_password
            // 
            txt_drop_password.Location = new Point(371, 142);
            txt_drop_password.Name = "txt_drop_password";
            txt_drop_password.Size = new Size(100, 23);
            txt_drop_password.TabIndex = 17;
            // 
            // lbl_drop_password
            // 
            lbl_drop_password.AutoSize = true;
            lbl_drop_password.Location = new Point(247, 145);
            lbl_drop_password.Name = "lbl_drop_password";
            lbl_drop_password.Size = new Size(116, 17);
            lbl_drop_password.TabIndex = 16;
            lbl_drop_password.Text = "表格（韩掉）密码：";
            // 
            // txt_change_password
            // 
            txt_change_password.Location = new Point(605, 142);
            txt_change_password.Name = "txt_change_password";
            txt_change_password.Size = new Size(100, 23);
            txt_change_password.TabIndex = 19;
            // 
            // lbl_change_password
            // 
            lbl_change_password.AutoSize = true;
            lbl_change_password.Location = new Point(481, 145);
            lbl_change_password.Name = "lbl_change_password";
            lbl_change_password.Size = new Size(116, 17);
            lbl_change_password.TabIndex = 18;
            lbl_change_password.Text = "表格（变更）密码：";
            // 
            // TwiceOpenTranslateForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(718, 294);
            Controls.Add(txt_change_password);
            Controls.Add(lbl_change_password);
            Controls.Add(txt_drop_password);
            Controls.Add(lbl_drop_password);
            Controls.Add(btn_run);
            Controls.Add(gbx_change_file);
            Controls.Add(gbx_drop_file);
            Controls.Add(txt_once_password);
            Controls.Add(lbl_once_password);
            Controls.Add(gbx_once_file);
            Controls.Add(txt_check_result);
            Controls.Add(lbl_task_title);
            Controls.Add(lbl_task);
            Controls.Add(lbl_remark);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "TwiceOpenTranslateForm";
            Text = "韩文翻译工具";
            gbx_once_file.ResumeLayout(false);
            gbx_once_file.PerformLayout();
            gbx_drop_file.ResumeLayout(false);
            gbx_drop_file.PerformLayout();
            gbx_change_file.ResumeLayout(false);
            gbx_change_file.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox txt_check_result;
        private Label lbl_task;
        private Label lbl_task_title;
        private Label lbl_remark;
        private GroupBox gbx_once_file;
        private Button btn_select_once_file;
        private Label lbl_once_file_path;
        private Label lbl_once_password;
        private TextBox txt_once_password;
        private Button btn_select_drop_file;
        private Label lbl_drop_file_path;
        private GroupBox gbx_drop_file;
        private Button btn_select_change_file;
        private Label lbl_change_file_path;
        private GroupBox gbx_change_file;
        private Button btn_run;
        private Label lbl_once_file_name;
        private Label lbl_drop_file_name;
        private Label lbl_change_file_name;
        private TextBox txt_drop_password;
        private Label lbl_drop_password;
        private TextBox txt_change_password;
        private Label lbl_change_password;
    }
}
