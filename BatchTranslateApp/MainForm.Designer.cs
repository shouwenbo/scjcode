namespace BatchTranslateApp
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            txt_check_result = new TextBox();
            lbl_task = new Label();
            lbl_task_title = new Label();
            lbl_remark = new Label();
            lbl_version = new Label();
            gbx_file = new GroupBox();
            btn_select_file = new Button();
            lbl_file_path = new Label();
            panel1 = new Panel();
            radio_jjb_mode = new RadioButton();
            radio_batch_mode = new RadioButton();
            gbx_file.SuspendLayout();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(399, 22);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(342, 145);
            txt_check_result.TabIndex = 6;
            // 
            // lbl_task
            // 
            lbl_task.AutoSize = true;
            lbl_task.Location = new Point(332, 150);
            lbl_task.Margin = new Padding(4, 0, 4, 0);
            lbl_task.Name = "lbl_task";
            lbl_task.Size = new Size(56, 17);
            lbl_task.TabIndex = 9;
            lbl_task.Text = "暂未开始";
            lbl_task.Click += lbl_task_Click;
            // 
            // lbl_task_title
            // 
            lbl_task_title.AutoSize = true;
            lbl_task_title.Location = new Point(204, 150);
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
            lbl_remark.Location = new Point(11, 175);
            lbl_remark.Margin = new Padding(2, 0, 2, 0);
            lbl_remark.Name = "lbl_remark";
            lbl_remark.Size = new Size(193, 17);
            lbl_remark.TabIndex = 10;
            lbl_remark.Text = "翻译来源：hanwenxingming.com";
            lbl_remark.Click += lbl_remark_Click;
            // 
            // lbl_version
            // 
            lbl_version.AutoSize = true;
            lbl_version.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            lbl_version.Location = new Point(629, 175);
            lbl_version.Name = "lbl_version";
            lbl_version.Size = new Size(112, 17);
            lbl_version.TabIndex = 11;
            lbl_version.Text = "版本号：41122601";
            lbl_version.Click += lbl_version_Click;
            // 
            // gbx_file
            // 
            gbx_file.Controls.Add(btn_select_file);
            gbx_file.Controls.Add(lbl_file_path);
            gbx_file.Location = new Point(10, 13);
            gbx_file.Margin = new Padding(4);
            gbx_file.Name = "gbx_file";
            gbx_file.Padding = new Padding(4);
            gbx_file.Size = new Size(378, 94);
            gbx_file.TabIndex = 12;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择表格";
            gbx_file.DragDrop += gbx_file_DragDrop;
            gbx_file.DragEnter += gbx_file_DragEnter;
            // 
            // btn_select_file
            // 
            btn_select_file.Location = new Point(14, 38);
            btn_select_file.Name = "btn_select_file";
            btn_select_file.Size = new Size(110, 34);
            btn_select_file.TabIndex = 4;
            btn_select_file.Text = "点此选择";
            btn_select_file.UseVisualStyleBackColor = true;
            btn_select_file.Click += btn_select_file_Click;
            // 
            // lbl_file_path
            // 
            lbl_file_path.AutoSize = true;
            lbl_file_path.ImeMode = ImeMode.NoControl;
            lbl_file_path.Location = new Point(132, 44);
            lbl_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_file_path.MaximumSize = new Size(509, 0);
            lbl_file_path.Name = "lbl_file_path";
            lbl_file_path.Size = new Size(124, 17);
            lbl_file_path.TabIndex = 3;
            lbl_file_path.Text = "或将 表格 拖拽到此处";
            // 
            // panel1
            // 
            panel1.Controls.Add(radio_jjb_mode);
            panel1.Controls.Add(radio_batch_mode);
            panel1.Location = new Point(12, 114);
            panel1.Name = "panel1";
            panel1.Size = new Size(376, 33);
            panel1.TabIndex = 13;
            // 
            // radio_jjb_mode
            // 
            radio_jjb_mode.AutoSize = true;
            radio_jjb_mode.Location = new Point(103, 3);
            radio_jjb_mode.Name = "radio_jjb_mode";
            radio_jjb_mode.Size = new Size(64, 21);
            radio_jjb_mode.TabIndex = 1;
            radio_jjb_mode.Text = "jjb模式";
            radio_jjb_mode.UseVisualStyleBackColor = true;
            // 
            // radio_batch_mode
            // 
            radio_batch_mode.AutoSize = true;
            radio_batch_mode.Checked = true;
            radio_batch_mode.Location = new Point(12, 3);
            radio_batch_mode.Name = "radio_batch_mode";
            radio_batch_mode.Size = new Size(74, 21);
            radio_batch_mode.TabIndex = 0;
            radio_batch_mode.TabStop = true;
            radio_batch_mode.Text = "普通模式";
            radio_batch_mode.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(751, 202);
            Controls.Add(panel1);
            Controls.Add(gbx_file);
            Controls.Add(txt_check_result);
            Controls.Add(lbl_task_title);
            Controls.Add(lbl_version);
            Controls.Add(lbl_task);
            Controls.Add(lbl_remark);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "MainForm";
            Text = "韩文翻译工具";
            gbx_file.ResumeLayout(false);
            gbx_file.PerformLayout();
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private TextBox txt_check_result;
        private Label lbl_task;
        private Label lbl_task_title;
        private Label lbl_remark;
        private Label lbl_version;
        private GroupBox gbx_file;
        private Button btn_select_file;
        private Label lbl_file_path;
        private Panel panel1;
        private RadioButton radio_batch_mode;
        private RadioButton radio_jjb_mode;
    }
}
