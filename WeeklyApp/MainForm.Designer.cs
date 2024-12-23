namespace WeeklyApp
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
            lbl_version = new Label();
            lbl_select_zbjl_file_1 = new Label();
            btn_zbjl_select_file = new Button();
            lbl_zbjl_file_path = new Label();
            lbl_select_zbjl_file_2 = new Label();
            lbl_select_zbjl_file_3 = new Label();
            lbl_select_dlbg_file_3 = new Label();
            lbl_select_dlbg_file_2 = new Label();
            lbl_dlbg_file_path = new Label();
            btn_dlbg_select_file = new Button();
            lbl_select_dlbg_file_1 = new Label();
            btn_yjsc = new Button();
            button1 = new Button();
            lbl_select_hwzb_file_3 = new Label();
            lbl_select_hwzb_file_2 = new Label();
            lbl_hwzb_file_path = new Label();
            btn_hwzb_select_file = new Button();
            lbl_select_hwzb_file_1 = new Label();
            label1 = new Label();
            label2 = new Label();
            SuspendLayout();
            // 
            // lbl_version
            // 
            lbl_version.AutoSize = true;
            lbl_version.Location = new Point(375, 73);
            lbl_version.Margin = new Padding(2, 0, 2, 0);
            lbl_version.Name = "lbl_version";
            lbl_version.Size = new Size(117, 17);
            lbl_version.TabIndex = 1;
            lbl_version.Text = "版本号：411208_01";
            // 
            // lbl_select_zbjl_file_1
            // 
            lbl_select_zbjl_file_1.AutoSize = true;
            lbl_select_zbjl_file_1.Location = new Point(15, 134);
            lbl_select_zbjl_file_1.Margin = new Padding(2, 0, 2, 0);
            lbl_select_zbjl_file_1.Name = "lbl_select_zbjl_file_1";
            lbl_select_zbjl_file_1.Size = new Size(44, 17);
            lbl_select_zbjl_file_1.TabIndex = 3;
            lbl_select_zbjl_file_1.Text = "请选择";
            // 
            // btn_zbjl_select_file
            // 
            btn_zbjl_select_file.Location = new Point(146, 130);
            btn_zbjl_select_file.Margin = new Padding(2);
            btn_zbjl_select_file.Name = "btn_zbjl_select_file";
            btn_zbjl_select_file.Size = new Size(71, 24);
            btn_zbjl_select_file.TabIndex = 4;
            btn_zbjl_select_file.Text = "选择文件";
            btn_zbjl_select_file.UseVisualStyleBackColor = true;
            btn_zbjl_select_file.Click += btn_zbjl_select_file_Click;
            // 
            // lbl_zbjl_file_path
            // 
            lbl_zbjl_file_path.AutoSize = true;
            lbl_zbjl_file_path.ForeColor = Color.Red;
            lbl_zbjl_file_path.Location = new Point(15, 161);
            lbl_zbjl_file_path.Margin = new Padding(2, 0, 2, 0);
            lbl_zbjl_file_path.Name = "lbl_zbjl_file_path";
            lbl_zbjl_file_path.Size = new Size(406, 17);
            lbl_zbjl_file_path.TabIndex = 5;
            lbl_zbjl_file_path.Text = "首先打开在线报告<周报接龙> → 文件 → 下载格式 → .xlsx → 保存到本地";
            // 
            // lbl_select_zbjl_file_2
            // 
            lbl_select_zbjl_file_2.AutoSize = true;
            lbl_select_zbjl_file_2.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_select_zbjl_file_2.Location = new Point(51, 134);
            lbl_select_zbjl_file_2.Margin = new Padding(2, 0, 2, 0);
            lbl_select_zbjl_file_2.Name = "lbl_select_zbjl_file_2";
            lbl_select_zbjl_file_2.Size = new Size(56, 17);
            lbl_select_zbjl_file_2.TabIndex = 6;
            lbl_select_zbjl_file_2.Text = "周报接龙";
            // 
            // lbl_select_zbjl_file_3
            // 
            lbl_select_zbjl_file_3.AutoSize = true;
            lbl_select_zbjl_file_3.Location = new Point(102, 134);
            lbl_select_zbjl_file_3.Margin = new Padding(2, 0, 2, 0);
            lbl_select_zbjl_file_3.Name = "lbl_select_zbjl_file_3";
            lbl_select_zbjl_file_3.Size = new Size(44, 17);
            lbl_select_zbjl_file_3.TabIndex = 7;
            lbl_select_zbjl_file_3.Text = "表格：";
            // 
            // lbl_select_dlbg_file_3
            // 
            lbl_select_dlbg_file_3.AutoSize = true;
            lbl_select_dlbg_file_3.Location = new Point(102, 185);
            lbl_select_dlbg_file_3.Margin = new Padding(2, 0, 2, 0);
            lbl_select_dlbg_file_3.Name = "lbl_select_dlbg_file_3";
            lbl_select_dlbg_file_3.Size = new Size(44, 17);
            lbl_select_dlbg_file_3.TabIndex = 12;
            lbl_select_dlbg_file_3.Text = "表格：";
            // 
            // lbl_select_dlbg_file_2
            // 
            lbl_select_dlbg_file_2.AutoSize = true;
            lbl_select_dlbg_file_2.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_select_dlbg_file_2.Location = new Point(51, 185);
            lbl_select_dlbg_file_2.Margin = new Padding(2, 0, 2, 0);
            lbl_select_dlbg_file_2.Name = "lbl_select_dlbg_file_2";
            lbl_select_dlbg_file_2.Size = new Size(56, 17);
            lbl_select_dlbg_file_2.TabIndex = 11;
            lbl_select_dlbg_file_2.Text = "掉落报告";
            // 
            // lbl_dlbg_file_path
            // 
            lbl_dlbg_file_path.AutoSize = true;
            lbl_dlbg_file_path.ForeColor = Color.Red;
            lbl_dlbg_file_path.Location = new Point(15, 211);
            lbl_dlbg_file_path.Margin = new Padding(2, 0, 2, 0);
            lbl_dlbg_file_path.Name = "lbl_dlbg_file_path";
            lbl_dlbg_file_path.Size = new Size(406, 17);
            lbl_dlbg_file_path.TabIndex = 10;
            lbl_dlbg_file_path.Text = "首先打开在线报告<掉落报告> → 文件 → 下载格式 → .xlsx → 保存到本地";
            // 
            // btn_dlbg_select_file
            // 
            btn_dlbg_select_file.Location = new Point(146, 181);
            btn_dlbg_select_file.Margin = new Padding(2);
            btn_dlbg_select_file.Name = "btn_dlbg_select_file";
            btn_dlbg_select_file.Size = new Size(71, 24);
            btn_dlbg_select_file.TabIndex = 9;
            btn_dlbg_select_file.Text = "选择文件";
            btn_dlbg_select_file.UseVisualStyleBackColor = true;
            btn_dlbg_select_file.Click += btn_dlbg_select_file_Click;
            // 
            // lbl_select_dlbg_file_1
            // 
            lbl_select_dlbg_file_1.AutoSize = true;
            lbl_select_dlbg_file_1.Location = new Point(15, 185);
            lbl_select_dlbg_file_1.Margin = new Padding(2, 0, 2, 0);
            lbl_select_dlbg_file_1.Name = "lbl_select_dlbg_file_1";
            lbl_select_dlbg_file_1.Size = new Size(44, 17);
            lbl_select_dlbg_file_1.TabIndex = 8;
            lbl_select_dlbg_file_1.Text = "请选择";
            // 
            // btn_yjsc
            // 
            btn_yjsc.Location = new Point(15, 305);
            btn_yjsc.Margin = new Padding(2);
            btn_yjsc.Name = "btn_yjsc";
            btn_yjsc.Size = new Size(71, 24);
            btn_yjsc.TabIndex = 13;
            btn_yjsc.Text = "一键生成";
            btn_yjsc.UseVisualStyleBackColor = true;
            btn_yjsc.Click += btn_yjsc_Click;
            // 
            // button1
            // 
            button1.Font = new Font("Microsoft YaHei UI", 16F, FontStyle.Regular, GraphicsUnit.Point);
            button1.Location = new Point(15, 11);
            button1.Margin = new Padding(2);
            button1.Name = "button1";
            button1.Size = new Size(327, 50);
            button1.TabIndex = 14;
            button1.Text = "一键生成韩文掉落报告";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // lbl_select_hwzb_file_3
            // 
            lbl_select_hwzb_file_3.AutoSize = true;
            lbl_select_hwzb_file_3.Location = new Point(102, 235);
            lbl_select_hwzb_file_3.Margin = new Padding(2, 0, 2, 0);
            lbl_select_hwzb_file_3.Name = "lbl_select_hwzb_file_3";
            lbl_select_hwzb_file_3.Size = new Size(44, 17);
            lbl_select_hwzb_file_3.TabIndex = 19;
            lbl_select_hwzb_file_3.Text = "表格：";
            // 
            // lbl_select_hwzb_file_2
            // 
            lbl_select_hwzb_file_2.AutoSize = true;
            lbl_select_hwzb_file_2.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_select_hwzb_file_2.Location = new Point(51, 235);
            lbl_select_hwzb_file_2.Margin = new Padding(2, 0, 2, 0);
            lbl_select_hwzb_file_2.Name = "lbl_select_hwzb_file_2";
            lbl_select_hwzb_file_2.Size = new Size(56, 17);
            lbl_select_hwzb_file_2.TabIndex = 18;
            lbl_select_hwzb_file_2.Text = "韩文周报";
            // 
            // lbl_hwzb_file_path
            // 
            lbl_hwzb_file_path.ForeColor = Color.Red;
            lbl_hwzb_file_path.Location = new Point(15, 262);
            lbl_hwzb_file_path.Margin = new Padding(2, 0, 2, 0);
            lbl_hwzb_file_path.Name = "lbl_hwzb_file_path";
            lbl_hwzb_file_path.Size = new Size(620, 40);
            lbl_hwzb_file_path.TabIndex = 17;
            lbl_hwzb_file_path.Text = "请选择上周提交的韩文周报“”";
            // 
            // btn_hwzb_select_file
            // 
            btn_hwzb_select_file.Location = new Point(146, 232);
            btn_hwzb_select_file.Margin = new Padding(2);
            btn_hwzb_select_file.Name = "btn_hwzb_select_file";
            btn_hwzb_select_file.Size = new Size(71, 24);
            btn_hwzb_select_file.TabIndex = 16;
            btn_hwzb_select_file.Text = "选择文件";
            btn_hwzb_select_file.UseVisualStyleBackColor = true;
            btn_hwzb_select_file.Click += btn_hwzb_select_file_Click;
            // 
            // lbl_select_hwzb_file_1
            // 
            lbl_select_hwzb_file_1.AutoSize = true;
            lbl_select_hwzb_file_1.Location = new Point(15, 235);
            lbl_select_hwzb_file_1.Margin = new Padding(2, 0, 2, 0);
            lbl_select_hwzb_file_1.Name = "lbl_select_hwzb_file_1";
            lbl_select_hwzb_file_1.Size = new Size(44, 17);
            lbl_select_hwzb_file_1.TabIndex = 15;
            lbl_select_hwzb_file_1.Text = "请选择";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Microsoft YaHei UI", 30F, FontStyle.Bold, GraphicsUnit.Point);
            label1.ForeColor = Color.Red;
            label1.Location = new Point(15, 352);
            label1.Margin = new Padding(2, 0, 2, 0);
            label1.Name = "label1";
            label1.Size = new Size(474, 52);
            label1.TabIndex = 20;
            label1.Text = "↑↑ 上面的没做完，不用管";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Microsoft YaHei UI", 25F, FontStyle.Bold, GraphicsUnit.Point);
            label2.ForeColor = Color.Red;
            label2.Location = new Point(81, 456);
            label2.Margin = new Padding(2, 0, 2, 0);
            label2.Name = "label2";
            label2.Size = new Size(336, 45);
            label2.TabIndex = 21;
            label2.Text = "↓↓ 下面的是原有功能";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(503, 99);
            Controls.Add(button1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(lbl_select_hwzb_file_3);
            Controls.Add(lbl_select_hwzb_file_2);
            Controls.Add(lbl_hwzb_file_path);
            Controls.Add(btn_hwzb_select_file);
            Controls.Add(lbl_select_hwzb_file_1);
            Controls.Add(btn_yjsc);
            Controls.Add(lbl_select_dlbg_file_3);
            Controls.Add(lbl_select_dlbg_file_2);
            Controls.Add(lbl_dlbg_file_path);
            Controls.Add(btn_dlbg_select_file);
            Controls.Add(lbl_select_dlbg_file_1);
            Controls.Add(lbl_select_zbjl_file_3);
            Controls.Add(lbl_select_zbjl_file_2);
            Controls.Add(lbl_zbjl_file_path);
            Controls.Add(btn_zbjl_select_file);
            Controls.Add(lbl_select_zbjl_file_1);
            Controls.Add(lbl_version);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(2);
            MaximizeBox = false;
            Name = "MainForm";
            Text = "周报助手";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
        private Label lbl_version;
        private Label lbl_select_zbjl_file_1;
        private Button btn_zbjl_select_file;
        private Label lbl_zbjl_file_path;
        private Label lbl_select_zbjl_file_2;
        private Label lbl_select_zbjl_file_3;
        private Label lbl_select_dlbg_file_3;
        private Label lbl_select_dlbg_file_2;
        private Label lbl_dlbg_file_path;
        private Button btn_dlbg_select_file;
        private Label lbl_select_dlbg_file_1;
        private Button btn_yjsc;
        private Button button1;
        private Label lbl_select_hwzb_file_3;
        private Label lbl_select_hwzb_file_2;
        private Label lbl_hwzb_file_path;
        private Button btn_hwzb_select_file;
        private Label lbl_select_hwzb_file_1;
        private Label label1;
        private Label label2;
    }
}
