namespace BatchTranslateApp
{
    partial class NormalTranslateForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NormalTranslateForm));
            txt_check_result = new TextBox();
            lbl_task = new Label();
            lbl_task_title = new Label();
            lbl_remark = new Label();
            gbx_file = new GroupBox();
            btn_select_file = new Button();
            lbl_file_path = new Label();
            gbx_file.SuspendLayout();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(373, 180);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(333, 102);
            txt_check_result.TabIndex = 6;
            // 
            // lbl_task
            // 
            lbl_task.AutoSize = true;
            lbl_task.Location = new Point(224, 219);
            lbl_task.Margin = new Padding(4, 0, 4, 0);
            lbl_task.Name = "lbl_task";
            lbl_task.Size = new Size(56, 17);
            lbl_task.TabIndex = 9;
            lbl_task.Text = "暂未开始";
            // 
            // lbl_task_title
            // 
            lbl_task_title.AutoSize = true;
            lbl_task_title.Location = new Point(96, 219);
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
            lbl_remark.Location = new Point(13, 265);
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
            gbx_file.Size = new Size(692, 160);
            gbx_file.TabIndex = 12;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择表格";
            // 
            // btn_select_file
            // 
            btn_select_file.Location = new Point(224, 69);
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
            lbl_file_path.Location = new Point(342, 75);
            lbl_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_file_path.MaximumSize = new Size(509, 0);
            lbl_file_path.Name = "lbl_file_path";
            lbl_file_path.Size = new Size(124, 17);
            lbl_file_path.TabIndex = 3;
            lbl_file_path.Text = "或将 表格 拖拽到此处";
            // 
            // NormalTranslateForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(718, 294);
            Controls.Add(gbx_file);
            Controls.Add(txt_check_result);
            Controls.Add(lbl_task_title);
            Controls.Add(lbl_task);
            Controls.Add(lbl_remark);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "NormalTranslateForm";
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
    }
}
