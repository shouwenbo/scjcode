namespace BatchPhotoRenameApp
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
            gbx_file = new GroupBox();
            btn_select_dir = new Button();
            lbl_file_path = new Label();
            lbl_version = new Label();
            lbl_file_count = new Label();
            btn_batch_rename = new Button();
            txt_check_result = new TextBox();
            gbx_file.SuspendLayout();
            SuspendLayout();
            // 
            // gbx_file
            // 
            gbx_file.Controls.Add(btn_select_dir);
            gbx_file.Controls.Add(lbl_file_path);
            gbx_file.Location = new Point(6, 9);
            gbx_file.Name = "gbx_file";
            gbx_file.Size = new Size(324, 67);
            gbx_file.TabIndex = 1;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择文件 或 将文件拖拽到此区域";
            gbx_file.DragDrop += gbx_file_DragDrop;
            gbx_file.DragEnter += gbx_file_DragEnter;
            // 
            // btn_select_dir
            // 
            btn_select_dir.Location = new Point(10, 29);
            btn_select_dir.Margin = new Padding(2, 2, 2, 2);
            btn_select_dir.Name = "btn_select_dir";
            btn_select_dir.Size = new Size(70, 24);
            btn_select_dir.TabIndex = 4;
            btn_select_dir.Text = "选择文件";
            btn_select_dir.UseVisualStyleBackColor = true;
            btn_select_dir.Click += btn_select_dir_Click;
            // 
            // lbl_file_path
            // 
            lbl_file_path.AutoSize = true;
            lbl_file_path.ImeMode = ImeMode.NoControl;
            lbl_file_path.Location = new Point(85, 33);
            lbl_file_path.MaximumSize = new Size(324, 0);
            lbl_file_path.Name = "lbl_file_path";
            lbl_file_path.Size = new Size(116, 17);
            lbl_file_path.TabIndex = 3;
            lbl_file_path.Text = "或将文件拖拽到此处";
            // 
            // lbl_version
            // 
            lbl_version.AutoSize = true;
            lbl_version.Location = new Point(220, 142);
            lbl_version.Name = "lbl_version";
            lbl_version.Size = new Size(112, 17);
            lbl_version.TabIndex = 2;
            lbl_version.Text = "版本号：41060601";
            // 
            // lbl_file_count
            // 
            lbl_file_count.AutoSize = true;
            lbl_file_count.Location = new Point(6, 76);
            lbl_file_count.Margin = new Padding(2, 0, 2, 0);
            lbl_file_count.Name = "lbl_file_count";
            lbl_file_count.Size = new Size(0, 17);
            lbl_file_count.TabIndex = 3;
            // 
            // btn_batch_rename
            // 
            btn_batch_rename.Enabled = false;
            btn_batch_rename.Location = new Point(219, 95);
            btn_batch_rename.Margin = new Padding(2, 2, 2, 2);
            btn_batch_rename.Name = "btn_batch_rename";
            btn_batch_rename.Size = new Size(109, 24);
            btn_batch_rename.TabIndex = 4;
            btn_batch_rename.Text = "批量重命名";
            btn_batch_rename.UseVisualStyleBackColor = true;
            btn_batch_rename.Click += btn_batch_rename_Click;
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(8, 95);
            txt_check_result.Margin = new Padding(2, 2, 2, 2);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(209, 65);
            txt_check_result.TabIndex = 5;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(338, 165);
            Controls.Add(txt_check_result);
            Controls.Add(btn_batch_rename);
            Controls.Add(lbl_file_count);
            Controls.Add(lbl_version);
            Controls.Add(gbx_file);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(2, 2, 2, 2);
            MaximizeBox = false;
            Name = "MainForm";
            Text = "批量修改照片名称";
            gbx_file.ResumeLayout(false);
            gbx_file.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox gbx_file;
        private Button btn_select_dir;
        private Label lbl_file_path;
        private Label lbl_version;
        private Label lbl_file_count;
        private Button btn_batch_rename;
        private TextBox txt_check_result;
    }
}
