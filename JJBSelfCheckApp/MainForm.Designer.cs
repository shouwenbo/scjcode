namespace JJBSelfCheckApp
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
            lbl_file_path = new Label();
            btn_select_file = new Button();
            lbl_version = new Label();
            gbx_file.SuspendLayout();
            SuspendLayout();
            // 
            // gbx_file
            // 
            gbx_file.Controls.Add(lbl_file_path);
            gbx_file.Controls.Add(btn_select_file);
            gbx_file.Location = new Point(15, 16);
            gbx_file.Margin = new Padding(4);
            gbx_file.Name = "gbx_file";
            gbx_file.Padding = new Padding(4);
            gbx_file.Size = new Size(778, 132);
            gbx_file.TabIndex = 0;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择文件 或 将文件拖拽到此区域";
            gbx_file.DragDrop += gbx_file_DragDrop;
            gbx_file.DragEnter += gbx_file_DragEnter;
            // 
            // lbl_file_path
            // 
            lbl_file_path.AutoSize = true;
            lbl_file_path.ImeMode = ImeMode.NoControl;
            lbl_file_path.Location = new Point(195, 62);
            lbl_file_path.Margin = new Padding(4, 0, 4, 0);
            lbl_file_path.MaximumSize = new Size(509, 0);
            lbl_file_path.Name = "lbl_file_path";
            lbl_file_path.Size = new Size(230, 31);
            lbl_file_path.TabIndex = 3;
            lbl_file_path.Text = "或将文件拖拽到此处";
            // 
            // btn_select_file
            // 
            btn_select_file.ImeMode = ImeMode.NoControl;
            btn_select_file.Location = new Point(32, 56);
            btn_select_file.Margin = new Padding(4);
            btn_select_file.Name = "btn_select_file";
            btn_select_file.Size = new Size(143, 44);
            btn_select_file.TabIndex = 2;
            btn_select_file.Text = "选择文件";
            btn_select_file.UseVisualStyleBackColor = true;
            btn_select_file.Click += btn_select_file_Click;
            // 
            // lbl_version
            // 
            lbl_version.AutoSize = true;
            lbl_version.Location = new Point(577, 160);
            lbl_version.Margin = new Padding(4, 0, 4, 0);
            lbl_version.Name = "lbl_version";
            lbl_version.Size = new Size(222, 31);
            lbl_version.TabIndex = 1;
            lbl_version.Text = "版本号：41061802";
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(14F, 31F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(808, 203);
            Controls.Add(lbl_version);
            Controls.Add(gbx_file);
            Font = new Font("Microsoft YaHei UI", 12F, FontStyle.Regular, GraphicsUnit.Point);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            Margin = new Padding(4);
            MaximizeBox = false;
            Name = "MainForm";
            Text = "jjb自测";
            gbx_file.ResumeLayout(false);
            gbx_file.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private GroupBox gbx_file;
        private Label lbl_file_path;
        private Button btn_select_file;
        private Label lbl_version;
    }
}
