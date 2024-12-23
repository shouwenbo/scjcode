
namespace DropReasonCreator
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
            lbl_version = new Label();
            gbx_file = new GroupBox();
            btn_select_dir = new Button();
            lbl_file_path = new Label();
            gbx_file.SuspendLayout();
            SuspendLayout();
            // 
            // txt_check_result
            // 
            txt_check_result.Location = new Point(12, 84);
            txt_check_result.Margin = new Padding(2);
            txt_check_result.Multiline = true;
            txt_check_result.Name = "txt_check_result";
            txt_check_result.ReadOnly = true;
            txt_check_result.ScrollBars = ScrollBars.Vertical;
            txt_check_result.Size = new Size(324, 65);
            txt_check_result.TabIndex = 8;
            // 
            // lbl_version
            // 
            lbl_version.AutoSize = true;
            lbl_version.Location = new Point(224, 151);
            lbl_version.Name = "lbl_version";
            lbl_version.Size = new Size(112, 17);
            lbl_version.TabIndex = 7;
            lbl_version.Text = "版本号：41011204";
            // 
            // gbx_file
            // 
            gbx_file.Controls.Add(btn_select_dir);
            gbx_file.Controls.Add(lbl_file_path);
            gbx_file.Location = new Point(12, 12);
            gbx_file.Name = "gbx_file";
            gbx_file.Size = new Size(324, 67);
            gbx_file.TabIndex = 6;
            gbx_file.TabStop = false;
            gbx_file.Text = "请选择文件 或 将文件拖拽到此区域";
            gbx_file.DragDrop += gbx_file_DragDrop;
            gbx_file.DragEnter += gbx_file_DragEnter;
            // 
            // btn_select_dir
            // 
            btn_select_dir.Location = new Point(10, 29);
            btn_select_dir.Margin = new Padding(2);
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
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(345, 170);
            Controls.Add(txt_check_result);
            Controls.Add(lbl_version);
            Controls.Add(gbx_file);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "MainForm";
            Text = "整理韩文掉落事由助手";
            gbx_file.ResumeLayout(false);
            gbx_file.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private TextBox txt_check_result;
        private Label lbl_version;
        private GroupBox gbx_file;
        private Button btn_select_dir;
        private Label lbl_file_path;
    }
}
