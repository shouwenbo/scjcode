namespace BatchTranslateApp
{
    partial class Mainform
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Mainform));
            panel_content = new Panel();
            lbl_title = new Label();
            panel_tab_list = new Panel();
            btn_twice_open_translate = new Button();
            btn_once_open_translate = new Button();
            btn_jjb_translate = new Button();
            btn_normal_translate = new Button();
            panel_content.SuspendLayout();
            panel_tab_list.SuspendLayout();
            SuspendLayout();
            // 
            // panel_content
            // 
            panel_content.Controls.Add(lbl_title);
            panel_content.Location = new Point(155, 12);
            panel_content.Name = "panel_content";
            panel_content.Size = new Size(758, 332);
            panel_content.TabIndex = 0;
            // 
            // lbl_title
            // 
            lbl_title.AutoSize = true;
            lbl_title.Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold, GraphicsUnit.Point);
            lbl_title.ForeColor = Color.Blue;
            lbl_title.Location = new Point(0, 0);
            lbl_title.Name = "lbl_title";
            lbl_title.Size = new Size(45, 17);
            lbl_title.TabIndex = 0;
            lbl_title.Text = "label1";
            // 
            // panel_tab_list
            // 
            panel_tab_list.Controls.Add(btn_twice_open_translate);
            panel_tab_list.Controls.Add(btn_once_open_translate);
            panel_tab_list.Controls.Add(btn_jjb_translate);
            panel_tab_list.Controls.Add(btn_normal_translate);
            panel_tab_list.Location = new Point(12, 12);
            panel_tab_list.Name = "panel_tab_list";
            panel_tab_list.Size = new Size(137, 332);
            panel_tab_list.TabIndex = 1;
            // 
            // btn_twice_open_translate
            // 
            btn_twice_open_translate.Location = new Point(0, 87);
            btn_twice_open_translate.Name = "btn_twice_open_translate";
            btn_twice_open_translate.Size = new Size(137, 23);
            btn_twice_open_translate.TabIndex = 3;
            btn_twice_open_translate.Text = "二开翻译";
            btn_twice_open_translate.UseVisualStyleBackColor = true;
            // 
            // btn_once_open_translate
            // 
            btn_once_open_translate.Location = new Point(0, 58);
            btn_once_open_translate.Name = "btn_once_open_translate";
            btn_once_open_translate.Size = new Size(137, 23);
            btn_once_open_translate.TabIndex = 2;
            btn_once_open_translate.Text = "一开翻译";
            btn_once_open_translate.UseVisualStyleBackColor = true;
            // 
            // btn_jjb_translate
            // 
            btn_jjb_translate.Location = new Point(0, 29);
            btn_jjb_translate.Name = "btn_jjb_translate";
            btn_jjb_translate.Size = new Size(137, 23);
            btn_jjb_translate.TabIndex = 1;
            btn_jjb_translate.Text = "JJB翻译";
            btn_jjb_translate.UseVisualStyleBackColor = true;
            // 
            // btn_normal_translate
            // 
            btn_normal_translate.Location = new Point(0, 0);
            btn_normal_translate.Name = "btn_normal_translate";
            btn_normal_translate.Size = new Size(137, 23);
            btn_normal_translate.TabIndex = 0;
            btn_normal_translate.Text = "普通翻译";
            btn_normal_translate.UseVisualStyleBackColor = true;
            // 
            // Mainform
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(925, 356);
            Controls.Add(panel_tab_list);
            Controls.Add(panel_content);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            Name = "Mainform";
            Text = "韩文翻译工具（版本：1.0.2）";
            FormClosing += Mainform_FormClosing;
            panel_content.ResumeLayout(false);
            panel_content.PerformLayout();
            panel_tab_list.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private Panel panel_content;
        private Panel panel_tab_list;
        private Button btn_normal_translate;
        private Label lbl_title;
        private Button btn_jjb_translate;
        private Button btn_once_open_translate;
        private Button btn_twice_open_translate;
    }
}