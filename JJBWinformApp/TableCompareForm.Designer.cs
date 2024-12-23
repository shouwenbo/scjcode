namespace JJBWinformApp
{
    partial class TableCompareForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TableCompareForm));
            this.lbl_ori_table_path = new System.Windows.Forms.Label();
            this.btn_ori_table = new System.Windows.Forms.Button();
            this.btn_new_table = new System.Windows.Forms.Button();
            this.lbl_new_table_path = new System.Windows.Forms.Label();
            this.btn_start_compare = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lbl_ori_table_path
            // 
            this.lbl_ori_table_path.AutoSize = true;
            this.lbl_ori_table_path.Location = new System.Drawing.Point(212, 125);
            this.lbl_ori_table_path.Name = "lbl_ori_table_path";
            this.lbl_ori_table_path.Size = new System.Drawing.Size(118, 24);
            this.lbl_ori_table_path.TabIndex = 0;
            this.lbl_ori_table_path.Text = "请选择原表格";
            // 
            // btn_ori_table
            // 
            this.btn_ori_table.Location = new System.Drawing.Point(78, 120);
            this.btn_ori_table.Name = "btn_ori_table";
            this.btn_ori_table.Size = new System.Drawing.Size(112, 34);
            this.btn_ori_table.TabIndex = 1;
            this.btn_ori_table.Text = "选择原表格";
            this.btn_ori_table.UseVisualStyleBackColor = true;
            this.btn_ori_table.Click += new System.EventHandler(this.btn_ori_table_Click);
            // 
            // btn_new_table
            // 
            this.btn_new_table.Location = new System.Drawing.Point(78, 203);
            this.btn_new_table.Name = "btn_new_table";
            this.btn_new_table.Size = new System.Drawing.Size(112, 34);
            this.btn_new_table.TabIndex = 3;
            this.btn_new_table.Text = "选择新表格";
            this.btn_new_table.UseVisualStyleBackColor = true;
            this.btn_new_table.Click += new System.EventHandler(this.btn_new_table_Click);
            // 
            // lbl_new_table_path
            // 
            this.lbl_new_table_path.AutoSize = true;
            this.lbl_new_table_path.Location = new System.Drawing.Point(212, 208);
            this.lbl_new_table_path.Name = "lbl_new_table_path";
            this.lbl_new_table_path.Size = new System.Drawing.Size(118, 24);
            this.lbl_new_table_path.TabIndex = 2;
            this.lbl_new_table_path.Text = "请选择原表格";
            // 
            // btn_start_compare
            // 
            this.btn_start_compare.Location = new System.Drawing.Point(78, 372);
            this.btn_start_compare.Name = "btn_start_compare";
            this.btn_start_compare.Size = new System.Drawing.Size(112, 34);
            this.btn_start_compare.TabIndex = 4;
            this.btn_start_compare.Text = "开始比较";
            this.btn_start_compare.UseVisualStyleBackColor = true;
            this.btn_start_compare.Click += new System.EventHandler(this.btn_start_compare_Click);
            // 
            // TableCompareForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1258, 664);
            this.Controls.Add(this.btn_start_compare);
            this.Controls.Add(this.btn_new_table);
            this.Controls.Add(this.lbl_new_table_path);
            this.Controls.Add(this.btn_ori_table);
            this.Controls.Add(this.lbl_ori_table_path);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "TableCompareForm";
            this.Text = "对比表格差异";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Label lbl_ori_table_path;
        private Button btn_ori_table;
        private Button btn_new_table;
        private Label lbl_new_table_path;
        private Button btn_start_compare;
    }
}