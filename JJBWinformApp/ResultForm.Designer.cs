namespace JJBWinformApp
{
    partial class ResultForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ResultForm));
            this.txt_check_result = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // txt_check_result
            // 
            this.txt_check_result.Location = new System.Drawing.Point(12, 12);
            this.txt_check_result.Multiline = true;
            this.txt_check_result.Name = "txt_check_result";
            this.txt_check_result.ReadOnly = true;
            this.txt_check_result.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txt_check_result.Size = new System.Drawing.Size(776, 426);
            this.txt_check_result.TabIndex = 0;
            // 
            // ResultForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.txt_check_result);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "ResultForm";
            this.Text = "检查结果";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private TextBox txt_check_result;
    }
}