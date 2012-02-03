namespace BibleVerseLinkerEx
{
    partial class MainForm
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
            this.tbPageName = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.cbSearchForUnderlineText = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // tbPageName
            // 
            this.tbPageName.Location = new System.Drawing.Point(13, 12);
            this.tbPageName.Name = "tbPageName";
            this.tbPageName.Size = new System.Drawing.Size(244, 20);
            this.tbPageName.TabIndex = 0;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(182, 68);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // cbSearchForUnderlineText
            // 
            this.cbSearchForUnderlineText.AutoSize = true;
            this.cbSearchForUnderlineText.Location = new System.Drawing.Point(13, 38);
            this.cbSearchForUnderlineText.Name = "cbSearchForUnderlineText";
            this.cbSearchForUnderlineText.Size = new System.Drawing.Size(168, 17);
            this.cbSearchForUnderlineText.TabIndex = 2;
            this.cbSearchForUnderlineText.Text = "Искать подчеркнутый текст";
            this.cbSearchForUnderlineText.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(269, 103);
            this.Controls.Add(this.cbSearchForUnderlineText);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.tbPageName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Укажите имя страницы";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbPageName;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.CheckBox cbSearchForUnderlineText;
    }
}

