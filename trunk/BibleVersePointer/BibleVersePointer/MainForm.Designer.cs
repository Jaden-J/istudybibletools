namespace BibleVersePointer
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
            this.tbVerse = new System.Windows.Forms.TextBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbVerse
            // 
            this.tbVerse.Location = new System.Drawing.Point(13, 12);
            this.tbVerse.Name = "tbVerse";
            this.tbVerse.Size = new System.Drawing.Size(244, 20);
            this.tbVerse.TabIndex = 0;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(182, 42);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "Go";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(269, 77);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.tbVerse);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Укажите место Писания";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbVerse;
        private System.Windows.Forms.Button btnOk;
    }
}

