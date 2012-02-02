namespace BibleNoteLinkerEx
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
            this.chkForce = new System.Windows.Forms.CheckBox();
            this.rbAnalyzeCurrentPage = new System.Windows.Forms.RadioButton();
            this.rbAnalyzeAllPages = new System.Windows.Forms.RadioButton();
            this.btnOk = new System.Windows.Forms.Button();
            this.chkDeleteNotes = new System.Windows.Forms.CheckBox();
            this.rbAnalyzeChangedPages = new System.Windows.Forms.RadioButton();
            this.SuspendLayout();
            // 
            // chkForce
            // 
            this.chkForce.AutoSize = true;
            this.chkForce.Location = new System.Drawing.Point(41, 90);
            this.chkForce.Name = "chkForce";
            this.chkForce.Size = new System.Drawing.Size(189, 17);
            this.chkForce.TabIndex = 0;
            this.chkForce.Text = "Повторный анализ всех ссылок";
            this.chkForce.UseVisualStyleBackColor = true;
            // 
            // rbAnalyzeCurrentPage
            // 
            this.rbAnalyzeCurrentPage.AutoSize = true;
            this.rbAnalyzeCurrentPage.Checked = true;
            this.rbAnalyzeCurrentPage.Location = new System.Drawing.Point(12, 12);
            this.rbAnalyzeCurrentPage.Name = "rbAnalyzeCurrentPage";
            this.rbAnalyzeCurrentPage.Size = new System.Drawing.Size(199, 17);
            this.rbAnalyzeCurrentPage.TabIndex = 1;
            this.rbAnalyzeCurrentPage.TabStop = true;
            this.rbAnalyzeCurrentPage.Text = "Анализировать текущую страницу";
            this.rbAnalyzeCurrentPage.UseVisualStyleBackColor = true;
            // 
            // rbAnalyzeAllPages
            // 
            this.rbAnalyzeAllPages.AutoSize = true;
            this.rbAnalyzeAllPages.Location = new System.Drawing.Point(12, 35);
            this.rbAnalyzeAllPages.Name = "rbAnalyzeAllPages";
            this.rbAnalyzeAllPages.Size = new System.Drawing.Size(218, 17);
            this.rbAnalyzeAllPages.TabIndex = 2;
            this.rbAnalyzeAllPages.TabStop = true;
            this.rbAnalyzeAllPages.Text = "Анализировать всю записную книжку";
            this.rbAnalyzeAllPages.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(204, 136);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 3;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // chkDeleteNotes
            // 
            this.chkDeleteNotes.AutoSize = true;
            this.chkDeleteNotes.Location = new System.Drawing.Point(41, 113);
            this.chkDeleteNotes.Name = "chkDeleteNotes";
            this.chkDeleteNotes.Size = new System.Drawing.Size(214, 17);
            this.chkDeleteNotes.TabIndex = 5;
            this.chkDeleteNotes.Text = "Удалить сводные страницы заметок";
            this.chkDeleteNotes.UseVisualStyleBackColor = true;
            this.chkDeleteNotes.CheckedChanged += new System.EventHandler(this.cbDeleteNotes_CheckedChanged);
            // 
            // rbAnalyzeChangedPages
            // 
            this.rbAnalyzeChangedPages.AutoSize = true;
            this.rbAnalyzeChangedPages.Location = new System.Drawing.Point(12, 58);
            this.rbAnalyzeChangedPages.Name = "rbAnalyzeChangedPages";
            this.rbAnalyzeChangedPages.Size = new System.Drawing.Size(222, 17);
            this.rbAnalyzeChangedPages.TabIndex = 6;
            this.rbAnalyzeChangedPages.TabStop = true;
            this.rbAnalyzeChangedPages.Text = "Анализировать изменённые страницы";
            this.rbAnalyzeChangedPages.UseVisualStyleBackColor = true;
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(291, 172);
            this.Controls.Add(this.rbAnalyzeChangedPages);
            this.Controls.Add(this.chkDeleteNotes);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.rbAnalyzeAllPages);
            this.Controls.Add(this.rbAnalyzeCurrentPage);
            this.Controls.Add(this.chkForce);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Укажите параметры";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkForce;
        private System.Windows.Forms.RadioButton rbAnalyzeCurrentPage;
        private System.Windows.Forms.RadioButton rbAnalyzeAllPages;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.CheckBox chkDeleteNotes;
        private System.Windows.Forms.RadioButton rbAnalyzeChangedPages;
    }
}

