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
            this.rbAnalyzeChangedPages = new System.Windows.Forms.RadioButton();
            this.lblInfo = new System.Windows.Forms.Label();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.lbLog = new System.Windows.Forms.ListBox();
            this.lblProgress = new System.Windows.Forms.Label();
            this.llblDetails = new System.Windows.Forms.LinkLabel();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiSeelctNotebooks = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkForce
            // 
            this.chkForce.AutoSize = true;
            this.chkForce.Location = new System.Drawing.Point(57, 109);
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
            this.rbAnalyzeCurrentPage.Location = new System.Drawing.Point(15, 31);
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
            this.rbAnalyzeAllPages.Location = new System.Drawing.Point(15, 77);
            this.rbAnalyzeAllPages.Name = "rbAnalyzeAllPages";
            this.rbAnalyzeAllPages.Size = new System.Drawing.Size(176, 17);
            this.rbAnalyzeAllPages.TabIndex = 2;
            this.rbAnalyzeAllPages.TabStop = true;
            this.rbAnalyzeAllPages.Text = "Анализировать все страницы";
            this.rbAnalyzeAllPages.UseVisualStyleBackColor = true;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(425, 127);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 3;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // rbAnalyzeChangedPages
            // 
            this.rbAnalyzeChangedPages.AutoSize = true;
            this.rbAnalyzeChangedPages.Location = new System.Drawing.Point(15, 54);
            this.rbAnalyzeChangedPages.Name = "rbAnalyzeChangedPages";
            this.rbAnalyzeChangedPages.Size = new System.Drawing.Size(260, 17);
            this.rbAnalyzeChangedPages.TabIndex = 6;
            this.rbAnalyzeChangedPages.TabStop = true;
            this.rbAnalyzeChangedPages.Text = "Анализировать только измененные страницы";
            this.rbAnalyzeChangedPages.UseVisualStyleBackColor = true;
            // 
            // lblInfo
            // 
            this.lblInfo.AutoSize = true;
            this.lblInfo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblInfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Underline, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.lblInfo.ForeColor = System.Drawing.Color.Red;
            this.lblInfo.Location = new System.Drawing.Point(9, 127);
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Size = new System.Drawing.Size(368, 26);
            this.lblInfo.TabIndex = 7;
            this.lblInfo.Text = "Доступна новая версия программы на сайте http://IStudyBibleTools.ru. \r\nКликните, " +
    "чтобы перейти на страницу загрузки.";
            this.lblInfo.Click += new System.EventHandler(this.lblInfo_Click);
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(12, 177);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(488, 23);
            this.pbMain.Step = 1;
            this.pbMain.TabIndex = 9;
            // 
            // lbLog
            // 
            this.lbLog.FormattingEnabled = true;
            this.lbLog.Location = new System.Drawing.Point(12, 228);
            this.lbLog.Name = "lbLog";
            this.lbLog.Size = new System.Drawing.Size(488, 199);
            this.lbLog.TabIndex = 10;
            // 
            // lblProgress
            // 
            this.lblProgress.AutoSize = true;
            this.lblProgress.Location = new System.Drawing.Point(9, 161);
            this.lblProgress.Name = "lblProgress";
            this.lblProgress.Size = new System.Drawing.Size(96, 13);
            this.lblProgress.TabIndex = 12;
            this.lblProgress.Text = "Инициализация...";
            // 
            // llblDetails
            // 
            this.llblDetails.AutoSize = true;
            this.llblDetails.Location = new System.Drawing.Point(9, 203);
            this.llblDetails.Name = "llblDetails";
            this.llblDetails.Size = new System.Drawing.Size(94, 13);
            this.llblDetails.TabIndex = 14;
            this.llblDetails.TabStop = true;
            this.llblDetails.Text = "Показать детали";
            this.llblDetails.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblDetails_LinkClicked);
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(514, 24);
            this.menuStrip1.TabIndex = 15;
            this.menuStrip1.Text = "Файл";
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiSeelctNotebooks});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(48, 20);
            this.toolStripMenuItem1.Text = "Меню";
            // 
            // tsmiSeelctNotebooks
            // 
            this.tsmiSeelctNotebooks.Name = "tsmiSeelctNotebooks";
            this.tsmiSeelctNotebooks.Size = new System.Drawing.Size(210, 22);
            this.tsmiSeelctNotebooks.Text = "Выбрать записные книжки";
            this.tsmiSeelctNotebooks.Click += new System.EventHandler(this.tsmiSeelctNotebooks_Click);
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(514, 440);
            this.Controls.Add(this.llblDetails);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.lbLog);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.rbAnalyzeChangedPages);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.rbAnalyzeAllPages);
            this.Controls.Add(this.rbAnalyzeCurrentPage);
            this.Controls.Add(this.chkForce);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Укажите параметры";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkForce;
        private System.Windows.Forms.RadioButton rbAnalyzeCurrentPage;
        private System.Windows.Forms.RadioButton rbAnalyzeAllPages;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.RadioButton rbAnalyzeChangedPages;
        private System.Windows.Forms.Label lblInfo;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.ListBox lbLog;
        private System.Windows.Forms.Label lblProgress;
        private System.Windows.Forms.LinkLabel llblDetails;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem tsmiSeelctNotebooks;
    }
}

