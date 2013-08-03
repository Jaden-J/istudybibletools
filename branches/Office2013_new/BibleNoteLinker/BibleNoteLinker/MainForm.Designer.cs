namespace BibleNoteLinker
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
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
            this.pbBaseElements = new System.Windows.Forms.Panel();
            this.cbCurrent = new System.Windows.Forms.ComboBox();
            this.llblShowErrors = new System.Windows.Forms.LinkLabel();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.menuStrip1.SuspendLayout();
            this.pbBaseElements.SuspendLayout();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // chkForce
            // 
            resources.ApplyResources(this.chkForce, "chkForce");
            this.chkForce.Name = "chkForce";
            this.chkForce.UseVisualStyleBackColor = true;
            // 
            // rbAnalyzeCurrentPage
            // 
            resources.ApplyResources(this.rbAnalyzeCurrentPage, "rbAnalyzeCurrentPage");
            this.rbAnalyzeCurrentPage.Checked = true;
            this.rbAnalyzeCurrentPage.Name = "rbAnalyzeCurrentPage";
            this.rbAnalyzeCurrentPage.TabStop = true;
            this.rbAnalyzeCurrentPage.UseVisualStyleBackColor = true;
            this.rbAnalyzeCurrentPage.CheckedChanged += new System.EventHandler(this.rbAnalyzeCurrentPage_CheckedChanged);
            // 
            // rbAnalyzeAllPages
            // 
            resources.ApplyResources(this.rbAnalyzeAllPages, "rbAnalyzeAllPages");
            this.rbAnalyzeAllPages.Name = "rbAnalyzeAllPages";
            this.rbAnalyzeAllPages.TabStop = true;
            this.rbAnalyzeAllPages.UseVisualStyleBackColor = true;
            this.rbAnalyzeAllPages.CheckedChanged += new System.EventHandler(this.rbAnalyzeAllPages_CheckedChanged);
            // 
            // btnOk
            // 
            resources.ApplyResources(this.btnOk, "btnOk");
            this.btnOk.Name = "btnOk";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // rbAnalyzeChangedPages
            // 
            resources.ApplyResources(this.rbAnalyzeChangedPages, "rbAnalyzeChangedPages");
            this.rbAnalyzeChangedPages.Name = "rbAnalyzeChangedPages";
            this.rbAnalyzeChangedPages.TabStop = true;
            this.rbAnalyzeChangedPages.UseVisualStyleBackColor = true;
            this.rbAnalyzeChangedPages.CheckedChanged += new System.EventHandler(this.rbAnalyzeChangedPages_CheckedChanged);
            // 
            // lblInfo
            // 
            resources.ApplyResources(this.lblInfo, "lblInfo");
            this.lblInfo.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblInfo.ForeColor = System.Drawing.Color.Red;
            this.lblInfo.Name = "lblInfo";
            this.lblInfo.Click += new System.EventHandler(this.lblInfo_Click);
            // 
            // pbMain
            // 
            resources.ApplyResources(this.pbMain, "pbMain");
            this.pbMain.Name = "pbMain";
            this.pbMain.Step = 1;
            // 
            // lbLog
            // 
            resources.ApplyResources(this.lbLog, "lbLog");
            this.lbLog.FormattingEnabled = true;
            this.lbLog.Name = "lbLog";
            // 
            // lblProgress
            // 
            resources.ApplyResources(this.lblProgress, "lblProgress");
            this.lblProgress.Name = "lblProgress";
            // 
            // llblDetails
            // 
            resources.ApplyResources(this.llblDetails, "llblDetails");
            this.llblDetails.Name = "llblDetails";
            this.llblDetails.TabStop = true;
            this.llblDetails.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblDetails_LinkClicked);
            // 
            // menuStrip1
            // 
            resources.ApplyResources(this.menuStrip1, "menuStrip1");
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem1});
            this.menuStrip1.Name = "menuStrip1";
            // 
            // toolStripMenuItem1
            // 
            resources.ApplyResources(this.toolStripMenuItem1, "toolStripMenuItem1");
            this.toolStripMenuItem1.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiSeelctNotebooks});
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            // 
            // tsmiSeelctNotebooks
            // 
            resources.ApplyResources(this.tsmiSeelctNotebooks, "tsmiSeelctNotebooks");
            this.tsmiSeelctNotebooks.Name = "tsmiSeelctNotebooks";
            this.tsmiSeelctNotebooks.Click += new System.EventHandler(this.tsmiSeelctNotebooks_Click);
            // 
            // pbBaseElements
            // 
            resources.ApplyResources(this.pbBaseElements, "pbBaseElements");
            this.pbBaseElements.Controls.Add(this.cbCurrent);
            this.pbBaseElements.Controls.Add(this.rbAnalyzeCurrentPage);
            this.pbBaseElements.Controls.Add(this.chkForce);
            this.pbBaseElements.Controls.Add(this.rbAnalyzeAllPages);
            this.pbBaseElements.Controls.Add(this.rbAnalyzeChangedPages);
            this.pbBaseElements.Name = "pbBaseElements";
            // 
            // cbCurrent
            // 
            resources.ApplyResources(this.cbCurrent, "cbCurrent");
            this.cbCurrent.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbCurrent.FormattingEnabled = true;
            this.cbCurrent.Items.AddRange(new object[] {
            resources.GetString("cbCurrent.Items"),
            resources.GetString("cbCurrent.Items1"),
            resources.GetString("cbCurrent.Items2"),
            resources.GetString("cbCurrent.Items3")});
            this.cbCurrent.Name = "cbCurrent";
            // 
            // llblShowErrors
            // 
            resources.ApplyResources(this.llblShowErrors, "llblShowErrors");
            this.llblShowErrors.Name = "llblShowErrors";
            this.llblShowErrors.TabStop = true;
            this.llblShowErrors.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.llblShowErrors_LinkClicked);
            // 
            // flowLayoutPanel1
            // 
            resources.ApplyResources(this.flowLayoutPanel1, "flowLayoutPanel1");
            this.flowLayoutPanel1.Controls.Add(this.llblShowErrors);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOk;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.pbBaseElements);
            this.Controls.Add(this.llblDetails);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblInfo);
            this.Controls.Add(this.lblProgress);
            this.Controls.Add(this.lbLog);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.menuStrip1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MainForm";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.MainForm_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.pbBaseElements.ResumeLayout(false);
            this.pbBaseElements.PerformLayout();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
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
        private System.Windows.Forms.Panel pbBaseElements;
        private System.Windows.Forms.LinkLabel llblShowErrors;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.ComboBox cbCurrent;
    }
}

