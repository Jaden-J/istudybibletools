namespace ISBTCommandHandler
{
    partial class NotesPageForm
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
            this.scMain = new System.Windows.Forms.SplitContainer();
            this.wbNotesPage = new System.Windows.Forms.WebBrowser();
            this.btnClose = new System.Windows.Forms.Button();
            this.chkAlwaysOnTop = new System.Windows.Forms.CheckBox();
            this.chkCloseOnClick = new System.Windows.Forms.CheckBox();
            this.scMain.Panel1.SuspendLayout();
            this.scMain.Panel2.SuspendLayout();
            this.scMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // scMain
            // 
            this.scMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.scMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scMain.Location = new System.Drawing.Point(0, 0);
            this.scMain.Name = "scMain";
            this.scMain.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // scMain.Panel1
            // 
            this.scMain.Panel1.Controls.Add(this.wbNotesPage);
            // 
            // scMain.Panel2
            // 
            this.scMain.Panel2.Controls.Add(this.btnClose);
            this.scMain.Panel2.Controls.Add(this.chkAlwaysOnTop);
            this.scMain.Panel2.Controls.Add(this.chkCloseOnClick);
            this.scMain.Size = new System.Drawing.Size(242, 226);
            this.scMain.SplitterDistance = 172;
            this.scMain.TabIndex = 0;
            // 
            // wbNotesPage
            // 
            this.wbNotesPage.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wbNotesPage.Location = new System.Drawing.Point(0, 0);
            this.wbNotesPage.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbNotesPage.Name = "wbNotesPage";
            this.wbNotesPage.Size = new System.Drawing.Size(242, 172);
            this.wbNotesPage.TabIndex = 0;
            this.wbNotesPage.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbNotesPage_Navigating);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnClose.Location = new System.Drawing.Point(155, 16);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 2;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // chkAlwaysOnTop
            // 
            this.chkAlwaysOnTop.AutoSize = true;
            this.chkAlwaysOnTop.Location = new System.Drawing.Point(12, 26);
            this.chkAlwaysOnTop.Name = "chkAlwaysOnTop";
            this.chkAlwaysOnTop.Size = new System.Drawing.Size(96, 17);
            this.chkAlwaysOnTop.TabIndex = 1;
            this.chkAlwaysOnTop.Text = "Always on Top";
            this.chkAlwaysOnTop.UseVisualStyleBackColor = true;
            this.chkAlwaysOnTop.CheckedChanged += new System.EventHandler(this.chkAlwaysOnTop_CheckedChanged);
            // 
            // chkCloseOnClick
            // 
            this.chkCloseOnClick.AutoSize = true;
            this.chkCloseOnClick.Location = new System.Drawing.Point(12, 3);
            this.chkCloseOnClick.Name = "chkCloseOnClick";
            this.chkCloseOnClick.Size = new System.Drawing.Size(93, 17);
            this.chkCloseOnClick.TabIndex = 0;
            this.chkCloseOnClick.Text = "Close on Click";
            this.chkCloseOnClick.UseVisualStyleBackColor = true;
            // 
            // NotesPageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(242, 226);
            this.Controls.Add(this.scMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.MinimumSize = new System.Drawing.Size(250, 250);
            this.Name = "NotesPageForm";
            this.Text = "NotesPage Form";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NotesPageForm_FormClosed);
            this.Load += new System.EventHandler(this.NotesPageForm_Load);            
            this.scMain.Panel1.ResumeLayout(false);
            this.scMain.Panel2.ResumeLayout(false);
            this.scMain.Panel2.PerformLayout();
            this.scMain.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer scMain;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckBox chkAlwaysOnTop;
        private System.Windows.Forms.CheckBox chkCloseOnClick;
        private System.Windows.Forms.WebBrowser wbNotesPage;
    }
}