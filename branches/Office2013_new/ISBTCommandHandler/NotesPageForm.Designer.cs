using ISBTCommandHandler.Controls;
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NotesPageForm));
            this.scMain = new System.Windows.Forms.SplitContainer();
            this.wbNotesPage = new ISBTCommandHandler.Controls.WebBrowserEx();
            this.label1 = new System.Windows.Forms.Label();
            this.trbScale = new System.Windows.Forms.TrackBar();
            this.btnClose = new System.Windows.Forms.Button();
            this.chkAlwaysOnTop = new System.Windows.Forms.CheckBox();
            this.chkCloseOnClick = new System.Windows.Forms.CheckBox();
            this.scMain.Panel1.SuspendLayout();
            this.scMain.Panel2.SuspendLayout();
            this.scMain.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trbScale)).BeginInit();
            this.SuspendLayout();
            // 
            // scMain
            // 
            resources.ApplyResources(this.scMain, "scMain");
            this.scMain.FixedPanel = System.Windows.Forms.FixedPanel.Panel2;
            this.scMain.Name = "scMain";
            // 
            // scMain.Panel1
            // 
            this.scMain.Panel1.Controls.Add(this.wbNotesPage);
            // 
            // scMain.Panel2
            // 
            this.scMain.Panel2.Controls.Add(this.label1);
            this.scMain.Panel2.Controls.Add(this.trbScale);
            this.scMain.Panel2.Controls.Add(this.btnClose);
            this.scMain.Panel2.Controls.Add(this.chkAlwaysOnTop);
            this.scMain.Panel2.Controls.Add(this.chkCloseOnClick);
            // 
            // wbNotesPage
            // 
            resources.ApplyResources(this.wbNotesPage, "wbNotesPage");
            this.wbNotesPage.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbNotesPage.Name = "wbNotesPage";
            this.wbNotesPage.DocumentCompleted += new System.Windows.Forms.WebBrowserDocumentCompletedEventHandler(this.wbNotesPage_DocumentCompleted);
            this.wbNotesPage.Navigating += new System.Windows.Forms.WebBrowserNavigatingEventHandler(this.wbNotesPage_Navigating);
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // trbScale
            // 
            resources.ApplyResources(this.trbScale, "trbScale");
            this.trbScale.Maximum = 20;
            this.trbScale.Name = "trbScale";
            this.trbScale.Value = 10;
            this.trbScale.Scroll += new System.EventHandler(this.trbScale_Scroll);
            // 
            // btnClose
            // 
            resources.ApplyResources(this.btnClose, "btnClose");
            this.btnClose.Name = "btnClose";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // chkAlwaysOnTop
            // 
            resources.ApplyResources(this.chkAlwaysOnTop, "chkAlwaysOnTop");
            this.chkAlwaysOnTop.Name = "chkAlwaysOnTop";
            this.chkAlwaysOnTop.UseVisualStyleBackColor = true;
            this.chkAlwaysOnTop.CheckedChanged += new System.EventHandler(this.chkAlwaysOnTop_CheckedChanged);
            // 
            // chkCloseOnClick
            // 
            resources.ApplyResources(this.chkCloseOnClick, "chkCloseOnClick");
            this.chkCloseOnClick.Name = "chkCloseOnClick";
            this.chkCloseOnClick.UseVisualStyleBackColor = true;
            // 
            // NotesPageForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.scMain);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
            this.Name = "NotesPageForm";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NotesPageForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NotesPageForm_FormClosed);
            this.Load += new System.EventHandler(this.NotesPageForm_Load);
            this.scMain.Panel1.ResumeLayout(false);
            this.scMain.Panel2.ResumeLayout(false);
            this.scMain.Panel2.PerformLayout();
            this.scMain.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.trbScale)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer scMain;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.CheckBox chkAlwaysOnTop;
        private System.Windows.Forms.CheckBox chkCloseOnClick;
        private WebBrowserEx wbNotesPage;
        private System.Windows.Forms.TrackBar trbScale;
        private System.Windows.Forms.Label label1;
    }
}