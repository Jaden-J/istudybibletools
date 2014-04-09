namespace BibleCommon.UI.Forms
{
    partial class ErrorsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorsForm));
            this.lbErrors = new System.Windows.Forms.ListBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnSaveToFile = new System.Windows.Forms.Button();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.btnOpenLog = new System.Windows.Forms.Button();
            this.lblDescription = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbErrors
            // 
            this.lbErrors.FormattingEnabled = true;
            resources.ApplyResources(this.lbErrors, "lbErrors");
            this.lbErrors.Name = "lbErrors";
            this.lbErrors.MouseClick += new System.Windows.Forms.MouseEventHandler(this.lbErrors_MouseClick);
            this.lbErrors.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lbErrors_KeyDown);
            this.lbErrors.MouseMove += new System.Windows.Forms.MouseEventHandler(this.lbErrors_MouseMove);
            // 
            // btnOk
            // 
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.btnOk, "btnOk");
            this.btnOk.Name = "btnOk";
            this.btnOk.UseVisualStyleBackColor = true;
            // 
            // btnSaveToFile
            // 
            resources.ApplyResources(this.btnSaveToFile, "btnSaveToFile");
            this.btnSaveToFile.Name = "btnSaveToFile";
            this.btnSaveToFile.UseVisualStyleBackColor = true;
            this.btnSaveToFile.Click += new System.EventHandler(this.btnSaveToFile_Click);
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.DefaultExt = "txt";
            this.saveFileDialog.FileName = "errors.txt";
            resources.ApplyResources(this.saveFileDialog, "saveFileDialog");
            // 
            // btnOpenLog
            // 
            resources.ApplyResources(this.btnOpenLog, "btnOpenLog");
            this.btnOpenLog.Name = "btnOpenLog";
            this.btnOpenLog.UseVisualStyleBackColor = true;
            this.btnOpenLog.Click += new System.EventHandler(this.btnOpenLog_Click);
            // 
            // lblDescription
            // 
            resources.ApplyResources(this.lblDescription, "lblDescription");
            this.lblDescription.Name = "lblDescription";
            // 
            // ErrorsForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.lblDescription);
            this.Controls.Add(this.btnOpenLog);
            this.Controls.Add(this.btnSaveToFile);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lbErrors);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ErrorsForm";
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ErrorsForm_FormClosed);
            this.Load += new System.EventHandler(this.Errors_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lbErrors;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnSaveToFile;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Button btnOpenLog;
        private System.Windows.Forms.Label lblDescription;
    }
}