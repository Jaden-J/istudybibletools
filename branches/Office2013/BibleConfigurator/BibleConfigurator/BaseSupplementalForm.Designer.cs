namespace BibleConfigurator
{
    abstract partial class BaseSupplementalForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(BaseSupplementalForm));
            this.chkUseSupplementalBible = new System.Windows.Forms.CheckBox();
            this.pnModules = new System.Windows.Forms.Panel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnSBFolder = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.SuspendLayout();
            // 
            // chkUseSupplementalBible
            // 
            resources.ApplyResources(this.chkUseSupplementalBible, "chkUseSupplementalBible");
            this.chkUseSupplementalBible.Name = "chkUseSupplementalBible";
            this.chkUseSupplementalBible.UseVisualStyleBackColor = true;
            this.chkUseSupplementalBible.CheckedChanged += new System.EventHandler(this.chkUseSupplementalBible_CheckedChanged);
            // 
            // pnModules
            // 
            resources.ApplyResources(this.pnModules, "pnModules");
            this.pnModules.Name = "pnModules";
            // 
            // btnOk
            // 
            resources.ApplyResources(this.btnOk, "btnOk");
            this.btnOk.Name = "btnOk";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnSBFolder
            // 
            resources.ApplyResources(this.btnSBFolder, "btnSBFolder");
            this.btnSBFolder.Name = "btnSBFolder";
            this.btnSBFolder.UseVisualStyleBackColor = true;
            this.btnSBFolder.Click += new System.EventHandler(this.btnSBFolder_Click);
            // 
            // BaseSupplementalForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.btnSBFolder);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.pnModules);
            this.Controls.Add(this.chkUseSupplementalBible);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "BaseSupplementalForm";
            this.ShowInTaskbar = false;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SupplementalBibleForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SupplementalBibleForm_FormClosed);
            this.Load += new System.EventHandler(this.SupplementalBibleForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SupplementalBibleForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkUseSupplementalBible;
        private System.Windows.Forms.Panel pnModules;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnSBFolder;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
    }
}