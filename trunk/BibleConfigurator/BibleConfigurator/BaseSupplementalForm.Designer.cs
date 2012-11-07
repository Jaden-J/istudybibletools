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
            this.btnSBFolder = new System.Windows.Forms.Button();
            this.cbExistingNotebooks = new System.Windows.Forms.ComboBox();
            this.rbUseExisting = new System.Windows.Forms.RadioButton();
            this.rbCreateNew = new System.Windows.Forms.RadioButton();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.pnModules.SuspendLayout();
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
            this.pnModules.Controls.Add(this.btnSBFolder);
            this.pnModules.Controls.Add(this.cbExistingNotebooks);
            this.pnModules.Controls.Add(this.rbUseExisting);
            this.pnModules.Controls.Add(this.rbCreateNew);
            this.pnModules.Name = "pnModules";
            // 
            // btnSBFolder
            // 
            resources.ApplyResources(this.btnSBFolder, "btnSBFolder");
            this.btnSBFolder.Name = "btnSBFolder";
            this.btnSBFolder.UseVisualStyleBackColor = true;
            this.btnSBFolder.Click += new System.EventHandler(this.btnSBFolder_Click);
            // 
            // cbExistingNotebooks
            // 
            resources.ApplyResources(this.cbExistingNotebooks, "cbExistingNotebooks");
            this.cbExistingNotebooks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbExistingNotebooks.FormattingEnabled = true;
            this.cbExistingNotebooks.Name = "cbExistingNotebooks";
            // 
            // rbUseExisting
            // 
            resources.ApplyResources(this.rbUseExisting, "rbUseExisting");
            this.rbUseExisting.Name = "rbUseExisting";
            this.rbUseExisting.TabStop = true;
            this.rbUseExisting.UseVisualStyleBackColor = true;
            this.rbUseExisting.CheckedChanged += new System.EventHandler(this.rbUseExisting_CheckedChanged);
            // 
            // rbCreateNew
            // 
            resources.ApplyResources(this.rbCreateNew, "rbCreateNew");
            this.rbCreateNew.Checked = true;
            this.rbCreateNew.Name = "rbCreateNew";
            this.rbCreateNew.TabStop = true;
            this.rbCreateNew.UseVisualStyleBackColor = true;
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
            // folderBrowserDialog
            // 
            resources.ApplyResources(this.folderBrowserDialog, "folderBrowserDialog");
            // 
            // BaseSupplementalForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
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
            this.pnModules.ResumeLayout(false);
            this.pnModules.PerformLayout();
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
        private System.Windows.Forms.ComboBox cbExistingNotebooks;
        private System.Windows.Forms.RadioButton rbUseExisting;
        private System.Windows.Forms.RadioButton rbCreateNew;
    }
}