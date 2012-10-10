namespace BibleConfigurator
{
    partial class SearchInDictionariesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SearchInDictionariesForm));
            this.cbDictionaries = new System.Windows.Forms.ComboBox();
            this.cbTerms = new System.Windows.Forms.ComboBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbDictionaries
            // 
            this.cbDictionaries.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbDictionaries.FormattingEnabled = true;
            resources.ApplyResources(this.cbDictionaries, "cbDictionaries");
            this.cbDictionaries.Name = "cbDictionaries";
            this.cbDictionaries.SelectedIndexChanged += new System.EventHandler(this.cbDictionaries_SelectedIndexChanged);
            // 
            // cbTerms
            // 
            this.cbTerms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.cbTerms.FormattingEnabled = true;
            resources.ApplyResources(this.cbTerms, "cbTerms");
            this.cbTerms.Name = "cbTerms";
            this.cbTerms.MouseClick += new System.Windows.Forms.MouseEventHandler(this.cbTerms_MouseClick);
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
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // SearchInDictionariesForm
            // 
            this.AcceptButton = this.btnOk;
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.CancelButton = this.btnCancel;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.cbTerms);
            this.Controls.Add(this.cbDictionaries);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SearchInDictionariesForm";
            this.TopMost = true;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SearchInDictionariesForm_FormClosed);
            this.Load += new System.EventHandler(this.SearchInDictionariesForm_Load);
            this.Shown += new System.EventHandler(this.SearchInDictionariesForm_Shown);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SearchInDictionariesForm_KeyDown);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cbDictionaries;
        private System.Windows.Forms.ComboBox cbTerms;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
    }
}