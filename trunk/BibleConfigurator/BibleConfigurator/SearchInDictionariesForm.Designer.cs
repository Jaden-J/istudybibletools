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
            this.cbAllTerms = new System.Windows.Forms.ComboBox();
            this.cbFoundInDictionaries = new System.Windows.Forms.ComboBox();
            this.lblFoundInDictionaries = new System.Windows.Forms.Label();
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
            // 
            // cbAllTerms
            // 
            this.cbAllTerms.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple;
            this.cbAllTerms.FormattingEnabled = true;
            resources.ApplyResources(this.cbAllTerms, "cbAllTerms");
            this.cbAllTerms.Name = "cbAllTerms";
            this.cbAllTerms.SelectedIndexChanged += new System.EventHandler(this.cbAllTerms_SelectedIndexChanged);
            // 
            // cbFoundInDictionaries
            // 
            this.cbFoundInDictionaries.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbFoundInDictionaries.FormattingEnabled = true;
            resources.ApplyResources(this.cbFoundInDictionaries, "cbFoundInDictionaries");
            this.cbFoundInDictionaries.Name = "cbFoundInDictionaries";
            // 
            // lblFoundInDictionaries
            // 
            resources.ApplyResources(this.lblFoundInDictionaries, "lblFoundInDictionaries");
            this.lblFoundInDictionaries.Name = "lblFoundInDictionaries";
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
            // SearchInDictionariesForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lblFoundInDictionaries);
            this.Controls.Add(this.cbFoundInDictionaries);
            this.Controls.Add(this.cbAllTerms);
            this.Controls.Add(this.cbDictionaries);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SearchInDictionariesForm";
            this.Load += new System.EventHandler(this.SearchInDictionariesForm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbDictionaries;
        private System.Windows.Forms.ComboBox cbAllTerms;
        private System.Windows.Forms.ComboBox cbFoundInDictionaries;
        private System.Windows.Forms.Label lblFoundInDictionaries;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnCancel;
    }
}