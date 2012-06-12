namespace BibleConfigurator
{
    partial class NotebookParametersForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(NotebookParametersForm));
            this.cbBibleSection = new System.Windows.Forms.ComboBox();
            this.cbBibleCommentsSection = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cbBibleStudySection = new System.Windows.Forms.ComboBox();
            this.btnBibleSectionRename = new System.Windows.Forms.Button();
            this.btnBibleCommentsSectionRename = new System.Windows.Forms.Button();
            this.btnBibleStudySectionRename = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // cbBibleSection
            // 
            this.cbBibleSection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleSection.FormattingEnabled = true;
            resources.ApplyResources(this.cbBibleSection, "cbBibleSection");
            this.cbBibleSection.Name = "cbBibleSection";
            // 
            // cbBibleCommentsSection
            // 
            this.cbBibleCommentsSection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleCommentsSection.FormattingEnabled = true;
            resources.ApplyResources(this.cbBibleCommentsSection, "cbBibleCommentsSection");
            this.cbBibleCommentsSection.Name = "cbBibleCommentsSection";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            // 
            // cbBibleStudySection
            // 
            this.cbBibleStudySection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleStudySection.FormattingEnabled = true;
            resources.ApplyResources(this.cbBibleStudySection, "cbBibleStudySection");
            this.cbBibleStudySection.Name = "cbBibleStudySection";
            // 
            // btnBibleSectionRename
            // 
            resources.ApplyResources(this.btnBibleSectionRename, "btnBibleSectionRename");
            this.btnBibleSectionRename.Name = "btnBibleSectionRename";
            this.btnBibleSectionRename.UseVisualStyleBackColor = true;
            this.btnBibleSectionRename.Click += new System.EventHandler(this.btnBibleSectionRename_Click);
            // 
            // btnBibleCommentsSectionRename
            // 
            resources.ApplyResources(this.btnBibleCommentsSectionRename, "btnBibleCommentsSectionRename");
            this.btnBibleCommentsSectionRename.Name = "btnBibleCommentsSectionRename";
            this.btnBibleCommentsSectionRename.UseVisualStyleBackColor = true;
            this.btnBibleCommentsSectionRename.Click += new System.EventHandler(this.btnBibleCommentsSectionRename_Click);
            // 
            // btnBibleStudySectionRename
            // 
            resources.ApplyResources(this.btnBibleStudySectionRename, "btnBibleStudySectionRename");
            this.btnBibleStudySectionRename.Name = "btnBibleStudySectionRename";
            this.btnBibleStudySectionRename.UseVisualStyleBackColor = true;
            this.btnBibleStudySectionRename.Click += new System.EventHandler(this.btnBibleStudySectionRename_Click);
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // NotebookParametersForm
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.btnBibleStudySectionRename);
            this.Controls.Add(this.btnBibleCommentsSectionRename);
            this.Controls.Add(this.btnBibleSectionRename);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbBibleStudySection);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbBibleCommentsSection);
            this.Controls.Add(this.cbBibleSection);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "NotebookParametersForm";
            this.ShowInTaskbar = false;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.NotebookParametersForm_FormClosed);
            this.Load += new System.EventHandler(this.NotebookParametersForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.NotebookParametersForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbBibleSection;
        private System.Windows.Forms.ComboBox cbBibleCommentsSection;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cbBibleStudySection;
        private System.Windows.Forms.Button btnBibleSectionRename;
        private System.Windows.Forms.Button btnBibleCommentsSectionRename;
        private System.Windows.Forms.Button btnBibleStudySectionRename;
        private System.Windows.Forms.Button btnOK;

    }
}