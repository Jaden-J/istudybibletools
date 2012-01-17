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
            this.cbBibleSection.Location = new System.Drawing.Point(16, 29);
            this.cbBibleSection.Name = "cbBibleSection";
            this.cbBibleSection.Size = new System.Drawing.Size(248, 21);
            this.cbBibleSection.TabIndex = 4;
            // 
            // cbBibleCommentsSection
            // 
            this.cbBibleCommentsSection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleCommentsSection.FormattingEnabled = true;
            this.cbBibleCommentsSection.Location = new System.Drawing.Point(16, 79);
            this.cbBibleCommentsSection.Name = "cbBibleCommentsSection";
            this.cbBibleCommentsSection.Size = new System.Drawing.Size(248, 21);
            this.cbBibleCommentsSection.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(191, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Веберите группу секций для Библии";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 63);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(280, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Выберите группу секций для комментариев к Библии";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 113);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(242, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Выберите группу секций для изучения Библии";
            // 
            // cbBibleStudySection
            // 
            this.cbBibleStudySection.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleStudySection.FormattingEnabled = true;
            this.cbBibleStudySection.Location = new System.Drawing.Point(16, 129);
            this.cbBibleStudySection.Name = "cbBibleStudySection";
            this.cbBibleStudySection.Size = new System.Drawing.Size(248, 21);
            this.cbBibleStudySection.TabIndex = 8;
            // 
            // btnBibleSectionRename
            // 
            this.btnBibleSectionRename.Location = new System.Drawing.Point(270, 27);
            this.btnBibleSectionRename.Name = "btnBibleSectionRename";
            this.btnBibleSectionRename.Size = new System.Drawing.Size(125, 23);
            this.btnBibleSectionRename.TabIndex = 10;
            this.btnBibleSectionRename.Text = "Переименовать";
            this.btnBibleSectionRename.UseVisualStyleBackColor = true;
            this.btnBibleSectionRename.Click += new System.EventHandler(this.btnBibleSectionRename_Click);
            // 
            // btnBibleCommentsSectionRename
            // 
            this.btnBibleCommentsSectionRename.Location = new System.Drawing.Point(270, 77);
            this.btnBibleCommentsSectionRename.Name = "btnBibleCommentsSectionRename";
            this.btnBibleCommentsSectionRename.Size = new System.Drawing.Size(125, 23);
            this.btnBibleCommentsSectionRename.TabIndex = 11;
            this.btnBibleCommentsSectionRename.Text = "Переименовать";
            this.btnBibleCommentsSectionRename.UseVisualStyleBackColor = true;
            this.btnBibleCommentsSectionRename.Click += new System.EventHandler(this.btnBibleCommentsSectionRename_Click);
            // 
            // btnBibleStudySectionRename
            // 
            this.btnBibleStudySectionRename.Location = new System.Drawing.Point(270, 127);
            this.btnBibleStudySectionRename.Name = "btnBibleStudySectionRename";
            this.btnBibleStudySectionRename.Size = new System.Drawing.Size(125, 23);
            this.btnBibleStudySectionRename.TabIndex = 12;
            this.btnBibleStudySectionRename.Text = "Переименовать";
            this.btnBibleStudySectionRename.UseVisualStyleBackColor = true;
            this.btnBibleStudySectionRename.Click += new System.EventHandler(this.btnBibleStudySectionRename_Click);
            // 
            // btnOK
            // 
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(320, 177);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 13;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // NotebookParametersForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(409, 212);
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
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "NotebookParameters";
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