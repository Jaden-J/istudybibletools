﻿namespace BibleNoteLinker
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
            this.SuspendLayout();
            // 
            // lbErrors
            // 
            resources.ApplyResources(this.lbErrors, "lbErrors");
            this.lbErrors.FormattingEnabled = true;
            this.lbErrors.Name = "lbErrors";
            // 
            // btnOk
            // 
            resources.ApplyResources(this.btnOk, "btnOk");
            this.btnOk.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOk.Name = "btnOk";
            this.btnOk.UseVisualStyleBackColor = true;
            // 
            // ErrorsForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lbErrors);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ErrorsForm";
            this.ShowInTaskbar = false;
            this.TopMost = true;
            this.Load += new System.EventHandler(this.Errors_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListBox lbErrors;
        private System.Windows.Forms.Button btnOk;
    }
}