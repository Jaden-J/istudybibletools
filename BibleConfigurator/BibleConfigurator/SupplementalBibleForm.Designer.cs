﻿namespace BibleConfigurator
{
    partial class SupplementalBibleForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SupplementalBibleForm));
            this.chkUseSupplementalBible = new System.Windows.Forms.CheckBox();
            this.pnModules = new System.Windows.Forms.Panel();
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
            // SupplementalBibleForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.pnModules);
            this.Controls.Add(this.chkUseSupplementalBible);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.KeyPreview = true;
            this.Name = "SupplementalBibleForm";
            this.ShowInTaskbar = false;
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SupplementalBibleForm_FormClosed);
            this.Load += new System.EventHandler(this.SupplementalBibleForm_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SupplementalBibleForm_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckBox chkUseSupplementalBible;
        private System.Windows.Forms.Panel pnModules;
    }
}