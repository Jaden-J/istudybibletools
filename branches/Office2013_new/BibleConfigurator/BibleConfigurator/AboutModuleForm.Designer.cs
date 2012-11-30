namespace BibleConfigurator
{
    partial class AboutModuleForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutModuleForm));
            this.btnOK = new System.Windows.Forms.Button();
            this.lblTitle = new System.Windows.Forms.Label();
            this.pnBooks = new System.Windows.Forms.Panel();
            this.label1 = new System.Windows.Forms.Label();
            this.lblLocation = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.wbBooks = new System.Windows.Forms.WebBrowser();
            this.pnBooks.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // lblTitle
            // 
            resources.ApplyResources(this.lblTitle, "lblTitle");
            this.lblTitle.Name = "lblTitle";
            // 
            // pnBooks
            // 
            resources.ApplyResources(this.pnBooks, "pnBooks");
            this.pnBooks.BackColor = System.Drawing.SystemColors.Control;
            this.pnBooks.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.pnBooks.Controls.Add(this.wbBooks);
            this.pnBooks.Name = "pnBooks";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // lblLocation
            // 
            resources.ApplyResources(this.lblLocation, "lblLocation");
            this.lblLocation.Name = "lblLocation";
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            // 
            // wbBooks
            // 
            resources.ApplyResources(this.wbBooks, "wbBooks");
            this.wbBooks.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbBooks.Name = "wbBooks";
            // 
            // AboutModuleForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            resources.ApplyResources(this, "$this");
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblLocation);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pnBooks);
            this.Controls.Add(this.lblTitle);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutModuleForm";
            this.Load += new System.EventHandler(this.AboutModule_Load);
            this.Shown += new System.EventHandler(this.AboutModuleForm_Shown);
            this.pnBooks.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label lblTitle;
        private System.Windows.Forms.Panel pnBooks;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblLocation;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.WebBrowser wbBooks;
    }
}