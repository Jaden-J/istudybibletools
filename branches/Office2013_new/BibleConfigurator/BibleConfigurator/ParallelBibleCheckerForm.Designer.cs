namespace BibleConfigurator
{
    partial class ParallelBibleCheckerForm
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
            this.cbBaseModule = new System.Windows.Forms.ComboBox();
            this.cbParallelModule = new System.Windows.Forms.ComboBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.rbCheckOneModule = new System.Windows.Forms.RadioButton();
            this.rbCheckAllModules = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.chkWithAllModules = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // cbBaseModule
            // 
            this.cbBaseModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBaseModule.FormattingEnabled = true;
            this.cbBaseModule.Location = new System.Drawing.Point(176, 12);
            this.cbBaseModule.Name = "cbBaseModule";
            this.cbBaseModule.Size = new System.Drawing.Size(121, 21);
            this.cbBaseModule.TabIndex = 2;
            // 
            // cbParallelModule
            // 
            this.cbParallelModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbParallelModule.FormattingEnabled = true;
            this.cbParallelModule.Location = new System.Drawing.Point(176, 39);
            this.cbParallelModule.Name = "cbParallelModule";
            this.cbParallelModule.Size = new System.Drawing.Size(121, 21);
            this.cbParallelModule.TabIndex = 3;
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(141, 112);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(222, 112);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 5;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // rbCheckOneModule
            // 
            this.rbCheckOneModule.AutoSize = true;
            this.rbCheckOneModule.Checked = true;
            this.rbCheckOneModule.Location = new System.Drawing.Point(27, 13);
            this.rbCheckOneModule.Name = "rbCheckOneModule";
            this.rbCheckOneModule.Size = new System.Drawing.Size(120, 17);
            this.rbCheckOneModule.TabIndex = 6;
            this.rbCheckOneModule.TabStop = true;
            this.rbCheckOneModule.Text = "Проверить модуль";
            this.rbCheckOneModule.UseVisualStyleBackColor = true;
            // 
            // rbCheckAllModules
            // 
            this.rbCheckAllModules.AutoSize = true;
            this.rbCheckAllModules.Location = new System.Drawing.Point(27, 89);
            this.rbCheckAllModules.Name = "rbCheckAllModules";
            this.rbCheckAllModules.Size = new System.Drawing.Size(214, 17);
            this.rbCheckAllModules.TabIndex = 7;
            this.rbCheckAllModules.Text = "Проверить все модули друг с другом";
            this.rbCheckAllModules.UseVisualStyleBackColor = true;
            this.rbCheckAllModules.CheckedChanged += new System.EventHandler(this.rbCheckAllModules_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(86, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(61, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "с модулем";
            // 
            // chkWithAllModules
            // 
            this.chkWithAllModules.AutoSize = true;
            this.chkWithAllModules.Location = new System.Drawing.Point(89, 66);
            this.chkWithAllModules.Name = "chkWithAllModules";
            this.chkWithAllModules.Size = new System.Drawing.Size(185, 17);
            this.chkWithAllModules.TabIndex = 9;
            this.chkWithAllModules.Text = "Проверить со всеми модулями";
            this.chkWithAllModules.UseVisualStyleBackColor = true;
            this.chkWithAllModules.CheckedChanged += new System.EventHandler(this.chkWithAllModules_CheckedChanged);
            // 
            // ParallelBibleCheckerForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(311, 148);
            this.Controls.Add(this.chkWithAllModules);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.rbCheckAllModules);
            this.Controls.Add(this.rbCheckOneModule);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.cbParallelModule);
            this.Controls.Add(this.cbBaseModule);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ParallelBibleCheckerForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Проверка параллельных переводов Библии";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ParallelBibleCheckerForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ParallelBibleCheckerForm_FormClosed);
            this.Load += new System.EventHandler(this.ParallelBibleChecker_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cbBaseModule;
        private System.Windows.Forms.ComboBox cbParallelModule;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.RadioButton rbCheckOneModule;
        private System.Windows.Forms.RadioButton rbCheckAllModules;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkWithAllModules;
    }
}