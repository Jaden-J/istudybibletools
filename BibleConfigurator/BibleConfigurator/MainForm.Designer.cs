namespace BibleConfigurator
{
    partial class MainForm
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
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.tbcMain = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.btnBibleStudyNotebookSetPath = new System.Windows.Forms.Button();
            this.btnBibleCommentsNotebookSetPath = new System.Windows.Forms.Button();
            this.btnBibleNotebookSetPath = new System.Windows.Forms.Button();
            this.btnSingleNotebookSetPath = new System.Windows.Forms.Button();
            this.btnSingleNotebookParameters = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.chkCreateBibleStudyNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbBibleStudyNotebook = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.chkCreateBibleCommentsNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbBibleCommentsNotebook = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.chkCreateBibleNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbBibleNotebook = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkCreateSingleNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbSingleNotebook = new System.Windows.Forms.ComboBox();
            this.rbMultiNotebook = new System.Windows.Forms.RadioButton();
            this.rbSingleNotebook = new System.Windows.Forms.RadioButton();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chkDefaultPageNameParameters = new System.Windows.Forms.CheckBox();
            this.tbNotesPageName = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbBookOverviewName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbPageDescriptionName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.button1 = new System.Windows.Forms.Button();
            this.btnResizeBibleTables = new System.Windows.Forms.Button();
            this.btnRelinkComments = new System.Windows.Forms.Button();
            this.btnOK = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lblProgressInfo = new System.Windows.Forms.Label();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.tbcMain.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tbcMain
            // 
            this.tbcMain.Controls.Add(this.tabPage1);
            this.tbcMain.Controls.Add(this.tabPage2);
            this.tbcMain.Controls.Add(this.tabPage3);
            this.tbcMain.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbcMain.Location = new System.Drawing.Point(0, 0);
            this.tbcMain.Name = "tbcMain";
            this.tbcMain.SelectedIndex = 0;
            this.tbcMain.Size = new System.Drawing.Size(490, 310);
            this.tbcMain.TabIndex = 16;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage1.Controls.Add(this.btnBibleStudyNotebookSetPath);
            this.tabPage1.Controls.Add(this.btnBibleCommentsNotebookSetPath);
            this.tabPage1.Controls.Add(this.btnBibleNotebookSetPath);
            this.tabPage1.Controls.Add(this.btnSingleNotebookSetPath);
            this.tabPage1.Controls.Add(this.btnSingleNotebookParameters);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.chkCreateBibleStudyNotebookFromTemplate);
            this.tabPage1.Controls.Add(this.cbBibleStudyNotebook);
            this.tabPage1.Controls.Add(this.label3);
            this.tabPage1.Controls.Add(this.chkCreateBibleCommentsNotebookFromTemplate);
            this.tabPage1.Controls.Add(this.cbBibleCommentsNotebook);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.chkCreateBibleNotebookFromTemplate);
            this.tabPage1.Controls.Add(this.cbBibleNotebook);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.chkCreateSingleNotebookFromTemplate);
            this.tabPage1.Controls.Add(this.cbSingleNotebook);
            this.tabPage1.Controls.Add(this.rbMultiNotebook);
            this.tabPage1.Controls.Add(this.rbSingleNotebook);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(482, 284);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Параметры OneNote";
            // 
            // btnBibleStudyNotebookSetPath
            // 
            this.btnBibleStudyNotebookSetPath.Location = new System.Drawing.Point(437, 241);
            this.btnBibleStudyNotebookSetPath.Name = "btnBibleStudyNotebookSetPath";
            this.btnBibleStudyNotebookSetPath.Size = new System.Drawing.Size(26, 23);
            this.btnBibleStudyNotebookSetPath.TabIndex = 35;
            this.btnBibleStudyNotebookSetPath.Text = "...";
            this.btnBibleStudyNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleStudyNotebookSetPath.Click += new System.EventHandler(this.btnBibleStudyNotebookSetPath_Click);
            // 
            // btnBibleCommentsNotebookSetPath
            // 
            this.btnBibleCommentsNotebookSetPath.Location = new System.Drawing.Point(437, 191);
            this.btnBibleCommentsNotebookSetPath.Name = "btnBibleCommentsNotebookSetPath";
            this.btnBibleCommentsNotebookSetPath.Size = new System.Drawing.Size(26, 23);
            this.btnBibleCommentsNotebookSetPath.TabIndex = 34;
            this.btnBibleCommentsNotebookSetPath.Text = "...";
            this.btnBibleCommentsNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleCommentsNotebookSetPath.Click += new System.EventHandler(this.btnBibleCommentsNotebookSetPath_Click);
            // 
            // btnBibleNotebookSetPath
            // 
            this.btnBibleNotebookSetPath.Location = new System.Drawing.Point(437, 141);
            this.btnBibleNotebookSetPath.Name = "btnBibleNotebookSetPath";
            this.btnBibleNotebookSetPath.Size = new System.Drawing.Size(26, 23);
            this.btnBibleNotebookSetPath.TabIndex = 33;
            this.btnBibleNotebookSetPath.Text = "...";
            this.btnBibleNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleNotebookSetPath.Click += new System.EventHandler(this.btnBibleNotebookSetPath_Click);
            // 
            // btnSingleNotebookSetPath
            // 
            this.btnSingleNotebookSetPath.Location = new System.Drawing.Point(437, 40);
            this.btnSingleNotebookSetPath.Name = "btnSingleNotebookSetPath";
            this.btnSingleNotebookSetPath.Size = new System.Drawing.Size(26, 23);
            this.btnSingleNotebookSetPath.TabIndex = 32;
            this.btnSingleNotebookSetPath.Text = "...";
            this.btnSingleNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnSingleNotebookSetPath.Click += new System.EventHandler(this.btnSingleNotebookSetPath_Click);
            // 
            // btnSingleNotebookParameters
            // 
            this.btnSingleNotebookParameters.Location = new System.Drawing.Point(47, 69);
            this.btnSingleNotebookParameters.Name = "btnSingleNotebookParameters";
            this.btnSingleNotebookParameters.Size = new System.Drawing.Size(102, 23);
            this.btnSingleNotebookParameters.TabIndex = 31;
            this.btnSingleNotebookParameters.Text = "Настроить";
            this.btnSingleNotebookParameters.UseVisualStyleBackColor = true;
            this.btnSingleNotebookParameters.Click += new System.EventHandler(this.btnSingleNotebookParameters_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(46, 227);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(259, 13);
            this.label4.TabIndex = 30;
            this.label4.Text = "Выберите записную книжку для изучения Библии";
            // 
            // chkCreateBibleStudyNotebookFromTemplate
            // 
            this.chkCreateBibleStudyNotebookFromTemplate.AutoSize = true;
            this.chkCreateBibleStudyNotebookFromTemplate.Location = new System.Drawing.Point(303, 245);
            this.chkCreateBibleStudyNotebookFromTemplate.Name = "chkCreateBibleStudyNotebookFromTemplate";
            this.chkCreateBibleStudyNotebookFromTemplate.Size = new System.Drawing.Size(130, 17);
            this.chkCreateBibleStudyNotebookFromTemplate.TabIndex = 29;
            this.chkCreateBibleStudyNotebookFromTemplate.Text = "Создать из шаблона";
            this.chkCreateBibleStudyNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleStudyNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleStudyNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleStudyNotebook
            // 
            this.cbBibleStudyNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleStudyNotebook.FormattingEnabled = true;
            this.cbBibleStudyNotebook.Location = new System.Drawing.Point(49, 243);
            this.cbBibleStudyNotebook.Name = "cbBibleStudyNotebook";
            this.cbBibleStudyNotebook.Size = new System.Drawing.Size(248, 21);
            this.cbBibleStudyNotebook.TabIndex = 28;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(46, 177);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(297, 13);
            this.label3.TabIndex = 27;
            this.label3.Text = "Выберите записную книжку для комментариев к Библии";
            // 
            // chkCreateBibleCommentsNotebookFromTemplate
            // 
            this.chkCreateBibleCommentsNotebookFromTemplate.AutoSize = true;
            this.chkCreateBibleCommentsNotebookFromTemplate.Location = new System.Drawing.Point(303, 195);
            this.chkCreateBibleCommentsNotebookFromTemplate.Name = "chkCreateBibleCommentsNotebookFromTemplate";
            this.chkCreateBibleCommentsNotebookFromTemplate.Size = new System.Drawing.Size(130, 17);
            this.chkCreateBibleCommentsNotebookFromTemplate.TabIndex = 26;
            this.chkCreateBibleCommentsNotebookFromTemplate.Text = "Создать из шаблона";
            this.chkCreateBibleCommentsNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleCommentsNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleCommentsNotebook
            // 
            this.cbBibleCommentsNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleCommentsNotebook.FormattingEnabled = true;
            this.cbBibleCommentsNotebook.Location = new System.Drawing.Point(49, 193);
            this.cbBibleCommentsNotebook.Name = "cbBibleCommentsNotebook";
            this.cbBibleCommentsNotebook.Size = new System.Drawing.Size(248, 21);
            this.cbBibleCommentsNotebook.TabIndex = 25;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(46, 127);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(210, 13);
            this.label2.TabIndex = 24;
            this.label2.Text = "Выберите записную книжку для Библии";
            // 
            // chkCreateBibleNotebookFromTemplate
            // 
            this.chkCreateBibleNotebookFromTemplate.AutoSize = true;
            this.chkCreateBibleNotebookFromTemplate.Location = new System.Drawing.Point(303, 145);
            this.chkCreateBibleNotebookFromTemplate.Name = "chkCreateBibleNotebookFromTemplate";
            this.chkCreateBibleNotebookFromTemplate.Size = new System.Drawing.Size(130, 17);
            this.chkCreateBibleNotebookFromTemplate.TabIndex = 23;
            this.chkCreateBibleNotebookFromTemplate.Text = "Создать из шаблона";
            this.chkCreateBibleNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleNotebook
            // 
            this.cbBibleNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleNotebook.FormattingEnabled = true;
            this.cbBibleNotebook.Location = new System.Drawing.Point(49, 143);
            this.cbBibleNotebook.Name = "cbBibleNotebook";
            this.cbBibleNotebook.Size = new System.Drawing.Size(248, 21);
            this.cbBibleNotebook.TabIndex = 22;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(44, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(149, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Выберите записную книжку";
            // 
            // chkCreateSingleNotebookFromTemplate
            // 
            this.chkCreateSingleNotebookFromTemplate.AutoSize = true;
            this.chkCreateSingleNotebookFromTemplate.Location = new System.Drawing.Point(301, 44);
            this.chkCreateSingleNotebookFromTemplate.Name = "chkCreateSingleNotebookFromTemplate";
            this.chkCreateSingleNotebookFromTemplate.Size = new System.Drawing.Size(130, 17);
            this.chkCreateSingleNotebookFromTemplate.TabIndex = 20;
            this.chkCreateSingleNotebookFromTemplate.Text = "Создать из шаблона";
            this.chkCreateSingleNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateSingleNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateSingleNotebookFromTemplate_CheckedChanged);
            // 
            // cbSingleNotebook
            // 
            this.cbSingleNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSingleNotebook.FormattingEnabled = true;
            this.cbSingleNotebook.Location = new System.Drawing.Point(47, 42);
            this.cbSingleNotebook.Name = "cbSingleNotebook";
            this.cbSingleNotebook.Size = new System.Drawing.Size(248, 21);
            this.cbSingleNotebook.TabIndex = 19;
            // 
            // rbMultiNotebook
            // 
            this.rbMultiNotebook.AutoSize = true;
            this.rbMultiNotebook.Checked = true;
            this.rbMultiNotebook.Location = new System.Drawing.Point(8, 107);
            this.rbMultiNotebook.Name = "rbMultiNotebook";
            this.rbMultiNotebook.Size = new System.Drawing.Size(176, 17);
            this.rbMultiNotebook.TabIndex = 18;
            this.rbMultiNotebook.TabStop = true;
            this.rbMultiNotebook.Text = "Отдельные записные книжки";
            this.rbMultiNotebook.UseVisualStyleBackColor = true;
            this.rbMultiNotebook.CheckedChanged += new System.EventHandler(this.rbMultiNotebook_CheckedChanged);
            // 
            // rbSingleNotebook
            // 
            this.rbSingleNotebook.AutoSize = true;
            this.rbSingleNotebook.Location = new System.Drawing.Point(6, 6);
            this.rbSingleNotebook.Name = "rbSingleNotebook";
            this.rbSingleNotebook.Size = new System.Drawing.Size(143, 17);
            this.rbSingleNotebook.TabIndex = 17;
            this.rbSingleNotebook.Text = "Одна записная книжка";
            this.rbSingleNotebook.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.chkDefaultPageNameParameters);
            this.tabPage2.Controls.Add(this.tbNotesPageName);
            this.tabPage2.Controls.Add(this.label7);
            this.tabPage2.Controls.Add(this.tbBookOverviewName);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.tbPageDescriptionName);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(482, 284);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Параметры программы";
            // 
            // chkDefaultPageNameParameters
            // 
            this.chkDefaultPageNameParameters.AutoSize = true;
            this.chkDefaultPageNameParameters.Location = new System.Drawing.Point(16, 171);
            this.chkDefaultPageNameParameters.Name = "chkDefaultPageNameParameters";
            this.chkDefaultPageNameParameters.Size = new System.Drawing.Size(223, 17);
            this.chkDefaultPageNameParameters.TabIndex = 6;
            this.chkDefaultPageNameParameters.Text = "Использовать значения по умолчанию";
            this.chkDefaultPageNameParameters.UseVisualStyleBackColor = true;
            this.chkDefaultPageNameParameters.CheckedChanged += new System.EventHandler(this.chkDefaultPageNameParameters_CheckedChanged);
            // 
            // tbNotesPageName
            // 
            this.tbNotesPageName.Location = new System.Drawing.Point(16, 129);
            this.tbNotesPageName.Name = "tbNotesPageName";
            this.tbNotesPageName.Size = new System.Drawing.Size(248, 20);
            this.tbNotesPageName.TabIndex = 5;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(13, 113);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(211, 13);
            this.label7.TabIndex = 4;
            this.label7.Text = "Название страницы \"Сводная заметок\"";
            // 
            // tbBookOverviewName
            // 
            this.tbBookOverviewName.Location = new System.Drawing.Point(16, 79);
            this.tbBookOverviewName.Name = "tbBookOverviewName";
            this.tbBookOverviewName.Size = new System.Drawing.Size(248, 20);
            this.tbBookOverviewName.TabIndex = 3;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(13, 63);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(302, 13);
            this.label6.TabIndex = 2;
            this.label6.Text = "Название страницы комментариев к книге по умолчанию";
            // 
            // tbPageDescriptionName
            // 
            this.tbPageDescriptionName.Location = new System.Drawing.Point(16, 29);
            this.tbPageDescriptionName.Name = "tbPageDescriptionName";
            this.tbPageDescriptionName.Size = new System.Drawing.Size(248, 20);
            this.tbPageDescriptionName.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 13);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(302, 13);
            this.label5.TabIndex = 0;
            this.label5.Text = "Название страницы комментариев к главе по умолчанию";
            // 
            // tabPage3
            // 
            this.tabPage3.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage3.Controls.Add(this.button1);
            this.tabPage3.Controls.Add(this.btnResizeBibleTables);
            this.tabPage3.Controls.Add(this.btnRelinkComments);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Size = new System.Drawing.Size(482, 284);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Дополнительные утилиты";
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(13, 93);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(230, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "Создать резервную копию данных";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // btnResizeBibleTables
            // 
            this.btnResizeBibleTables.Enabled = false;
            this.btnResizeBibleTables.Location = new System.Drawing.Point(13, 53);
            this.btnResizeBibleTables.Name = "btnResizeBibleTables";
            this.btnResizeBibleTables.Size = new System.Drawing.Size(230, 23);
            this.btnResizeBibleTables.TabIndex = 1;
            this.btnResizeBibleTables.Text = "Изменить ширину страниц Библии";
            this.btnResizeBibleTables.UseVisualStyleBackColor = true;
            // 
            // btnRelinkComments
            // 
            this.btnRelinkComments.Location = new System.Drawing.Point(13, 13);
            this.btnRelinkComments.Name = "btnRelinkComments";
            this.btnRelinkComments.Size = new System.Drawing.Size(230, 23);
            this.btnRelinkComments.TabIndex = 0;
            this.btnRelinkComments.Text = "Обновить ссылки на комментарии";
            this.btnRelinkComments.UseVisualStyleBackColor = true;
            this.btnRelinkComments.Click += new System.EventHandler(this.btnRelinkComments_Click);
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(392, 16);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 16;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tbcMain);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.lblProgressInfo);
            this.splitContainer1.Panel2.Controls.Add(this.pbMain);
            this.splitContainer1.Panel2.Controls.Add(this.btnOK);
            this.splitContainer1.Size = new System.Drawing.Size(490, 365);
            this.splitContainer1.SplitterDistance = 310;
            this.splitContainer1.TabIndex = 17;
            // 
            // lblProgressInfo
            // 
            this.lblProgressInfo.AutoSize = true;
            this.lblProgressInfo.Location = new System.Drawing.Point(12, 0);
            this.lblProgressInfo.Name = "lblProgressInfo";
            this.lblProgressInfo.Size = new System.Drawing.Size(0, 13);
            this.lblProgressInfo.TabIndex = 18;
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(12, 16);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(374, 23);
            this.pbMain.Step = 3;
            this.pbMain.TabIndex = 17;
            this.pbMain.Visible = false;
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(490, 365);
            this.Controls.Add(this.splitContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Укажите параметры программы OneNote IStudyBibleTools";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.tbcMain.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.TabControl tbcMain;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.Button btnSingleNotebookParameters;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckBox chkCreateBibleStudyNotebookFromTemplate;
        private System.Windows.Forms.ComboBox cbBibleStudyNotebook;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox chkCreateBibleCommentsNotebookFromTemplate;
        private System.Windows.Forms.ComboBox cbBibleCommentsNotebook;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox chkCreateBibleNotebookFromTemplate;
        private System.Windows.Forms.ComboBox cbBibleNotebook;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkCreateSingleNotebookFromTemplate;
        private System.Windows.Forms.ComboBox cbSingleNotebook;
        private System.Windows.Forms.RadioButton rbMultiNotebook;
        private System.Windows.Forms.RadioButton rbSingleNotebook;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox tbNotesPageName;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox tbBookOverviewName;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox tbPageDescriptionName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button btnResizeBibleTables;
        private System.Windows.Forms.Button btnRelinkComments;
        private System.Windows.Forms.Button btnBibleStudyNotebookSetPath;
        private System.Windows.Forms.Button btnBibleCommentsNotebookSetPath;
        private System.Windows.Forms.Button btnBibleNotebookSetPath;
        private System.Windows.Forms.Button btnSingleNotebookSetPath;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox chkDefaultPageNameParameters;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.Label lblProgressInfo;

    }
}

