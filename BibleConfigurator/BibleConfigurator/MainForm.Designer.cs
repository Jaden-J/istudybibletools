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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tbcMain = new System.Windows.Forms.TabControl();
            this.tpNotebooks = new System.Windows.Forms.TabPage();
            this.chkUseFolderForBibleNotesPages = new System.Windows.Forms.CheckBox();
            this.btnBibleNotesPagesSetFolder = new System.Windows.Forms.Button();
            this.tbBibleNotesPagesFolder = new System.Windows.Forms.TextBox();
            this.btnBibleNotesPagesNotebookSetPath = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.chkCreateBibleNotesPagesNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbBibleNotesPagesNotebook = new System.Windows.Forms.ComboBox();
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
            this.lblSelectSingleNotebook = new System.Windows.Forms.Label();
            this.chkCreateSingleNotebookFromTemplate = new System.Windows.Forms.CheckBox();
            this.cbSingleNotebook = new System.Windows.Forms.ComboBox();
            this.rbMultiNotebook = new System.Windows.Forms.RadioButton();
            this.rbSingleNotebook = new System.Windows.Forms.RadioButton();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chkUseCommaDelimeter = new System.Windows.Forms.CheckBox();
            this.chkUseAdvancedProxyForOneNoteLinks = new System.Windows.Forms.CheckBox();
            this.chkUseProxyLinksForBibleVerses = new System.Windows.Forms.CheckBox();
            this.chkUseProxyLinksForLinks = new System.Windows.Forms.CheckBox();
            this.chkUseProxyLinksForStrong = new System.Windows.Forms.CheckBox();
            this.label9 = new System.Windows.Forms.Label();
            this.cbLanguage = new System.Windows.Forms.ComboBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.chkUseDifferentPages = new System.Windows.Forms.CheckBox();
            this.label7 = new System.Windows.Forms.Label();
            this.tbNotesPageName = new System.Windows.Forms.TextBox();
            this.chkExcludedVersesLinking = new System.Windows.Forms.CheckBox();
            this.tbNotesPageWidth = new System.Windows.Forms.TextBox();
            this.chkExpandMultiVersesLinking = new System.Windows.Forms.CheckBox();
            this.lblNotesPageWidth = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkUseRubbishPage = new System.Windows.Forms.CheckBox();
            this.label10 = new System.Windows.Forms.Label();
            this.chkRubbishExcludedVersesLinking = new System.Windows.Forms.CheckBox();
            this.tbRubbishNotesPageName = new System.Windows.Forms.TextBox();
            this.chkRubbishExpandMultiVersesLinking = new System.Windows.Forms.CheckBox();
            this.tbRubbishNotesPageWidth = new System.Windows.Forms.TextBox();
            this.lblRubbishNotesPageWidth = new System.Windows.Forms.Label();
            this.chkDefaultParameters = new System.Windows.Forms.CheckBox();
            this.tbBookOverviewName = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.tbCommentsPageName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnModuleChecker = new System.Windows.Forms.Button();
            this.btnConverter = new System.Windows.Forms.Button();
            this.btnBackup = new System.Windows.Forms.Button();
            this.btnResizeBibleTables = new System.Windows.Forms.Button();
            this.btnDeleteNotesPages = new System.Windows.Forms.Button();
            this.btnRelinkComments = new System.Windows.Forms.Button();
            this.tpModules = new System.Windows.Forms.TabPage();
            this.btnDictionariesManagement = new System.Windows.Forms.Button();
            this.btnSupplementalBibleManagement = new System.Windows.Forms.Button();
            this.hlModules = new System.Windows.Forms.LinkLabel();
            this.lblModulesLink = new System.Windows.Forms.Label();
            this.pnModules = new System.Windows.Forms.Panel();
            this.lblMustSelectModule = new System.Windows.Forms.Label();
            this.lblMustUploadModule = new System.Windows.Forms.Label();
            this.btnUploadModule = new System.Windows.Forms.Button();
            this.btnApply = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.lblWarning = new System.Windows.Forms.Label();
            this.lblProgressInfo = new System.Windows.Forms.Label();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.btnOK = new System.Windows.Forms.Button();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.openModuleFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.notesPagesFolderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.ttNotesPageFolder = new System.Windows.Forms.ToolTip(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tbcMain.SuspendLayout();
            this.tpNotebooks.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tpModules.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            resources.ApplyResources(this.splitContainer1, "splitContainer1");
            this.splitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            resources.ApplyResources(this.splitContainer1.Panel1, "splitContainer1.Panel1");
            this.splitContainer1.Panel1.Controls.Add(this.tbcMain);
            this.ttNotesPageFolder.SetToolTip(this.splitContainer1.Panel1, resources.GetString("splitContainer1.Panel1.ToolTip"));
            // 
            // splitContainer1.Panel2
            // 
            resources.ApplyResources(this.splitContainer1.Panel2, "splitContainer1.Panel2");
            this.splitContainer1.Panel2.Controls.Add(this.btnApply);
            this.splitContainer1.Panel2.Controls.Add(this.btnClose);
            this.splitContainer1.Panel2.Controls.Add(this.lblWarning);
            this.splitContainer1.Panel2.Controls.Add(this.lblProgressInfo);
            this.splitContainer1.Panel2.Controls.Add(this.pbMain);
            this.splitContainer1.Panel2.Controls.Add(this.btnOK);
            this.ttNotesPageFolder.SetToolTip(this.splitContainer1.Panel2, resources.GetString("splitContainer1.Panel2.ToolTip"));
            this.ttNotesPageFolder.SetToolTip(this.splitContainer1, resources.GetString("splitContainer1.ToolTip"));
            // 
            // tbcMain
            // 
            resources.ApplyResources(this.tbcMain, "tbcMain");
            this.tbcMain.Controls.Add(this.tpNotebooks);
            this.tbcMain.Controls.Add(this.tabPage2);
            this.tbcMain.Controls.Add(this.tabPage3);
            this.tbcMain.Controls.Add(this.tpModules);
            this.tbcMain.Name = "tbcMain";
            this.tbcMain.SelectedIndex = 0;
            this.ttNotesPageFolder.SetToolTip(this.tbcMain, resources.GetString("tbcMain.ToolTip"));
            // 
            // tpNotebooks
            // 
            resources.ApplyResources(this.tpNotebooks, "tpNotebooks");
            this.tpNotebooks.BackColor = System.Drawing.SystemColors.Control;
            this.tpNotebooks.Controls.Add(this.chkUseFolderForBibleNotesPages);
            this.tpNotebooks.Controls.Add(this.btnBibleNotesPagesSetFolder);
            this.tpNotebooks.Controls.Add(this.tbBibleNotesPagesFolder);
            this.tpNotebooks.Controls.Add(this.btnBibleNotesPagesNotebookSetPath);
            this.tpNotebooks.Controls.Add(this.label8);
            this.tpNotebooks.Controls.Add(this.chkCreateBibleNotesPagesNotebookFromTemplate);
            this.tpNotebooks.Controls.Add(this.cbBibleNotesPagesNotebook);
            this.tpNotebooks.Controls.Add(this.btnBibleStudyNotebookSetPath);
            this.tpNotebooks.Controls.Add(this.btnBibleCommentsNotebookSetPath);
            this.tpNotebooks.Controls.Add(this.btnBibleNotebookSetPath);
            this.tpNotebooks.Controls.Add(this.btnSingleNotebookSetPath);
            this.tpNotebooks.Controls.Add(this.btnSingleNotebookParameters);
            this.tpNotebooks.Controls.Add(this.label4);
            this.tpNotebooks.Controls.Add(this.chkCreateBibleStudyNotebookFromTemplate);
            this.tpNotebooks.Controls.Add(this.cbBibleStudyNotebook);
            this.tpNotebooks.Controls.Add(this.label3);
            this.tpNotebooks.Controls.Add(this.chkCreateBibleCommentsNotebookFromTemplate);
            this.tpNotebooks.Controls.Add(this.cbBibleCommentsNotebook);
            this.tpNotebooks.Controls.Add(this.label2);
            this.tpNotebooks.Controls.Add(this.chkCreateBibleNotebookFromTemplate);
            this.tpNotebooks.Controls.Add(this.cbBibleNotebook);
            this.tpNotebooks.Controls.Add(this.lblSelectSingleNotebook);
            this.tpNotebooks.Controls.Add(this.chkCreateSingleNotebookFromTemplate);
            this.tpNotebooks.Controls.Add(this.cbSingleNotebook);
            this.tpNotebooks.Controls.Add(this.rbMultiNotebook);
            this.tpNotebooks.Controls.Add(this.rbSingleNotebook);
            this.tpNotebooks.Name = "tpNotebooks";
            this.ttNotesPageFolder.SetToolTip(this.tpNotebooks, resources.GetString("tpNotebooks.ToolTip"));
            this.tpNotebooks.Enter += new System.EventHandler(this.tabPage1_Enter);
            // 
            // chkUseFolderForBibleNotesPages
            // 
            resources.ApplyResources(this.chkUseFolderForBibleNotesPages, "chkUseFolderForBibleNotesPages");
            this.chkUseFolderForBibleNotesPages.Checked = true;
            this.chkUseFolderForBibleNotesPages.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseFolderForBibleNotesPages.Name = "chkUseFolderForBibleNotesPages";
            this.ttNotesPageFolder.SetToolTip(this.chkUseFolderForBibleNotesPages, resources.GetString("chkUseFolderForBibleNotesPages.ToolTip"));
            this.chkUseFolderForBibleNotesPages.UseVisualStyleBackColor = true;
            this.chkUseFolderForBibleNotesPages.CheckedChanged += new System.EventHandler(this.chkUseFolderForBibleNotesPages_CheckedChanged);
            // 
            // btnBibleNotesPagesSetFolder
            // 
            resources.ApplyResources(this.btnBibleNotesPagesSetFolder, "btnBibleNotesPagesSetFolder");
            this.btnBibleNotesPagesSetFolder.Name = "btnBibleNotesPagesSetFolder";
            this.ttNotesPageFolder.SetToolTip(this.btnBibleNotesPagesSetFolder, resources.GetString("btnBibleNotesPagesSetFolder.ToolTip"));
            this.btnBibleNotesPagesSetFolder.UseVisualStyleBackColor = true;
            this.btnBibleNotesPagesSetFolder.Click += new System.EventHandler(this.btnBibleNotesPagesSetFolder_Click);
            // 
            // tbBibleNotesPagesFolder
            // 
            resources.ApplyResources(this.tbBibleNotesPagesFolder, "tbBibleNotesPagesFolder");
            this.tbBibleNotesPagesFolder.Name = "tbBibleNotesPagesFolder";
            this.tbBibleNotesPagesFolder.ReadOnly = true;
            this.ttNotesPageFolder.SetToolTip(this.tbBibleNotesPagesFolder, resources.GetString("tbBibleNotesPagesFolder.ToolTip"));
            this.tbBibleNotesPagesFolder.Click += new System.EventHandler(this.tbBibleNotesPagesFolder_Click);
            // 
            // btnBibleNotesPagesNotebookSetPath
            // 
            resources.ApplyResources(this.btnBibleNotesPagesNotebookSetPath, "btnBibleNotesPagesNotebookSetPath");
            this.btnBibleNotesPagesNotebookSetPath.Name = "btnBibleNotesPagesNotebookSetPath";
            this.ttNotesPageFolder.SetToolTip(this.btnBibleNotesPagesNotebookSetPath, resources.GetString("btnBibleNotesPagesNotebookSetPath.ToolTip"));
            this.btnBibleNotesPagesNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleNotesPagesNotebookSetPath.Click += new System.EventHandler(this.btnBibleNotesPagesNotebookSetPath_Click);
            // 
            // label8
            // 
            resources.ApplyResources(this.label8, "label8");
            this.label8.Name = "label8";
            this.ttNotesPageFolder.SetToolTip(this.label8, resources.GetString("label8.ToolTip"));
            // 
            // chkCreateBibleNotesPagesNotebookFromTemplate
            // 
            resources.ApplyResources(this.chkCreateBibleNotesPagesNotebookFromTemplate, "chkCreateBibleNotesPagesNotebookFromTemplate");
            this.chkCreateBibleNotesPagesNotebookFromTemplate.Name = "chkCreateBibleNotesPagesNotebookFromTemplate";
            this.ttNotesPageFolder.SetToolTip(this.chkCreateBibleNotesPagesNotebookFromTemplate, resources.GetString("chkCreateBibleNotesPagesNotebookFromTemplate.ToolTip"));
            this.chkCreateBibleNotesPagesNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleNotesPagesNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleNotesPagesNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleNotesPagesNotebook
            // 
            resources.ApplyResources(this.cbBibleNotesPagesNotebook, "cbBibleNotesPagesNotebook");
            this.cbBibleNotesPagesNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleNotesPagesNotebook.FormattingEnabled = true;
            this.cbBibleNotesPagesNotebook.Name = "cbBibleNotesPagesNotebook";
            this.ttNotesPageFolder.SetToolTip(this.cbBibleNotesPagesNotebook, resources.GetString("cbBibleNotesPagesNotebook.ToolTip"));
            // 
            // btnBibleStudyNotebookSetPath
            // 
            resources.ApplyResources(this.btnBibleStudyNotebookSetPath, "btnBibleStudyNotebookSetPath");
            this.btnBibleStudyNotebookSetPath.Name = "btnBibleStudyNotebookSetPath";
            this.ttNotesPageFolder.SetToolTip(this.btnBibleStudyNotebookSetPath, resources.GetString("btnBibleStudyNotebookSetPath.ToolTip"));
            this.btnBibleStudyNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleStudyNotebookSetPath.Click += new System.EventHandler(this.btnBibleStudyNotebookSetPath_Click);
            // 
            // btnBibleCommentsNotebookSetPath
            // 
            resources.ApplyResources(this.btnBibleCommentsNotebookSetPath, "btnBibleCommentsNotebookSetPath");
            this.btnBibleCommentsNotebookSetPath.Name = "btnBibleCommentsNotebookSetPath";
            this.ttNotesPageFolder.SetToolTip(this.btnBibleCommentsNotebookSetPath, resources.GetString("btnBibleCommentsNotebookSetPath.ToolTip"));
            this.btnBibleCommentsNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleCommentsNotebookSetPath.Click += new System.EventHandler(this.btnBibleCommentsNotebookSetPath_Click);
            // 
            // btnBibleNotebookSetPath
            // 
            resources.ApplyResources(this.btnBibleNotebookSetPath, "btnBibleNotebookSetPath");
            this.btnBibleNotebookSetPath.Name = "btnBibleNotebookSetPath";
            this.ttNotesPageFolder.SetToolTip(this.btnBibleNotebookSetPath, resources.GetString("btnBibleNotebookSetPath.ToolTip"));
            this.btnBibleNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnBibleNotebookSetPath.Click += new System.EventHandler(this.btnBibleNotebookSetPath_Click);
            // 
            // btnSingleNotebookSetPath
            // 
            resources.ApplyResources(this.btnSingleNotebookSetPath, "btnSingleNotebookSetPath");
            this.btnSingleNotebookSetPath.Name = "btnSingleNotebookSetPath";
            this.ttNotesPageFolder.SetToolTip(this.btnSingleNotebookSetPath, resources.GetString("btnSingleNotebookSetPath.ToolTip"));
            this.btnSingleNotebookSetPath.UseVisualStyleBackColor = true;
            this.btnSingleNotebookSetPath.Click += new System.EventHandler(this.btnSingleNotebookSetPath_Click);
            // 
            // btnSingleNotebookParameters
            // 
            resources.ApplyResources(this.btnSingleNotebookParameters, "btnSingleNotebookParameters");
            this.btnSingleNotebookParameters.Name = "btnSingleNotebookParameters";
            this.ttNotesPageFolder.SetToolTip(this.btnSingleNotebookParameters, resources.GetString("btnSingleNotebookParameters.ToolTip"));
            this.btnSingleNotebookParameters.UseVisualStyleBackColor = true;
            this.btnSingleNotebookParameters.Click += new System.EventHandler(this.btnSingleNotebookParameters_Click);
            // 
            // label4
            // 
            resources.ApplyResources(this.label4, "label4");
            this.label4.Name = "label4";
            this.ttNotesPageFolder.SetToolTip(this.label4, resources.GetString("label4.ToolTip"));
            // 
            // chkCreateBibleStudyNotebookFromTemplate
            // 
            resources.ApplyResources(this.chkCreateBibleStudyNotebookFromTemplate, "chkCreateBibleStudyNotebookFromTemplate");
            this.chkCreateBibleStudyNotebookFromTemplate.Name = "chkCreateBibleStudyNotebookFromTemplate";
            this.ttNotesPageFolder.SetToolTip(this.chkCreateBibleStudyNotebookFromTemplate, resources.GetString("chkCreateBibleStudyNotebookFromTemplate.ToolTip"));
            this.chkCreateBibleStudyNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleStudyNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleStudyNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleStudyNotebook
            // 
            resources.ApplyResources(this.cbBibleStudyNotebook, "cbBibleStudyNotebook");
            this.cbBibleStudyNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleStudyNotebook.FormattingEnabled = true;
            this.cbBibleStudyNotebook.Name = "cbBibleStudyNotebook";
            this.ttNotesPageFolder.SetToolTip(this.cbBibleStudyNotebook, resources.GetString("cbBibleStudyNotebook.ToolTip"));
            // 
            // label3
            // 
            resources.ApplyResources(this.label3, "label3");
            this.label3.Name = "label3";
            this.ttNotesPageFolder.SetToolTip(this.label3, resources.GetString("label3.ToolTip"));
            // 
            // chkCreateBibleCommentsNotebookFromTemplate
            // 
            resources.ApplyResources(this.chkCreateBibleCommentsNotebookFromTemplate, "chkCreateBibleCommentsNotebookFromTemplate");
            this.chkCreateBibleCommentsNotebookFromTemplate.Name = "chkCreateBibleCommentsNotebookFromTemplate";
            this.ttNotesPageFolder.SetToolTip(this.chkCreateBibleCommentsNotebookFromTemplate, resources.GetString("chkCreateBibleCommentsNotebookFromTemplate.ToolTip"));
            this.chkCreateBibleCommentsNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleCommentsNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleCommentsNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleCommentsNotebook
            // 
            resources.ApplyResources(this.cbBibleCommentsNotebook, "cbBibleCommentsNotebook");
            this.cbBibleCommentsNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleCommentsNotebook.FormattingEnabled = true;
            this.cbBibleCommentsNotebook.Name = "cbBibleCommentsNotebook";
            this.ttNotesPageFolder.SetToolTip(this.cbBibleCommentsNotebook, resources.GetString("cbBibleCommentsNotebook.ToolTip"));
            // 
            // label2
            // 
            resources.ApplyResources(this.label2, "label2");
            this.label2.Name = "label2";
            this.ttNotesPageFolder.SetToolTip(this.label2, resources.GetString("label2.ToolTip"));
            // 
            // chkCreateBibleNotebookFromTemplate
            // 
            resources.ApplyResources(this.chkCreateBibleNotebookFromTemplate, "chkCreateBibleNotebookFromTemplate");
            this.chkCreateBibleNotebookFromTemplate.Name = "chkCreateBibleNotebookFromTemplate";
            this.ttNotesPageFolder.SetToolTip(this.chkCreateBibleNotebookFromTemplate, resources.GetString("chkCreateBibleNotebookFromTemplate.ToolTip"));
            this.chkCreateBibleNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateBibleNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateBibleNotebookFromTemplate_CheckedChanged);
            // 
            // cbBibleNotebook
            // 
            resources.ApplyResources(this.cbBibleNotebook, "cbBibleNotebook");
            this.cbBibleNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbBibleNotebook.FormattingEnabled = true;
            this.cbBibleNotebook.Name = "cbBibleNotebook";
            this.ttNotesPageFolder.SetToolTip(this.cbBibleNotebook, resources.GetString("cbBibleNotebook.ToolTip"));
            // 
            // lblSelectSingleNotebook
            // 
            resources.ApplyResources(this.lblSelectSingleNotebook, "lblSelectSingleNotebook");
            this.lblSelectSingleNotebook.Name = "lblSelectSingleNotebook";
            this.ttNotesPageFolder.SetToolTip(this.lblSelectSingleNotebook, resources.GetString("lblSelectSingleNotebook.ToolTip"));
            // 
            // chkCreateSingleNotebookFromTemplate
            // 
            resources.ApplyResources(this.chkCreateSingleNotebookFromTemplate, "chkCreateSingleNotebookFromTemplate");
            this.chkCreateSingleNotebookFromTemplate.Name = "chkCreateSingleNotebookFromTemplate";
            this.ttNotesPageFolder.SetToolTip(this.chkCreateSingleNotebookFromTemplate, resources.GetString("chkCreateSingleNotebookFromTemplate.ToolTip"));
            this.chkCreateSingleNotebookFromTemplate.UseVisualStyleBackColor = true;
            this.chkCreateSingleNotebookFromTemplate.CheckedChanged += new System.EventHandler(this.chkCreateSingleNotebookFromTemplate_CheckedChanged);
            // 
            // cbSingleNotebook
            // 
            resources.ApplyResources(this.cbSingleNotebook, "cbSingleNotebook");
            this.cbSingleNotebook.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbSingleNotebook.FormattingEnabled = true;
            this.cbSingleNotebook.Name = "cbSingleNotebook";
            this.ttNotesPageFolder.SetToolTip(this.cbSingleNotebook, resources.GetString("cbSingleNotebook.ToolTip"));
            // 
            // rbMultiNotebook
            // 
            resources.ApplyResources(this.rbMultiNotebook, "rbMultiNotebook");
            this.rbMultiNotebook.Checked = true;
            this.rbMultiNotebook.Name = "rbMultiNotebook";
            this.rbMultiNotebook.TabStop = true;
            this.ttNotesPageFolder.SetToolTip(this.rbMultiNotebook, resources.GetString("rbMultiNotebook.ToolTip"));
            this.rbMultiNotebook.UseVisualStyleBackColor = true;
            this.rbMultiNotebook.CheckedChanged += new System.EventHandler(this.rbMultiNotebook_CheckedChanged);
            // 
            // rbSingleNotebook
            // 
            resources.ApplyResources(this.rbSingleNotebook, "rbSingleNotebook");
            this.rbSingleNotebook.Name = "rbSingleNotebook";
            this.ttNotesPageFolder.SetToolTip(this.rbSingleNotebook, resources.GetString("rbSingleNotebook.ToolTip"));
            this.rbSingleNotebook.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            resources.ApplyResources(this.tabPage2, "tabPage2");
            this.tabPage2.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage2.Controls.Add(this.chkUseCommaDelimeter);
            this.tabPage2.Controls.Add(this.chkUseAdvancedProxyForOneNoteLinks);
            this.tabPage2.Controls.Add(this.chkUseProxyLinksForBibleVerses);
            this.tabPage2.Controls.Add(this.chkUseProxyLinksForLinks);
            this.tabPage2.Controls.Add(this.chkUseProxyLinksForStrong);
            this.tabPage2.Controls.Add(this.label9);
            this.tabPage2.Controls.Add(this.cbLanguage);
            this.tabPage2.Controls.Add(this.groupBox2);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.chkDefaultParameters);
            this.tabPage2.Controls.Add(this.tbBookOverviewName);
            this.tabPage2.Controls.Add(this.label6);
            this.tabPage2.Controls.Add(this.tbCommentsPageName);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Name = "tabPage2";
            this.ttNotesPageFolder.SetToolTip(this.tabPage2, resources.GetString("tabPage2.ToolTip"));
            this.tabPage2.Enter += new System.EventHandler(this.tabPage1_Enter);
            // 
            // chkUseCommaDelimeter
            // 
            resources.ApplyResources(this.chkUseCommaDelimeter, "chkUseCommaDelimeter");
            this.chkUseCommaDelimeter.Name = "chkUseCommaDelimeter";
            this.ttNotesPageFolder.SetToolTip(this.chkUseCommaDelimeter, resources.GetString("chkUseCommaDelimeter.ToolTip"));
            this.chkUseCommaDelimeter.UseVisualStyleBackColor = true;
            // 
            // chkUseAdvancedProxyForOneNoteLinks
            // 
            resources.ApplyResources(this.chkUseAdvancedProxyForOneNoteLinks, "chkUseAdvancedProxyForOneNoteLinks");
            this.chkUseAdvancedProxyForOneNoteLinks.Name = "chkUseAdvancedProxyForOneNoteLinks";
            this.ttNotesPageFolder.SetToolTip(this.chkUseAdvancedProxyForOneNoteLinks, resources.GetString("chkUseAdvancedProxyForOneNoteLinks.ToolTip"));
            this.chkUseAdvancedProxyForOneNoteLinks.UseVisualStyleBackColor = true;
            // 
            // chkUseProxyLinksForBibleVerses
            // 
            resources.ApplyResources(this.chkUseProxyLinksForBibleVerses, "chkUseProxyLinksForBibleVerses");
            this.chkUseProxyLinksForBibleVerses.Name = "chkUseProxyLinksForBibleVerses";
            this.ttNotesPageFolder.SetToolTip(this.chkUseProxyLinksForBibleVerses, resources.GetString("chkUseProxyLinksForBibleVerses.ToolTip"));
            this.chkUseProxyLinksForBibleVerses.UseVisualStyleBackColor = true;
            // 
            // chkUseProxyLinksForLinks
            // 
            resources.ApplyResources(this.chkUseProxyLinksForLinks, "chkUseProxyLinksForLinks");
            this.chkUseProxyLinksForLinks.Name = "chkUseProxyLinksForLinks";
            this.ttNotesPageFolder.SetToolTip(this.chkUseProxyLinksForLinks, resources.GetString("chkUseProxyLinksForLinks.ToolTip"));
            this.chkUseProxyLinksForLinks.UseVisualStyleBackColor = true;
            // 
            // chkUseProxyLinksForStrong
            // 
            resources.ApplyResources(this.chkUseProxyLinksForStrong, "chkUseProxyLinksForStrong");
            this.chkUseProxyLinksForStrong.Name = "chkUseProxyLinksForStrong";
            this.ttNotesPageFolder.SetToolTip(this.chkUseProxyLinksForStrong, resources.GetString("chkUseProxyLinksForStrong.ToolTip"));
            this.chkUseProxyLinksForStrong.UseVisualStyleBackColor = true;
            this.chkUseProxyLinksForStrong.CheckedChanged += new System.EventHandler(this.chkNotOneNoteControls_CheckedChanged);
            // 
            // label9
            // 
            resources.ApplyResources(this.label9, "label9");
            this.label9.Name = "label9";
            this.ttNotesPageFolder.SetToolTip(this.label9, resources.GetString("label9.ToolTip"));
            // 
            // cbLanguage
            // 
            resources.ApplyResources(this.cbLanguage, "cbLanguage");
            this.cbLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbLanguage.FormattingEnabled = true;
            this.cbLanguage.Name = "cbLanguage";
            this.ttNotesPageFolder.SetToolTip(this.cbLanguage, resources.GetString("cbLanguage.ToolTip"));
            // 
            // groupBox2
            // 
            resources.ApplyResources(this.groupBox2, "groupBox2");
            this.groupBox2.Controls.Add(this.chkUseDifferentPages);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.tbNotesPageName);
            this.groupBox2.Controls.Add(this.chkExcludedVersesLinking);
            this.groupBox2.Controls.Add(this.tbNotesPageWidth);
            this.groupBox2.Controls.Add(this.chkExpandMultiVersesLinking);
            this.groupBox2.Controls.Add(this.lblNotesPageWidth);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.TabStop = false;
            this.ttNotesPageFolder.SetToolTip(this.groupBox2, resources.GetString("groupBox2.ToolTip"));
            // 
            // chkUseDifferentPages
            // 
            resources.ApplyResources(this.chkUseDifferentPages, "chkUseDifferentPages");
            this.chkUseDifferentPages.Checked = true;
            this.chkUseDifferentPages.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseDifferentPages.Name = "chkUseDifferentPages";
            this.ttNotesPageFolder.SetToolTip(this.chkUseDifferentPages, resources.GetString("chkUseDifferentPages.ToolTip"));
            this.chkUseDifferentPages.UseVisualStyleBackColor = true;
            // 
            // label7
            // 
            resources.ApplyResources(this.label7, "label7");
            this.label7.Name = "label7";
            this.ttNotesPageFolder.SetToolTip(this.label7, resources.GetString("label7.ToolTip"));
            // 
            // tbNotesPageName
            // 
            resources.ApplyResources(this.tbNotesPageName, "tbNotesPageName");
            this.tbNotesPageName.Name = "tbNotesPageName";
            this.ttNotesPageFolder.SetToolTip(this.tbNotesPageName, resources.GetString("tbNotesPageName.ToolTip"));
            // 
            // chkExcludedVersesLinking
            // 
            resources.ApplyResources(this.chkExcludedVersesLinking, "chkExcludedVersesLinking");
            this.chkExcludedVersesLinking.Name = "chkExcludedVersesLinking";
            this.ttNotesPageFolder.SetToolTip(this.chkExcludedVersesLinking, resources.GetString("chkExcludedVersesLinking.ToolTip"));
            this.chkExcludedVersesLinking.UseVisualStyleBackColor = true;
            // 
            // tbNotesPageWidth
            // 
            resources.ApplyResources(this.tbNotesPageWidth, "tbNotesPageWidth");
            this.tbNotesPageWidth.Name = "tbNotesPageWidth";
            this.ttNotesPageFolder.SetToolTip(this.tbNotesPageWidth, resources.GetString("tbNotesPageWidth.ToolTip"));
            // 
            // chkExpandMultiVersesLinking
            // 
            resources.ApplyResources(this.chkExpandMultiVersesLinking, "chkExpandMultiVersesLinking");
            this.chkExpandMultiVersesLinking.Checked = true;
            this.chkExpandMultiVersesLinking.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkExpandMultiVersesLinking.Name = "chkExpandMultiVersesLinking";
            this.ttNotesPageFolder.SetToolTip(this.chkExpandMultiVersesLinking, resources.GetString("chkExpandMultiVersesLinking.ToolTip"));
            this.chkExpandMultiVersesLinking.UseVisualStyleBackColor = true;
            // 
            // lblNotesPageWidth
            // 
            resources.ApplyResources(this.lblNotesPageWidth, "lblNotesPageWidth");
            this.lblNotesPageWidth.Name = "lblNotesPageWidth";
            this.ttNotesPageFolder.SetToolTip(this.lblNotesPageWidth, resources.GetString("lblNotesPageWidth.ToolTip"));
            // 
            // groupBox1
            // 
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Controls.Add(this.chkUseRubbishPage);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.chkRubbishExcludedVersesLinking);
            this.groupBox1.Controls.Add(this.tbRubbishNotesPageName);
            this.groupBox1.Controls.Add(this.chkRubbishExpandMultiVersesLinking);
            this.groupBox1.Controls.Add(this.tbRubbishNotesPageWidth);
            this.groupBox1.Controls.Add(this.lblRubbishNotesPageWidth);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            this.ttNotesPageFolder.SetToolTip(this.groupBox1, resources.GetString("groupBox1.ToolTip"));
            // 
            // chkUseRubbishPage
            // 
            resources.ApplyResources(this.chkUseRubbishPage, "chkUseRubbishPage");
            this.chkUseRubbishPage.Name = "chkUseRubbishPage";
            this.ttNotesPageFolder.SetToolTip(this.chkUseRubbishPage, resources.GetString("chkUseRubbishPage.ToolTip"));
            this.chkUseRubbishPage.UseVisualStyleBackColor = true;
            this.chkUseRubbishPage.CheckedChanged += new System.EventHandler(this.chkUseRubbishPage_CheckedChanged);
            // 
            // label10
            // 
            resources.ApplyResources(this.label10, "label10");
            this.label10.Name = "label10";
            this.ttNotesPageFolder.SetToolTip(this.label10, resources.GetString("label10.ToolTip"));
            // 
            // chkRubbishExcludedVersesLinking
            // 
            resources.ApplyResources(this.chkRubbishExcludedVersesLinking, "chkRubbishExcludedVersesLinking");
            this.chkRubbishExcludedVersesLinking.Name = "chkRubbishExcludedVersesLinking";
            this.ttNotesPageFolder.SetToolTip(this.chkRubbishExcludedVersesLinking, resources.GetString("chkRubbishExcludedVersesLinking.ToolTip"));
            this.chkRubbishExcludedVersesLinking.UseVisualStyleBackColor = true;
            // 
            // tbRubbishNotesPageName
            // 
            resources.ApplyResources(this.tbRubbishNotesPageName, "tbRubbishNotesPageName");
            this.tbRubbishNotesPageName.Name = "tbRubbishNotesPageName";
            this.ttNotesPageFolder.SetToolTip(this.tbRubbishNotesPageName, resources.GetString("tbRubbishNotesPageName.ToolTip"));
            // 
            // chkRubbishExpandMultiVersesLinking
            // 
            resources.ApplyResources(this.chkRubbishExpandMultiVersesLinking, "chkRubbishExpandMultiVersesLinking");
            this.chkRubbishExpandMultiVersesLinking.Name = "chkRubbishExpandMultiVersesLinking";
            this.ttNotesPageFolder.SetToolTip(this.chkRubbishExpandMultiVersesLinking, resources.GetString("chkRubbishExpandMultiVersesLinking.ToolTip"));
            this.chkRubbishExpandMultiVersesLinking.UseVisualStyleBackColor = true;
            // 
            // tbRubbishNotesPageWidth
            // 
            resources.ApplyResources(this.tbRubbishNotesPageWidth, "tbRubbishNotesPageWidth");
            this.tbRubbishNotesPageWidth.Name = "tbRubbishNotesPageWidth";
            this.ttNotesPageFolder.SetToolTip(this.tbRubbishNotesPageWidth, resources.GetString("tbRubbishNotesPageWidth.ToolTip"));
            // 
            // lblRubbishNotesPageWidth
            // 
            resources.ApplyResources(this.lblRubbishNotesPageWidth, "lblRubbishNotesPageWidth");
            this.lblRubbishNotesPageWidth.Name = "lblRubbishNotesPageWidth";
            this.ttNotesPageFolder.SetToolTip(this.lblRubbishNotesPageWidth, resources.GetString("lblRubbishNotesPageWidth.ToolTip"));
            // 
            // chkDefaultParameters
            // 
            resources.ApplyResources(this.chkDefaultParameters, "chkDefaultParameters");
            this.chkDefaultParameters.Name = "chkDefaultParameters";
            this.ttNotesPageFolder.SetToolTip(this.chkDefaultParameters, resources.GetString("chkDefaultParameters.ToolTip"));
            this.chkDefaultParameters.UseVisualStyleBackColor = true;
            this.chkDefaultParameters.CheckedChanged += new System.EventHandler(this.chkDefaultPageNameParameters_CheckedChanged);
            // 
            // tbBookOverviewName
            // 
            resources.ApplyResources(this.tbBookOverviewName, "tbBookOverviewName");
            this.tbBookOverviewName.Name = "tbBookOverviewName";
            this.ttNotesPageFolder.SetToolTip(this.tbBookOverviewName, resources.GetString("tbBookOverviewName.ToolTip"));
            // 
            // label6
            // 
            resources.ApplyResources(this.label6, "label6");
            this.label6.Name = "label6";
            this.ttNotesPageFolder.SetToolTip(this.label6, resources.GetString("label6.ToolTip"));
            // 
            // tbCommentsPageName
            // 
            resources.ApplyResources(this.tbCommentsPageName, "tbCommentsPageName");
            this.tbCommentsPageName.Name = "tbCommentsPageName";
            this.ttNotesPageFolder.SetToolTip(this.tbCommentsPageName, resources.GetString("tbCommentsPageName.ToolTip"));
            // 
            // label5
            // 
            resources.ApplyResources(this.label5, "label5");
            this.label5.Name = "label5";
            this.ttNotesPageFolder.SetToolTip(this.label5, resources.GetString("label5.ToolTip"));
            // 
            // tabPage3
            // 
            resources.ApplyResources(this.tabPage3, "tabPage3");
            this.tabPage3.BackColor = System.Drawing.SystemColors.Control;
            this.tabPage3.Controls.Add(this.btnModuleChecker);
            this.tabPage3.Controls.Add(this.btnConverter);
            this.tabPage3.Controls.Add(this.btnBackup);
            this.tabPage3.Controls.Add(this.btnResizeBibleTables);
            this.tabPage3.Controls.Add(this.btnDeleteNotesPages);
            this.tabPage3.Controls.Add(this.btnRelinkComments);
            this.tabPage3.Name = "tabPage3";
            this.ttNotesPageFolder.SetToolTip(this.tabPage3, resources.GetString("tabPage3.ToolTip"));
            this.tabPage3.Enter += new System.EventHandler(this.tabPage1_Enter);
            // 
            // btnModuleChecker
            // 
            resources.ApplyResources(this.btnModuleChecker, "btnModuleChecker");
            this.btnModuleChecker.Name = "btnModuleChecker";
            this.ttNotesPageFolder.SetToolTip(this.btnModuleChecker, resources.GetString("btnModuleChecker.ToolTip"));
            this.btnModuleChecker.UseVisualStyleBackColor = true;
            this.btnModuleChecker.Click += new System.EventHandler(this.btnModuleChecker_Click);
            // 
            // btnConverter
            // 
            resources.ApplyResources(this.btnConverter, "btnConverter");
            this.btnConverter.Name = "btnConverter";
            this.ttNotesPageFolder.SetToolTip(this.btnConverter, resources.GetString("btnConverter.ToolTip"));
            this.btnConverter.UseVisualStyleBackColor = true;
            this.btnConverter.Click += new System.EventHandler(this.btnConverter_Click);
            // 
            // btnBackup
            // 
            resources.ApplyResources(this.btnBackup, "btnBackup");
            this.btnBackup.Name = "btnBackup";
            this.ttNotesPageFolder.SetToolTip(this.btnBackup, resources.GetString("btnBackup.ToolTip"));
            this.btnBackup.UseVisualStyleBackColor = true;
            this.btnBackup.Click += new System.EventHandler(this.btnBackup_Click);
            // 
            // btnResizeBibleTables
            // 
            resources.ApplyResources(this.btnResizeBibleTables, "btnResizeBibleTables");
            this.btnResizeBibleTables.Name = "btnResizeBibleTables";
            this.ttNotesPageFolder.SetToolTip(this.btnResizeBibleTables, resources.GetString("btnResizeBibleTables.ToolTip"));
            this.btnResizeBibleTables.UseVisualStyleBackColor = true;
            this.btnResizeBibleTables.Click += new System.EventHandler(this.btnResizeBibleTables_Click);
            // 
            // btnDeleteNotesPages
            // 
            resources.ApplyResources(this.btnDeleteNotesPages, "btnDeleteNotesPages");
            this.btnDeleteNotesPages.Name = "btnDeleteNotesPages";
            this.ttNotesPageFolder.SetToolTip(this.btnDeleteNotesPages, resources.GetString("btnDeleteNotesPages.ToolTip"));
            this.btnDeleteNotesPages.UseVisualStyleBackColor = true;
            this.btnDeleteNotesPages.Click += new System.EventHandler(this.btnDeleteNotesPages_Click);
            // 
            // btnRelinkComments
            // 
            resources.ApplyResources(this.btnRelinkComments, "btnRelinkComments");
            this.btnRelinkComments.Name = "btnRelinkComments";
            this.ttNotesPageFolder.SetToolTip(this.btnRelinkComments, resources.GetString("btnRelinkComments.ToolTip"));
            this.btnRelinkComments.UseVisualStyleBackColor = true;
            this.btnRelinkComments.Click += new System.EventHandler(this.btnRelinkComments_Click);
            // 
            // tpModules
            // 
            resources.ApplyResources(this.tpModules, "tpModules");
            this.tpModules.BackColor = System.Drawing.SystemColors.Control;
            this.tpModules.Controls.Add(this.btnDictionariesManagement);
            this.tpModules.Controls.Add(this.btnSupplementalBibleManagement);
            this.tpModules.Controls.Add(this.hlModules);
            this.tpModules.Controls.Add(this.lblModulesLink);
            this.tpModules.Controls.Add(this.pnModules);
            this.tpModules.Controls.Add(this.lblMustSelectModule);
            this.tpModules.Controls.Add(this.lblMustUploadModule);
            this.tpModules.Controls.Add(this.btnUploadModule);
            this.tpModules.Name = "tpModules";
            this.ttNotesPageFolder.SetToolTip(this.tpModules, resources.GetString("tpModules.ToolTip"));
            this.tpModules.Enter += new System.EventHandler(this.tabPage4_Enter);
            // 
            // btnDictionariesManagement
            // 
            resources.ApplyResources(this.btnDictionariesManagement, "btnDictionariesManagement");
            this.btnDictionariesManagement.Name = "btnDictionariesManagement";
            this.ttNotesPageFolder.SetToolTip(this.btnDictionariesManagement, resources.GetString("btnDictionariesManagement.ToolTip"));
            this.btnDictionariesManagement.UseVisualStyleBackColor = true;
            this.btnDictionariesManagement.Click += new System.EventHandler(this.btnDictionariesManagement_Click);
            // 
            // btnSupplementalBibleManagement
            // 
            resources.ApplyResources(this.btnSupplementalBibleManagement, "btnSupplementalBibleManagement");
            this.btnSupplementalBibleManagement.Name = "btnSupplementalBibleManagement";
            this.ttNotesPageFolder.SetToolTip(this.btnSupplementalBibleManagement, resources.GetString("btnSupplementalBibleManagement.ToolTip"));
            this.btnSupplementalBibleManagement.UseVisualStyleBackColor = true;
            this.btnSupplementalBibleManagement.Click += new System.EventHandler(this.btnSupplementalBibleManagement_Click);
            // 
            // hlModules
            // 
            resources.ApplyResources(this.hlModules, "hlModules");
            this.hlModules.Name = "hlModules";
            this.hlModules.TabStop = true;
            this.ttNotesPageFolder.SetToolTip(this.hlModules, resources.GetString("hlModules.ToolTip"));
            this.hlModules.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.hlModules_LinkClicked);
            // 
            // lblModulesLink
            // 
            resources.ApplyResources(this.lblModulesLink, "lblModulesLink");
            this.lblModulesLink.Name = "lblModulesLink";
            this.ttNotesPageFolder.SetToolTip(this.lblModulesLink, resources.GetString("lblModulesLink.ToolTip"));
            // 
            // pnModules
            // 
            resources.ApplyResources(this.pnModules, "pnModules");
            this.pnModules.Name = "pnModules";
            this.ttNotesPageFolder.SetToolTip(this.pnModules, resources.GetString("pnModules.ToolTip"));
            // 
            // lblMustSelectModule
            // 
            resources.ApplyResources(this.lblMustSelectModule, "lblMustSelectModule");
            this.lblMustSelectModule.ForeColor = System.Drawing.Color.Red;
            this.lblMustSelectModule.Name = "lblMustSelectModule";
            this.ttNotesPageFolder.SetToolTip(this.lblMustSelectModule, resources.GetString("lblMustSelectModule.ToolTip"));
            // 
            // lblMustUploadModule
            // 
            resources.ApplyResources(this.lblMustUploadModule, "lblMustUploadModule");
            this.lblMustUploadModule.ForeColor = System.Drawing.Color.Red;
            this.lblMustUploadModule.Name = "lblMustUploadModule";
            this.ttNotesPageFolder.SetToolTip(this.lblMustUploadModule, resources.GetString("lblMustUploadModule.ToolTip"));
            // 
            // btnUploadModule
            // 
            resources.ApplyResources(this.btnUploadModule, "btnUploadModule");
            this.btnUploadModule.Name = "btnUploadModule";
            this.ttNotesPageFolder.SetToolTip(this.btnUploadModule, resources.GetString("btnUploadModule.ToolTip"));
            this.btnUploadModule.UseVisualStyleBackColor = true;
            this.btnUploadModule.Click += new System.EventHandler(this.btnUploadModule_Click);
            // 
            // btnApply
            // 
            resources.ApplyResources(this.btnApply, "btnApply");
            this.btnApply.Name = "btnApply";
            this.ttNotesPageFolder.SetToolTip(this.btnApply, resources.GetString("btnApply.ToolTip"));
            this.btnApply.UseVisualStyleBackColor = true;
            this.btnApply.Click += new System.EventHandler(this.btnApply_Click);
            // 
            // btnClose
            // 
            resources.ApplyResources(this.btnClose, "btnClose");
            this.btnClose.Name = "btnClose";
            this.ttNotesPageFolder.SetToolTip(this.btnClose, resources.GetString("btnClose.ToolTip"));
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // lblWarning
            // 
            resources.ApplyResources(this.lblWarning, "lblWarning");
            this.lblWarning.ForeColor = System.Drawing.Color.Red;
            this.lblWarning.Name = "lblWarning";
            this.ttNotesPageFolder.SetToolTip(this.lblWarning, resources.GetString("lblWarning.ToolTip"));
            // 
            // lblProgressInfo
            // 
            resources.ApplyResources(this.lblProgressInfo, "lblProgressInfo");
            this.lblProgressInfo.Name = "lblProgressInfo";
            this.ttNotesPageFolder.SetToolTip(this.lblProgressInfo, resources.GetString("lblProgressInfo.ToolTip"));
            // 
            // pbMain
            // 
            resources.ApplyResources(this.pbMain, "pbMain");
            this.pbMain.Name = "pbMain";
            this.pbMain.Step = 3;
            this.ttNotesPageFolder.SetToolTip(this.pbMain, resources.GetString("pbMain.ToolTip"));
            // 
            // btnOK
            // 
            resources.ApplyResources(this.btnOK, "btnOK");
            this.btnOK.Name = "btnOK";
            this.ttNotesPageFolder.SetToolTip(this.btnOK, resources.GetString("btnOK.ToolTip"));
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // folderBrowserDialog
            // 
            resources.ApplyResources(this.folderBrowserDialog, "folderBrowserDialog");
            // 
            // saveFileDialog
            // 
            resources.ApplyResources(this.saveFileDialog, "saveFileDialog");
            // 
            // openModuleFileDialog
            // 
            this.openModuleFileDialog.DefaultExt = "isbt";
            resources.ApplyResources(this.openModuleFileDialog, "openModuleFileDialog");
            // 
            // notesPagesFolderBrowserDialog
            // 
            resources.ApplyResources(this.notesPagesFolderBrowserDialog, "notesPagesFolderBrowserDialog");
            // 
            // MainForm
            // 
            this.AcceptButton = this.btnOK;
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.splitContainer1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.ttNotesPageFolder.SetToolTip(this, resources.GetString("$this.ToolTip"));
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainForm_FormClosed);
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Shown += new System.EventHandler(this.MainForm_Shown);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.tbcMain.ResumeLayout(false);
            this.tpNotebooks.ResumeLayout(false);
            this.tpNotebooks.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tpModules.ResumeLayout(false);
            this.tpModules.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.TabControl tbcMain;
        private System.Windows.Forms.TabPage tpNotebooks;
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
        private System.Windows.Forms.Label lblSelectSingleNotebook;
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
        private System.Windows.Forms.TextBox tbCommentsPageName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button btnDeleteNotesPages;
        private System.Windows.Forms.Button btnRelinkComments;
        private System.Windows.Forms.Button btnBibleStudyNotebookSetPath;
        private System.Windows.Forms.Button btnBibleCommentsNotebookSetPath;
        private System.Windows.Forms.Button btnBibleNotebookSetPath; 
        private System.Windows.Forms.Button btnSingleNotebookSetPath;
        private System.Windows.Forms.Button btnResizeBibleTables;
        private System.Windows.Forms.CheckBox chkDefaultParameters;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.Label lblProgressInfo;
        private System.Windows.Forms.Label lblWarning;
        private System.Windows.Forms.Button btnBackup;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox chkExcludedVersesLinking;
        private System.Windows.Forms.TextBox tbNotesPageWidth;
        private System.Windows.Forms.CheckBox chkExpandMultiVersesLinking;
        private System.Windows.Forms.Label lblNotesPageWidth;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkUseRubbishPage;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.CheckBox chkRubbishExcludedVersesLinking;
        private System.Windows.Forms.TextBox tbRubbishNotesPageName;
        private System.Windows.Forms.CheckBox chkRubbishExpandMultiVersesLinking;
        private System.Windows.Forms.TextBox tbRubbishNotesPageWidth;
        private System.Windows.Forms.Label lblRubbishNotesPageWidth;
        private System.Windows.Forms.CheckBox chkUseDifferentPages;
        private System.Windows.Forms.Button btnBibleNotesPagesNotebookSetPath;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.CheckBox chkCreateBibleNotesPagesNotebookFromTemplate;
        private System.Windows.Forms.ComboBox cbBibleNotesPagesNotebook;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox cbLanguage;
        private System.Windows.Forms.TabPage tpModules;
        private System.Windows.Forms.Button btnUploadModule;
        private System.Windows.Forms.Label lblMustUploadModule;
        private System.Windows.Forms.Label lblMustSelectModule;
        private System.Windows.Forms.Panel pnModules;
        private System.Windows.Forms.OpenFileDialog openModuleFileDialog;
        private System.Windows.Forms.LinkLabel hlModules;
        private System.Windows.Forms.Label lblModulesLink;
        private System.Windows.Forms.Button btnSupplementalBibleManagement;
        private System.Windows.Forms.Button btnDictionariesManagement;
        private System.Windows.Forms.Button btnApply;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnModuleChecker;
        private System.Windows.Forms.Button btnConverter;
        private System.Windows.Forms.CheckBox chkUseProxyLinksForStrong;
        private System.Windows.Forms.CheckBox chkUseProxyLinksForLinks;
        private System.Windows.Forms.CheckBox chkUseFolderForBibleNotesPages;
        private System.Windows.Forms.Button btnBibleNotesPagesSetFolder;
        private System.Windows.Forms.TextBox tbBibleNotesPagesFolder;
        private System.Windows.Forms.FolderBrowserDialog notesPagesFolderBrowserDialog;
        private System.Windows.Forms.CheckBox chkUseProxyLinksForBibleVerses;
        private System.Windows.Forms.ToolTip ttNotesPageFolder;
        private System.Windows.Forms.CheckBox chkUseAdvancedProxyForOneNoteLinks;
        private System.Windows.Forms.CheckBox chkUseCommaDelimeter;

    }
}

