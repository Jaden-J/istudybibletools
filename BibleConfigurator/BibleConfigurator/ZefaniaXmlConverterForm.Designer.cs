﻿namespace BibleConfigurator
{
    partial class ZefaniaXmlConverterForm
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
            this.btnOk = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.tbShortName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.tbDisplayName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.tbZefaniaXmlFilePath = new System.Windows.Forms.TextBox();
            this.btnZefaniaXmlFilePath = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.tbVersion = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.tbNotebookSummaryOfNotesName = new System.Windows.Forms.TextBox();
            this.tbNotebookBibleCommentsName = new System.Windows.Forms.TextBox();
            this.tbNotebookBibleName = new System.Windows.Forms.TextBox();
            this.chkNotebookSummaryOfNotesGenerate = new System.Windows.Forms.CheckBox();
            this.chkNotebookBibleCommentsGenerate = new System.Windows.Forms.CheckBox();
            this.chkNotebookBibleGenerate = new System.Windows.Forms.CheckBox();
            this.cbNotebookSummaryOfNotes = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.cbNotebookBibleComments = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.cbNotebookBibleStudy = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.cbNotebookBible = new System.Windows.Forms.ComboBox();
            this.gbAdditionalParameters = new System.Windows.Forms.GroupBox();
            this.chkRemoveStrongNumbers = new System.Windows.Forms.CheckBox();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tbLocale = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.tbResultDirectory = new System.Windows.Forms.TextBox();
            this.btnResultFilePath = new System.Windows.Forms.Button();
            this.pbMain = new System.Windows.Forms.ProgressBar();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox1.SuspendLayout();
            this.gbAdditionalParameters.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(575, 305);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 0;
            this.btnOk.Text = "OK";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(656, 305);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(412, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Локаль";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 41);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Имя модуля";
            // 
            // tbShortName
            // 
            this.tbShortName.Location = new System.Drawing.Point(234, 38);
            this.tbShortName.Name = "tbShortName";
            this.tbShortName.Size = new System.Drawing.Size(100, 20);
            this.tbShortName.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 67);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(136, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "Полное название модуля";
            // 
            // tbDisplayName
            // 
            this.tbDisplayName.Location = new System.Drawing.Point(234, 64);
            this.tbDisplayName.Name = "tbDisplayName";
            this.tbDisplayName.Size = new System.Drawing.Size(497, 20);
            this.tbDisplayName.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(13, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(194, 13);
            this.label4.TabIndex = 8;
            this.label4.Text = "Файл Библии в формате ZefaniaXML";
            // 
            // tbZefaniaXmlFilePath
            // 
            this.tbZefaniaXmlFilePath.Location = new System.Drawing.Point(234, 12);
            this.tbZefaniaXmlFilePath.Name = "tbZefaniaXmlFilePath";
            this.tbZefaniaXmlFilePath.Size = new System.Drawing.Size(465, 20);
            this.tbZefaniaXmlFilePath.TabIndex = 9;
            this.tbZefaniaXmlFilePath.MouseClick += new System.Windows.Forms.MouseEventHandler(this.tbZefaniaXmlFilePath_MouseClick);
            // 
            // btnZefaniaXmlFilePath
            // 
            this.btnZefaniaXmlFilePath.Location = new System.Drawing.Point(705, 9);
            this.btnZefaniaXmlFilePath.Name = "btnZefaniaXmlFilePath";
            this.btnZefaniaXmlFilePath.Size = new System.Drawing.Size(26, 23);
            this.btnZefaniaXmlFilePath.TabIndex = 10;
            this.btnZefaniaXmlFilePath.Text = "...";
            this.btnZefaniaXmlFilePath.UseVisualStyleBackColor = true;
            this.btnZefaniaXmlFilePath.Click += new System.EventHandler(this.btnZefaniaXmlFilePath_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(602, 41);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 13);
            this.label9.TabIndex = 25;
            this.label9.Text = "Версия";
            // 
            // tbVersion
            // 
            this.tbVersion.Location = new System.Drawing.Point(656, 38);
            this.tbVersion.Name = "tbVersion";
            this.tbVersion.Size = new System.Drawing.Size(75, 20);
            this.tbVersion.TabIndex = 26;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(13, 226);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(0, 13);
            this.label10.TabIndex = 27;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.tbNotebookSummaryOfNotesName);
            this.groupBox1.Controls.Add(this.tbNotebookBibleCommentsName);
            this.groupBox1.Controls.Add(this.tbNotebookBibleName);
            this.groupBox1.Controls.Add(this.chkNotebookSummaryOfNotesGenerate);
            this.groupBox1.Controls.Add(this.chkNotebookBibleCommentsGenerate);
            this.groupBox1.Controls.Add(this.chkNotebookBibleGenerate);
            this.groupBox1.Controls.Add(this.cbNotebookSummaryOfNotes);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.cbNotebookBibleComments);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.cbNotebookBibleStudy);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.cbNotebookBible);
            this.groupBox1.Location = new System.Drawing.Point(16, 116);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(715, 134);
            this.groupBox1.TabIndex = 30;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Параметры записных книжек";
            // 
            // tbNotebookSummaryOfNotesName
            // 
            this.tbNotebookSummaryOfNotesName.Enabled = false;
            this.tbNotebookSummaryOfNotesName.Location = new System.Drawing.Point(553, 103);
            this.tbNotebookSummaryOfNotesName.Name = "tbNotebookSummaryOfNotesName";
            this.tbNotebookSummaryOfNotesName.Size = new System.Drawing.Size(156, 20);
            this.tbNotebookSummaryOfNotesName.TabIndex = 45;
            // 
            // tbNotebookBibleCommentsName
            // 
            this.tbNotebookBibleCommentsName.Enabled = false;
            this.tbNotebookBibleCommentsName.Location = new System.Drawing.Point(553, 76);
            this.tbNotebookBibleCommentsName.Name = "tbNotebookBibleCommentsName";
            this.tbNotebookBibleCommentsName.Size = new System.Drawing.Size(156, 20);
            this.tbNotebookBibleCommentsName.TabIndex = 44;
            // 
            // tbNotebookBibleName
            // 
            this.tbNotebookBibleName.Enabled = false;
            this.tbNotebookBibleName.Location = new System.Drawing.Point(553, 22);
            this.tbNotebookBibleName.Name = "tbNotebookBibleName";
            this.tbNotebookBibleName.Size = new System.Drawing.Size(156, 20);
            this.tbNotebookBibleName.TabIndex = 43;
            // 
            // chkNotebookSummaryOfNotesGenerate
            // 
            this.chkNotebookSummaryOfNotesGenerate.AutoSize = true;
            this.chkNotebookSummaryOfNotesGenerate.Location = new System.Drawing.Point(444, 105);
            this.chkNotebookSummaryOfNotesGenerate.Name = "chkNotebookSummaryOfNotesGenerate";
            this.chkNotebookSummaryOfNotesGenerate.Size = new System.Drawing.Size(103, 17);
            this.chkNotebookSummaryOfNotesGenerate.TabIndex = 42;
            this.chkNotebookSummaryOfNotesGenerate.Text = "Сгенерировать";
            this.chkNotebookSummaryOfNotesGenerate.UseVisualStyleBackColor = true;
            this.chkNotebookSummaryOfNotesGenerate.CheckedChanged += new System.EventHandler(this.chkNotebookSummaryOfNotesGenerate_CheckedChanged);
            // 
            // chkNotebookBibleCommentsGenerate
            // 
            this.chkNotebookBibleCommentsGenerate.AutoSize = true;
            this.chkNotebookBibleCommentsGenerate.Location = new System.Drawing.Point(444, 78);
            this.chkNotebookBibleCommentsGenerate.Name = "chkNotebookBibleCommentsGenerate";
            this.chkNotebookBibleCommentsGenerate.Size = new System.Drawing.Size(103, 17);
            this.chkNotebookBibleCommentsGenerate.TabIndex = 41;
            this.chkNotebookBibleCommentsGenerate.Text = "Сгенерировать";
            this.chkNotebookBibleCommentsGenerate.UseVisualStyleBackColor = true;
            this.chkNotebookBibleCommentsGenerate.CheckedChanged += new System.EventHandler(this.chkNotebookBibleCommentsGenerate_CheckedChanged);
            // 
            // chkNotebookBibleGenerate
            // 
            this.chkNotebookBibleGenerate.AutoSize = true;
            this.chkNotebookBibleGenerate.Location = new System.Drawing.Point(444, 24);
            this.chkNotebookBibleGenerate.Name = "chkNotebookBibleGenerate";
            this.chkNotebookBibleGenerate.Size = new System.Drawing.Size(103, 17);
            this.chkNotebookBibleGenerate.TabIndex = 40;
            this.chkNotebookBibleGenerate.Text = "Сгенерировать";
            this.chkNotebookBibleGenerate.UseVisualStyleBackColor = true;
            this.chkNotebookBibleGenerate.CheckedChanged += new System.EventHandler(this.chkNotebookBibleGenerate_CheckedChanged);
            // 
            // cbNotebookSummaryOfNotes
            // 
            this.cbNotebookSummaryOfNotes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbNotebookSummaryOfNotes.FormattingEnabled = true;
            this.cbNotebookSummaryOfNotes.Location = new System.Drawing.Point(218, 103);
            this.cbNotebookSummaryOfNotes.Name = "cbNotebookSummaryOfNotes";
            this.cbNotebookSummaryOfNotes.Size = new System.Drawing.Size(220, 21);
            this.cbNotebookSummaryOfNotes.TabIndex = 37;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(63, 106);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(98, 13);
            this.label13.TabIndex = 36;
            this.label13.Text = "Сводные заметок";
            // 
            // cbNotebookBibleComments
            // 
            this.cbNotebookBibleComments.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbNotebookBibleComments.FormattingEnabled = true;
            this.cbNotebookBibleComments.Location = new System.Drawing.Point(218, 76);
            this.cbNotebookBibleComments.Name = "cbNotebookBibleComments";
            this.cbNotebookBibleComments.Size = new System.Drawing.Size(220, 21);
            this.cbNotebookBibleComments.TabIndex = 35;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(63, 79);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(126, 13);
            this.label12.TabIndex = 34;
            this.label12.Text = "Комментарии к Библии";
            // 
            // cbNotebookBibleStudy
            // 
            this.cbNotebookBibleStudy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbNotebookBibleStudy.FormattingEnabled = true;
            this.cbNotebookBibleStudy.Location = new System.Drawing.Point(218, 49);
            this.cbNotebookBibleStudy.Name = "cbNotebookBibleStudy";
            this.cbNotebookBibleStudy.Size = new System.Drawing.Size(220, 21);
            this.cbNotebookBibleStudy.TabIndex = 33;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(63, 52);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(95, 13);
            this.label11.TabIndex = 32;
            this.label11.Text = "Изучение Библии";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(63, 25);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(44, 13);
            this.label8.TabIndex = 31;
            this.label8.Text = "Библия";
            // 
            // cbNotebookBible
            // 
            this.cbNotebookBible.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbNotebookBible.FormattingEnabled = true;
            this.cbNotebookBible.Location = new System.Drawing.Point(218, 22);
            this.cbNotebookBible.Name = "cbNotebookBible";
            this.cbNotebookBible.Size = new System.Drawing.Size(220, 21);
            this.cbNotebookBible.TabIndex = 30;
            // 
            // gbAdditionalParameters
            // 
            this.gbAdditionalParameters.Controls.Add(this.chkRemoveStrongNumbers);
            this.gbAdditionalParameters.Location = new System.Drawing.Point(16, 256);
            this.gbAdditionalParameters.Name = "gbAdditionalParameters";
            this.gbAdditionalParameters.Size = new System.Drawing.Size(715, 44);
            this.gbAdditionalParameters.TabIndex = 32;
            this.gbAdditionalParameters.TabStop = false;
            this.gbAdditionalParameters.Text = "Дополнительные параметры";
            // 
            // chkRemoveStrongNumbers
            // 
            this.chkRemoveStrongNumbers.AutoSize = true;
            this.chkRemoveStrongNumbers.Location = new System.Drawing.Point(66, 20);
            this.chkRemoveStrongNumbers.Name = "chkRemoveStrongNumbers";
            this.chkRemoveStrongNumbers.Size = new System.Drawing.Size(154, 17);
            this.chkRemoveStrongNumbers.TabIndex = 0;
            this.chkRemoveStrongNumbers.Text = "Удалить номера Стронга";
            this.chkRemoveStrongNumbers.UseVisualStyleBackColor = true;
            // 
            // openFileDialog
            // 
            this.openFileDialog.DefaultExt = "xml";
            this.openFileDialog.Filter = ".xml files|*.xml";
            // 
            // tbLocale
            // 
            this.tbLocale.Location = new System.Drawing.Point(463, 38);
            this.tbLocale.Name = "tbLocale";
            this.tbLocale.Size = new System.Drawing.Size(100, 20);
            this.tbLocale.TabIndex = 33;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(13, 93);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(122, 13);
            this.label5.TabIndex = 11;
            this.label5.Text = "Папка с результатами";
            // 
            // tbResultDirectory
            // 
            this.tbResultDirectory.Location = new System.Drawing.Point(234, 90);
            this.tbResultDirectory.Name = "tbResultDirectory";
            this.tbResultDirectory.Size = new System.Drawing.Size(465, 20);
            this.tbResultDirectory.TabIndex = 12;
            // 
            // btnResultFilePath
            // 
            this.btnResultFilePath.Location = new System.Drawing.Point(705, 88);
            this.btnResultFilePath.Name = "btnResultFilePath";
            this.btnResultFilePath.Size = new System.Drawing.Size(26, 23);
            this.btnResultFilePath.TabIndex = 13;
            this.btnResultFilePath.Text = "...";
            this.btnResultFilePath.UseVisualStyleBackColor = true;
            this.btnResultFilePath.Click += new System.EventHandler(this.btnResultFilePath_Click);
            // 
            // pbMain
            // 
            this.pbMain.Location = new System.Drawing.Point(16, 306);
            this.pbMain.Name = "pbMain";
            this.pbMain.Size = new System.Drawing.Size(553, 23);
            this.pbMain.TabIndex = 34;
            this.pbMain.Visible = false;
            // 
            // ZefaniaXmlConverterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(743, 341);
            this.Controls.Add(this.pbMain);
            this.Controls.Add(this.tbLocale);
            this.Controls.Add(this.gbAdditionalParameters);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.tbVersion);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.btnResultFilePath);
            this.Controls.Add(this.tbResultDirectory);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.btnZefaniaXmlFilePath);
            this.Controls.Add(this.tbZefaniaXmlFilePath);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.tbDisplayName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.tbShortName);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnOk);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ZefaniaXmlConverterForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ZefaniaXML Конвертер";
            this.Load += new System.EventHandler(this.ZefaniaXmlConverterForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.gbAdditionalParameters.ResumeLayout(false);
            this.gbAdditionalParameters.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox tbShortName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox tbDisplayName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox tbZefaniaXmlFilePath;
        private System.Windows.Forms.Button btnZefaniaXmlFilePath;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox tbVersion;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox cbNotebookSummaryOfNotes;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox cbNotebookBibleComments;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox cbNotebookBibleStudy;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox cbNotebookBible;
        private System.Windows.Forms.GroupBox gbAdditionalParameters;
        private System.Windows.Forms.CheckBox chkRemoveStrongNumbers;
        private System.Windows.Forms.CheckBox chkNotebookSummaryOfNotesGenerate;
        private System.Windows.Forms.CheckBox chkNotebookBibleCommentsGenerate;
        private System.Windows.Forms.CheckBox chkNotebookBibleGenerate;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TextBox tbLocale;
        private System.Windows.Forms.TextBox tbNotebookBibleName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbResultDirectory;
        private System.Windows.Forms.Button btnResultFilePath;
        private System.Windows.Forms.TextBox tbNotebookSummaryOfNotesName;
        private System.Windows.Forms.TextBox tbNotebookBibleCommentsName;
        private System.Windows.Forms.ProgressBar pbMain;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
    }
}