using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using BibleCommon.Services;
using System.IO;
using BibleCommon.Helpers;
using BibleCommon.Common;
using System.Diagnostics;

namespace BibleCommon.UI.Forms
{
    public partial class ErrorsForm : Form
    {
        public List<ErrorsList> AllErrors { get; set; }
        private Microsoft.Office.Interop.OneNote.Application _oneNoteApp;

        public string LogFilePath { get; set; }

        public ErrorsForm()
        {
            AllErrors = new List<ErrorsList>();            

            InitializeComponent();            
        }

        public ErrorsForm(List<string> errors)
            : this()
        {
            AllErrors.Add(new ErrorsList(errors));
        }

        public ErrorsForm(List<LogItem> errors)
            : this()
        {
            AllErrors.Add(new ErrorsList(errors));
        }

        public void ClearErrors()
        {
            AllErrors.Clear();
            lbErrors.Items.Clear();
        }

        private void Errors_Load(object sender, EventArgs e)
        {
            try
            {
                if (AllErrors.All(errors => errors.Count == 0))
                    Close();

                FormExtensions.SetFocus(this);

                using (Graphics g = lbErrors.CreateGraphics())
                {
                    foreach (var errors in AllErrors)
                    {
                        if (!string.IsNullOrEmpty(errors.ErrorsDecription))
                            lbErrors.Items.Add(string.Format("{0} ({1})", errors.ErrorsDecription, errors.Count));

                        int index = 1;

                        foreach (LogItem item in errors)
                        {
                            var errorItem = item;
                            errorItem.Message = string.Format("{0}. {1}", index++, errorItem.Message);
                            lbErrors.Items.Add(errorItem);

                            //int width = Convert.ToInt32(error.Length * 5.75);
                            int width = (int)g.MeasureString(errorItem, lbErrors.Font).Width + 100;
                            if (width > lbErrors.HorizontalExtent)
                                lbErrors.HorizontalExtent = width;
                        }
                        lbErrors.Items.Add(string.Empty);
                    }
                }

                if (string.IsNullOrEmpty(LogFilePath))
                    btnOpenLog.Visible = false;
            }
            catch (Exception ex)
            {
                FormLogger.LogError(ex);
            }
        }

        public void SaveErrorsToFile(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Create))
            {
                using (StreamWriter sw = new StreamWriter(fs))
                {
                    foreach (var errors in AllErrors)
                    {
                        if (!string.IsNullOrEmpty(errors.ErrorsDecription))
                            sw.WriteLine(string.Format("{0} ({1})", errors.ErrorsDecription, errors.Count));

                        int index = 1;
                        foreach (var error in errors)
                        {
                            sw.WriteLine(string.Format("{0}. {1}", index++, error));
                        }
                        sw.WriteLine(string.Empty);
                    }
                    sw.Flush();
                }
            }
        }

        private void btnSaveToFile_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                SaveErrorsToFile(saveFileDialog.FileName);

                MessageBox.Show(BibleCommon.Resources.Constants.SuccessfullySaved);
            }
        }

        private void btnOpenLog_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(LogFilePath))
                Process.Start(LogFilePath);
        }      

        private void lbErrors_MouseClick(object sender, MouseEventArgs e)
        {
            TryToGoToErrorObject();
        }

        private void TryToGoToErrorObject()
        {
            if (lbErrors.SelectedItem != null)
            {
                if (lbErrors.SelectedItem is LogItem)
                {
                    var item = (LogItem)lbErrors.SelectedItem;
                    if (!string.IsNullOrEmpty(item.PageId) && !string.IsNullOrEmpty(item.ContentObjectId))
                    {
                        OneNoteUtils.UseOneNoteAPI(ref _oneNoteApp, () =>
                        {
                            _oneNoteApp.NavigateTo(item.PageId, item.ContentObjectId);                            
                        });
                        OneNoteUtils.SetActiveCurrentWindow(ref _oneNoteApp);
                    }
                }
            }
        }

        private void ErrorsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            OneNoteUtils.ReleaseOneNoteApp(ref _oneNoteApp);
        }
    }
}
