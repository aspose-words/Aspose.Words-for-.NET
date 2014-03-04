// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Words;

namespace Excel2Word
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            string wordsLicenseFile = Path.Combine(Application.StartupPath, "Aspose.Words.lic");
            if (File.Exists(wordsLicenseFile))
            {
                //This shows how to license Aspose.Words.
                //If you don't specify a license, Aspose.Words works in evaluation mode.
                Aspose.Words.License license = new Aspose.Words.License();
                license.SetLicense(wordsLicenseFile);
            }

            string cellsLicenseFile = Path.Combine(Application.StartupPath, "Aspose.Cells.lic");
            if (File.Exists(cellsLicenseFile))
            {
                //This shows how to license Aspose.Cells.
                //If you don't specify a license, Aspose.Cells works in evaluation mode.
                Aspose.Cells.License license = new Aspose.Cells.License();
                license.SetLicense(cellsLicenseFile);
            }
        }

        private void buttonConvert_Click(object sender, EventArgs e)
        {
            try
            {
                //Show the open dialog
                if (!openFileDialog.ShowDialog().Equals(DialogResult.OK))
                    return;

                //Show the save dialog to select the destination file name and then run the demo.
                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + " Out";
                if (!saveFileDialog.ShowDialog().Equals(DialogResult.OK))
                    return;

                RunConvert(openFileDialog.FileName, saveFileDialog.FileName);

                MessageBox.Show("Done!", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void RunConvert(string srcFileName, string dstFileName)
        {
            //Open Excel Workbook using Aspose.Cells.
            Workbook workbook = new Workbook(srcFileName);
            
            //Convert workbook to Word document
            ConverterXls2Doc converter = new ConverterXls2Doc();
            Document doc = converter.Convert(workbook);

            // Save using Aspose.Words. 
            doc.Save(dstFileName);
        }
    }
}