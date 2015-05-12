//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Windows.Forms;
using Aspose.Words;
#if CellsInstalled
using Aspose.Cells;
#endif

namespace Excel2WordExample
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
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
            #if CellsInstalled
            //Open Excel Workbook using Aspose.Cells.
            Workbook workbook = new Workbook(srcFileName);
            
            //Convert workbook to Word document
            ConverterXls2Doc converter = new ConverterXls2Doc();
            Document doc = converter.Convert(workbook);

            // Save using Aspose.Words. 
            doc.Save(dstFileName);
#else
            throw new InvalidOperationException(@"This example requires the use of Aspose.Cells." + 
                                    "Make sure Aspose.Cells.dll is present in the bin\net2.0 folder.");
#endif
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}