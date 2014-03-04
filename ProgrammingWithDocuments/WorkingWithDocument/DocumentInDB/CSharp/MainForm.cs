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

namespace DocumentInDBExample
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
                        
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");
            
            
        }

        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }
}