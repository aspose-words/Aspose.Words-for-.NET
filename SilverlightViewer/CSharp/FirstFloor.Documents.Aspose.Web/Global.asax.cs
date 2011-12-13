//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;
using Aspose.Words;

namespace FirstFloor.Documents.Aspose.Web
{
    public class Global : System.Web.HttpApplication
    {

        protected void Application_Start(object sender, EventArgs e)
        {
            // TODO 0 Do not ship source code of this demo project with Aspose.Words.lic embedded in the project. Delete Aspose.Words.lic and this comment before shipping.

            using (Stream licenseStream = Assembly.GetExecutingAssembly().GetManifestResourceStream("FirstFloor.Documents.Aspose.Web.Aspose.Words.lic"))
            {
                if (licenseStream != null)
                {
                    License license = new License();
                    license.SetLicense(licenseStream);
                }
            }
        }
    }
}