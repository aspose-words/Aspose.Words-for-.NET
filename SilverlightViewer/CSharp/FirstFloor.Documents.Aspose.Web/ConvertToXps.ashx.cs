//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2011 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Web;

using Aspose.Words;

namespace FirstFloor.Documents.Aspose.Web
{
    public class ConvertToXps : IHttpHandler
    {
        public void ProcessRequest(HttpContext context)
        {
            if (context.Request.ContentLength > 2 << 18) {
                throw new NotSupportedException();
            }
            var document = new Document(context.Request.InputStream);
            document.Save(context.Response.OutputStream, SaveFormat.Xps);
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}
