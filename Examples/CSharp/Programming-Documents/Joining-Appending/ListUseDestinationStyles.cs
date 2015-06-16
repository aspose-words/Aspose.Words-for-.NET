﻿//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Fields;
using System.Text;
using System.Collections;
using Aspose.Words.Lists;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class ListUseDestinationStyles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.DestinationList.doc");
            Document srcDoc = new Document(dataDir + "TestFile.SourceList.doc");

            // Set the source document to continue straight after the end of the destination document.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Keep track of the lists that are created.
            Hashtable newLists = new Hashtable();

            // Iterate through all paragraphs in the document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                if (para.IsListItem)
                {
                    int listId = para.ListFormat.List.ListId;

                    // Check if the destination document contains a list with this ID already. If it does then this may
                    // cause the two lists to run together. Create a copy of the list in the source document instead.
                    if (dstDoc.Lists.GetListByListId(listId) != null)
                    {
                        List currentList;
                        // A newly copied list already exists for this ID, retrieve the stored list and use it on 
                        // the current paragraph.
                        if (newLists.Contains(listId))
                        {
                            currentList = (List)newLists[listId];
                        }
                        else
                        {
                            // Add a copy of this list to the document and store it for later reference.
                            currentList = srcDoc.Lists.AddCopy(para.ListFormat.List);
                            newLists.Add(listId, currentList);
                        }

                        // Set the list of this paragraph  to the copied list.
                        para.ListFormat.List = currentList;
                    }
                }
            }

            // Append the source document to end of the destination document.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);

            // Save the combined document to disk.
            dstDoc.Save(dataDir + "TestFile.ListUseDestinationStyles Out.docx");

            Console.WriteLine("\nDocument appended successfully without continuing any list numberings.\nFile saved at " + dataDir + "TestFile.ListUseDestinationStyles Out.docx");
        }
    }
}
