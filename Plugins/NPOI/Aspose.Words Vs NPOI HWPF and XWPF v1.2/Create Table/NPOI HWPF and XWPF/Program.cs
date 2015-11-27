using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NPOI_HWPF_and_XWPF
{
    class Program
    {
        static void Main(string[] args)
        {
            XWPFDocument doc = new XWPFDocument();

            // New 2x2 table
            XWPFTable tableOne = doc.CreateTable();
            XWPFTableRow tableOneRowOne = tableOne.GetRow(0);
            tableOneRowOne.GetCell(0).SetText("Hello");
            tableOneRowOne.AddNewTableCell().SetText("Word");

            XWPFTableRow tableOneRowTwo = tableOne.CreateRow();
            tableOneRowTwo.GetCell(0).SetText("This is");
            tableOneRowTwo.GetCell(1).SetText("a table");

            using (FileStream sw = File.Create("Create_Table.docx"))
            {
                doc.Write(sw);
            }


        }
    }
}
