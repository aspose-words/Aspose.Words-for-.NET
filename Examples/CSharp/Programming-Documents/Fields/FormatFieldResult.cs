using Aspose.Words.Fields;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Fields
{
    class FormatFieldResult
    {
        public static void Run()
        {
            // ExStart:FormatFieldResult
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            DocumentBuilder builder = new DocumentBuilder();
            Document document = builder.Document;

            Field field = builder.InsertField("=-1234567.89 \\# \"### ### ###.000\"", null);
            document.FieldOptions.ResultFormatter = new FieldResultFormatter("[{0:N2}]", null);

            field.Update();

            dataDir = dataDir + "FormatFieldResult_out.docx";
            builder.Document.Save(dataDir);
            // ExEnd:FormatFieldResult
            Console.WriteLine("\nFormat field result successfully.\nFile saved at " + dataDir);
        }

        // ExStart:FieldResultFormatter
        class FieldResultFormatter : IFieldResultFormatter
        {
            public FieldResultFormatter(string numberFormat, string dateFormat)
            {
                mNumberFormat = numberFormat;
                mDateFormat = dateFormat;
            }

            public FieldResultFormatter()
                : this(null, null)
            {
            }

            public override string FormatNumeric(double value, string format)
            {
                mNumberFormatInvocations.Add(new object[] { value, format });

                return string.IsNullOrEmpty(mNumberFormat)
                    ? null
                    : string.Format(mNumberFormat, value);
            }

            public override string FormatDateTime(DateTime value, string format, CalendarType calendarType)
            {
                mDateFormatInvocations.Add(new object[] { value, format, calendarType });

                return string.IsNullOrEmpty(mDateFormat)
                    ? null
                    : string.Format(mDateFormat, value);
            }

            public override string Format(string value, GeneralFormat format)
            {
                throw new NotImplementedException();
            }

            public override string Format(double value, GeneralFormat format)
            {
                throw new NotImplementedException();
            }

            private readonly string mNumberFormat;
            private readonly string mDateFormat;

            private readonly ArrayList mNumberFormatInvocations = new ArrayList();
            private readonly ArrayList mDateFormatInvocations = new ArrayList();
        }
        // ExEnd:FieldResultFormatter
    }
}
