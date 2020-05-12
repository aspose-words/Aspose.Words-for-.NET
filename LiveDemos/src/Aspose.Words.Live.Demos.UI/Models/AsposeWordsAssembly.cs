using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;
using Aspose.Words.Live.Demos.UI.Models;
using Aspose.Words;
using Aspose.Words.Reporting;
using Aspose.Words.Tables;
using Microsoft.VisualBasic.FileIO;
using Aspose.Cells;
using Newtonsoft.Json;

using File = System.IO.File;
using Row = Aspose.Words.Tables.Row;


namespace Aspose.Words.Live.Demos.UI.Models
{
	/// <summary>
	///AsposeWordsAssembly class to assemble word document
	/// </summary>
	public class AsposeWordsAssembly : AsposeWordsBase
  {
		/// <summary>
		/// Assemble method to assemmble word document
		/// </summary>
		public Response Assemble(string folderName, string templateFilename, string datasourceFilename,
      string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
    {
      Opts.AppName = "Assembly";
      Opts.MethodName = System.Reflection.MethodBase.GetCurrentMethod().Name;
      Opts.FolderName = folderName;
      Opts.FileName = datasourceFilename;
      Opts.OutputType = Path.GetExtension(templateFilename);
      Opts.ResultFileName = Path.GetFileNameWithoutExtension(templateFilename) + " Assembled";
      Opts.DeleteSourceFolder = false;

      var dataTable = PrepareDataTable(Opts.WorkingFileName, datasourceName, datasourceTableIndex, delimiter);

      if (dataTable != null)
        return  Process((inFilePath, outPath, zipOutFolder) =>
        {
          var doc = new Document(Path.Combine(Config.Configuration.WorkingDirectory, folderName, templateFilename));
          var engine = new ReportingEngine {Options = ReportBuildOptions.AllowMissingMembers};

          engine.BuildReport(doc, dataTable, datasourceName);
          doc.Save(outPath);
        });

      return new Response
      {
        FileName = null,
        FolderName = folderName,
        Status = "Can't process the data source",
        StatusCode = 500,
		FileProcessingErrorCode = FileProcessingErrorCode.OK
      };
    }

    private static DataTable PrepareDataTable(string filename, string datasourceName, int datasourceTableIndex = 0, string delimiter = ",")
    {
      try
      {
        switch (Path.GetExtension(filename).ToLower())
        {
          case ".json":
            return PrepareDataTableFromJson(filename, datasourceName);
          case ".xml":
            return PrepareDataTableFromXML(filename, datasourceName);
          case ".csv":
            return PrepareDataTableFromCSV(filename, datasourceName, delimiter);
          case ".xls":
          case ".xlsx":
            return PrepareDataTableFromExcel(filename, datasourceTableIndex);
          default:
            return PrepareDataTableFromDocument(filename, datasourceName, datasourceTableIndex);
        }
      }
      catch (Exception ex)
      {
				Console.WriteLine(ex.Message);
        return null;
      }
    }

    private static DataTable PrepareDataTableFromDocument(string filename, string datasourceName, int datasourceTableIndex)
    {
      var data = new Document(filename);

      var table = (Table)data.GetChild(NodeType.Table, datasourceTableIndex, true);
      if (table == null)
        return null;
      var properties = table.FirstRow.Cells.Select(x => x.GetText().Replace("\a", "")).ToArray();

      var dataTable = new DataTable(datasourceName);
      foreach (var property in properties)
        if (!dataTable.Columns.Contains(property))
          dataTable.Columns.Add(property);

      foreach (var row in table.Rows.Skip(1).Select(x => (Row)x))
        dataTable.Rows.Add(row.Cells.Select(x => x.GetText().Replace("\a", "") as object).ToArray()); // Cells have special symbol '\a'

      return dataTable;
    }

    private static DataTable PrepareDataTableFromJson(string filename, string datasourceName)
    {
      var xml = JsonConvert.DeserializeXmlNode(File.ReadAllText(filename), "RootElement");
      var dataSet = new DataSet(datasourceName);
      dataSet.ReadXml(new MemoryStream(Encoding.UTF8.GetBytes(xml.InnerXml)));
      return dataSet.Tables[datasourceName];
    }

    private static DataTable PrepareDataTableFromXML(string filename, string datasourceName)
    {
      var dataSet = new DataSet(datasourceName);
      dataSet.ReadXml(XmlReader.Create(filename));
      return dataSet.Tables[datasourceName];
    }

    private static DataTable PrepareDataTableFromCSV(string filename, string datasourceName, string delimiter)
    {
      var dataTable = new DataTable(datasourceName);
      using (var csvReader = new TextFieldParser(filename))
      {
        csvReader.SetDelimiters(delimiter);
        csvReader.HasFieldsEnclosedInQuotes = true;
        var colFields = csvReader.ReadFields();
        foreach (var column in colFields)
        {
          var datecolumn = new DataColumn(column);
          datecolumn.AllowDBNull = true;
          dataTable.Columns.Add(datecolumn);
        }
        while (!csvReader.EndOfData)
        {
          var fieldData = csvReader.ReadFields();
          //Making empty value as null
          for (int i = 0; i < fieldData.Length; i++)
            if (fieldData[i] == "")
              fieldData[i] = null;
          dataTable.Rows.Add(fieldData);
        }
      }
      return dataTable;
    }

    private static DataTable PrepareDataTableFromExcel(string filename, int datasourceTableIndex)
    {
      
      var excel = new Workbook(filename);
      var cells = excel.Worksheets[datasourceTableIndex].Cells;
      var lastColumn = cells.MaxColumn;
      var lastRow = int.MinValue;
      for (int i = 0; i < lastColumn; i++)
        if (cells.GetLastDataRow(i) > lastRow)
          lastRow = cells.GetLastDataRow(i);
      if (lastRow == int.MinValue)
        return null;
      return cells.ExportDataTable(0, 0, lastRow + 1, lastColumn + 1, true);
    }
  }
}
