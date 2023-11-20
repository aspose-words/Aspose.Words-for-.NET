program Aspose.Words.Examples;

uses
  System.SysUtils,
  System.Win.ComObj,
  Aspose_Words_TLB in '..\22.0\Imports\Aspose_Words_TLB.pas';

procedure ConvertDocument();
var
  helper, doc: OLEvariant;
begin
  helper:= CreateOleObject('Aspose.Words.ComHelper');
  doc := helper.Open('X:\Aspose.Words-for-.NET.git\Examples\Data\Absolute position tab.docx');
  doc.Save('X:\Aspose.Words-for-.NET.git\Examples\Data\Absolute position tab.pdf');
end;

procedure CreateNewDocument();
var
  builder: OLEvariant;
begin
  builder := CreateOleObject('Aspose.Words.DocumentBuilder');
  builder.Writeln('Hello world!');
  builder.Document.Save('X:\Aspose.Words-for-.NET.git\Examples\Data\New document.docx');
end;


var
  license : _License;
begin
  Set8087CW($133f);
  try
    license := CoLicense.Create;
    license.SetLicense('Use your license');

    ConvertDocument();
    CreateNewDocument();
  except
    on E: Exception do
      Writeln(E.ClassName, ': ', E.Message);
  end;
end.


