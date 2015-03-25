unit Main;

interface

uses
Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
vcl.Dialogs, Buttons, StdCtrls, ComObj, WordXP, ActiveX;

type
TmainFrm = class(TForm)
Edit1: TEdit;
SpeedButton1: TSpeedButton;
Label1: TLabel;
Button1: TButton;
Label2: TLabel;
Edit2: TEdit;
OpenDialog1: TOpenDialog;
SaveDialog1: TSaveDialog;
SpeedButton2: TSpeedButton;
procedure SpeedButton1Click(Sender: TObject);
procedure SpeedButton2Click(Sender: TObject);
procedure Button1Click(Sender: TObject);
private
{ Private declarations }
public
{ Public declarations }
end;

var
mainFrm: TmainFrm;

implementation

{$R *.dfm}

procedure TmainFrm.Button1Click(Sender: TObject);
const
wdExportFormatPDF = 17;
ppFixedFormatTypePDF = 32;
msoFalse = TOleEnum(False);
msoTrue = TOleEnum(True);
var
Word, Doc: OleVariant;
oWB : Variant;
begin
if (ExtractFileExt(Edit1.Text) = '.doc') OR (ExtractFileExt(Edit1.Text) = '.docx') then
begin
Word := CreateOLEObject('Word.Application');
Doc := Word.Documents.Open(Edit1.Text);
Doc.ExportAsFixedFormat(Edit2.Text,wdExportFormatPDF);
Word.quit;
end
else if (ExtractFileExt(Edit1.Text) = '.xls') OR (ExtractFileExt(Edit1.Text) = '.xlsx') then
begin
Word := CreateOLEObject('Excel.Application');
Doc := Word.WorkBooks.Open(Edit1.Text);
oWB := Doc.ActiveSheet;
oWB.ExportAsFixedFormat(0,Edit2.Text);
Word.quit;
end
else if (ExtractFileExt(Edit1.Text) = '.ppt') OR (ExtractFileExt(Edit1.Text) = '.pptx') then
begin
Word := CreateOLEObject('PowerPoint.Application');
Word.Visible := True;
Doc := Word.Presentations.Open(Edit1.Text,msoFalse,msoFalse,msoTrue);
Doc.SaveAs(Edit2.Text,ppFixedFormatTypePDF);
Doc.Close;
Word.quit;
end
else
begin
ShowMessage('File Type Not supported. Please Select .doc, .docx, .xls, .xlsx, .ppt');
exit;
end;
ShowMessage('Process Completed' );
oWB := UnAssigned;
Doc := UnAssigned;
Word := Unassigned;
end;

procedure TmainFrm.SpeedButton1Click(Sender: TObject);
begin
if OpenDialog1.Execute then
Edit1.Text := OpenDialog1.FileName;
end;

procedure TmainFrm.SpeedButton2Click(Sender: TObject);
begin
if SaveDialog1.Execute then
Edit2.Text := SaveDialog1.FileName;
end;

end.
