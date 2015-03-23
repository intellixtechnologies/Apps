unit Main;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls,OleAuto;

type
  TForm1 = class(TForm)
    Edit1: TEdit;
    Edit2: TEdit;
    Button1: TButton;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
const
wdExportFormatPDF = 17;
var
  Word, PDF: OleVariant;
begin
  Showmessage('Test');
  Word:= CreateOLEObject('Word.Application');
  PDF:= Word.Documents.Open(Edit1.Text);
  PDF.ExportAsFixedFormat(Edit2.Text,wdExportFormatPDF);
end;

end.
