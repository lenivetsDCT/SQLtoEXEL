unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  sqldb, odbcconn, Forms, StdCtrls, ComObj, Variants;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    Edit2: TEdit;
    Edit3: TEdit;
    Edit4: TEdit;
    ODBCConnection1: TODBCConnection;
    SQLQuery1: TSQLQuery;
    SQLTransaction1: TSQLTransaction;
    procedure Button1Click(Sender: TObject);
  private
    { private declarations }
  public
    { public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
var
  ExcelApp, Workbook, Range, Cell1, Cell2, ArrayData: variant;
  BeginCol, BeginRow, i, j: integer;
begin
  ODBCConnection1.HostName := Edit1.Text;
  ODBCConnection1.Driver := Edit2.Text;
  ODBCCOnnection1.DatabaseName := Edit3.Text;
  SQLQuery1.ReadOnly := True;
  ODBCConnection1.Connected := True;
  ODBCConnection1.KeepConnection := True;
  SQLTransaction1.Active := True;

  SQLQuery1.SQL.add(Edit4.Text);
  SQLQuery1.Open;

  BeginCol := 1;
  BeginRow := 1;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Application.EnableEvents := False;
  //  Создаем Книгу (Workbook)
  // Если заполняем шаблон, то
  // Workbook := ExcelApp.WorkBooks.Add('C:\MyTemplate.xls');
  Workbook := ExcelApp.WorkBooks.Add;
  with SQLQuery1 do
  begin
    Last;
    First;
    ArrayData := VarArrayCreate([1, RecordCount, 1, FieldCount], varVariant);
    for I := 1 to RecordCount do
    begin
      for J := 1 to FieldCount do
        ArrayData[I, J] := Fields[j - 1].AsVariant;
      Next;
    end;
  end;
  Cell1 := WorkBook.WorkSheets[1].Cells[BeginRow, BeginCol];
  Cell2 := WorkBook.WorkSheets[1].Cells[BeginRow + SQLQuery1.RecordCount -
    1, BeginCol + SQLQuery1.FieldCount - 1];
  Range := WorkBook.WorkSheets[1].Range[Cell1, Cell2];
  Range.Value := ArrayData;
  ExcelApp.Visible := True;
  SQLQuery1.Close;
  ODBCConnection1.Connected := False;
  ArrayData.Free;
end;

end.
