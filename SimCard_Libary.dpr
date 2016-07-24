program SimCard_Libary;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  SimCard_Lib in 'SimCard_Lib.pas',
  XML_Lib in 'XML_Lib.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.Title := '血糖計資料讀取';
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
