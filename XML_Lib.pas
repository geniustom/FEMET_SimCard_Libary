unit XML_Lib;

interface
uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

  
type
  XMLRecord=record
     IDNO:string;
     MeasureData:TStringlist;
  end;

  function CreateXML(XML:XMLRecord):string;

implementation


function CreateXML(XML:XMLRecord):string;
var
  XMLFile:Textfile;
  SelfPath:string;
  DatePath:string;
  TimeStr:string;
  Path:string;
  i:integer;
begin
  result:='';
  if XML.MeasureData.Count=0 then exit;

  SelfPath:=ExtractFileDir(application.ExeName)+'\Data';
  DatePath:=formatdatetime('YY_MM_DD',now);
  //TimeStr:=formatdatetime('HH_NN_SS',now);
  //Path:=SelfPath+'\'+DatePath+'\'+TimeStr+'.csv';
  //result:='Data\'+DatePath+'\'+TimeStr+'.csv';
  TimeStr:=formatdatetime('YYYYMMDDHHNNSS',now);
  Path:=SelfPath+'\'+TimeStr+'.csv';
  result:='Data\'+TimeStr+'.csv';

  if DirectoryExists(SelfPath)=false then
    CreateDir(SelfPath);

  //if DirectoryExists(SelfPath+'\'+DatePath)=false then
  //  CreateDir(SelfPath+'\'+DatePath);

   assignFile(XMLFile,Path);
   rewrite(XMLFile);
  XML.MeasureData.Delimiter:=',';
  for i:=0 to XML.MeasureData.Count-1 do
  begin
     writeln(XMLFile,XML.IDNO+','+XML.MeasureData.Strings[i]);
  end;
   closeFile(XMLFile);
end;


end.
