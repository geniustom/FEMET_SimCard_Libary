unit SimCard_Lib;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, PCSCConnector;


type
  THCardRecord =record
    CardNo: string;
    HolderName: string;
    IDN: string;
    BirthDate: string;
    Sex: string;
    IssueDate: string;
  end;

  TCallBack = procedure;
  THCardReader= class(TThread)
  public
    SetSuspend:boolean;
    CardRecord:THCardRecord;
    PCSC: TPCSCConnector;
    CardIsRemove:boolean;
    CardIsOK:boolean;
    ID_NO:string;
    ReaderIsOK:boolean;
    InsertState:integer; //0拔1插
    LastCardInsert:boolean;
    procedure ReadCard();
    constructor Create(CreateSuspended:Boolean;CB:TCallBack);
  private
    CallBack:TCallBack;
  protected
    function GetHCardRecord(Data: string): THCardRecord;
    procedure ClearRecord();
    procedure PCSCConnector1StateChange(Sender: TObject;StateChange: TPCSCStateChange; StateChangeStr: String);
    procedure PCSCConnector1Error(Sender: TObject; ProcSource: TProcSource;ErrCode: Integer);
    procedure Execute; override;

  end;




implementation

constructor THCardReader.Create(CreateSuspended:Boolean;CB:TCallBack);
begin
  inherited Create(CreateSuspended);
  CallBack:=CB;
end;

procedure THCardReader.Execute;
var i:integer;
begin
  self.Priority:= tpLower;
  PCSC:=TPCSCConnector.Create(nil);
  PCSC.AutoConnect:=true;
  PCSC.Protocol:=pcpT1;
  PCSC.OnStateChange:= PCSCConnector1StateChange;
  PCSC.OnError:=PCSCConnector1Error;

  while true do
  begin
    if SetSuspend then self.Suspend;
    PCSC.Refresh;
    PCSC.ResetAPI;
    //if ReaderIsOK=true then
    //begin
      CardIsRemove:=PCSC.ReaderPadding;
      CardIsOK:=PCSC.CardPresent;
    //end
    //else
    //begin
    //  CardIsRemove:=true;
    //  CardIsOK:=false;
    //end;
    application.ProcessMessages;
    ReadCard();
    for i:=0 to 9 do
    begin
      sleep(50);
      if SetSuspend then self.Suspend;
      application.ProcessMessages;
    end;
  end;
end;


procedure THCardReader.ReadCard();
var
  NowCardInsert:boolean;
  i:integer;
begin
  if PCSC.ReaderList.Count =0 then exit;
  if PCSC.ReaderList.Count > 0 then
  begin
    PCSC.ReaderIndex := 0;
    PCSC.Connect;
    if PCSC.Connected then
    begin
      //ReaderIsOK:=true;
      //PCSC.Disconnect;
    end;
    
    PCSC.GetResponseFromCard('00 A4 04 00 10 D1 58 00 00 01 00 00 00 00 00 00 00 00 00 11 00');
    self.CardRecord:=GetHCardRecord(PCSC.GetResponseFromCard('00 CA 11 00 02 00 00'));
    if CardRecord.HolderName='' then CardRecord.BirthDate:='';  //避免出現1912/1/1
    if CardRecord.IDN<>'' then
      ID_NO:=CardRecord.IDN;           //身分證有讀到才更新

  end;
  
  NowCardInsert:=(CardRecord.HolderName<>'');

  if (NowCardInsert=true) and  (LastCardInsert=false) then //插卡
  begin
    Beep;
    InsertState:=1;
    if Assigned(CallBack) then CallBack;
  end;

  if (NowCardInsert=false) and  (LastCardInsert=true) then //拔卡
  begin
    Beep;
    InsertState:=0;
    if Assigned(CallBack) then CallBack;
    ClearRecord;
  end;
  LastCardInsert:=NowCardInsert;
end;


procedure THCardReader.PCSCConnector1Error(Sender: TObject;ProcSource: TProcSource; ErrCode: Integer);
begin

  //if ErrCode<0 then
  if (ErrCode=-2146435026) or (ErrCode=-2146435063) then
  begin
    ReaderIsOK:=false;
    CardIsRemove:=false;
    ClearRecord;
  end
  else
  begin
    if (CardIsRemove=false) and (CardIsOK=false) then
    begin
      ReaderIsOK:=false;
    end
    else
    begin
      ReaderIsOK:=(CardIsRemove or CardIsOK);
    end;
  end;

end;

procedure THCardReader.PCSCConnector1StateChange(Sender: TObject;StateChange: TPCSCStateChange; StateChangeStr: String);
begin
{
  if StateChangeStr='Card Insert' then
  begin

  end;

  if (StateChangeStr='Card Remove') and (CardRecord.HolderName<>'') then
  begin

  end;
}
end;

function THCardReader.GetHCardRecord(Data: string): THCardRecord;
begin
  Data := Hex2Bin(Data);
  Result.CardNo := Copy(Data, 1, 12);
  Result.HolderName := Copy(Data, 13, 6);
  Result.IDN := Copy(Data, 33, 10);
  Result.BirthDate := DatetimeTostr(EncodeDate(StrToIntDef(Copy(Data, 44, 2), 1) + 1911, StrToIntDef(Copy(Data, 46, 2), 1), StrToIntDef(Copy(Data, 48, 2), 1)));
  Result.Sex := Copy(Data, 50, 1);
  Result.IssueDate := DatetimeTostr(EncodeDate(StrToIntDef(Copy(Data, 52, 2), 1) + 1911, StrToIntDef(Copy(Data, 54, 2), 1), StrToIntDef(Copy(Data, 56, 2), 1)));
end;

procedure THCardReader.ClearRecord();
begin
  CardRecord.CardNo := '';
  CardRecord.HolderName := '';
  CardRecord.IDN := '';
  CardRecord.BirthDate := '';
  CardRecord.Sex := '';
  CardRecord.IssueDate := '';
end;

end.
