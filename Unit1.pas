unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Grids, StdCtrls, ComCtrls,SimCard_Lib,inifiles, SPComm,XML_Lib,
  ExtCtrls, jpeg;

type
  TCMD = array of byte;
  PTCMD =^TCMD;
  
  TForm1 = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    AbbottData: TMemo;
    TabSheet3: TTabSheet;
    Label3: TLabel;
    Button2: TButton;
    TabSheet4: TTabSheet;
    Label4: TLabel;
    TabSheet5: TTabSheet;
    Label2: TLabel;
    Label1: TLabel;
    ID: TEdit;
    Label6: TLabel;
    COMPORT: TComboBox;
    MD: TStringGrid;
    Comm1: TComm;
    Image1: TImage;
    Button1: TButton;
    Image2: TImage;
    Label5: TLabel;
    DeviceSel: TComboBox;
    Image3: TImage;
    Progress: TProgressBar;
    ACCU_Log: TMemo;
    Image4: TImage;
    DbgModChk: TCheckBox;
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Comm1ReceiveData(Sender: TObject; Buffer: Pointer;
      BufferLength: Word);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure COMPORTChange(Sender: TObject);
    procedure DeviceSelChange(Sender: TObject);
    procedure DbgModChkClick(Sender: TObject);
  private
    { Private declarations }
  public
     procedure ShowXceedDataOld;
     procedure ShowXceedDataNew;
     procedure ShowFreeStyleData;
     procedure ShowACCUData;
  end;

var
  DataPath:string;
  ConfigINI:tinifile;
  Form1: TForm1;
  SIM:THCardReader;
  CloseFlag:boolean;
  FinishFlag:boolean;
  AbbottDataList:string;
  ShowAbbott:boolean;
  DebugMode:boolean;
  DeviceIndexOffset:integer; //若亞培有選，offset=0，若沒選，offset=2

  ACCUDataStr:string;
  ACCUDataBuf:array [0..255] of char;
  ACCUFinally:boolean;
  ACCUBufIndex:integer;
  ACCUDataCount:integer;
  ACCUDataList:TStringList;
  ACCUDataValue:string;
  ACCUDataTime:string;
  ACCUDataDate:string;
  ShowAccu:boolean;

  procedure ShowCard;
  procedure DeleteDataXceedOld(out COM:TComm);
  procedure DeleteDataXceedNew(out COM:TComm);
  procedure DeleteDataFreeStyle(out COM:TComm);
implementation

{$R *.dfm}

procedure Delay(ms:int64);
var counter:int64;
begin
   counter:=0;
   while (CloseFlag=false) and (counter<ms) do
   begin
      sleep(10);
      counter:=counter+10;
      application.ProcessMessages;
   end;
end;

procedure ShowCard;
begin
  form1.ID.Text:=SIM.ID_NO;
end;

procedure TForm1.FormCreate(Sender: TObject);
var tempindex:integer;
begin
   CloseFlag:=false;
   FinishFlag:=false;
   SIM:= THCardReader.Create(false,ShowCard);

   DataPath:=ExtractFilePath(Application.ExeName);

   ConfigINI:=tinifile.create(DataPath+'App.ini');

   COMPORT.ItemIndex:=ConfigINI.ReadInteger('Device','COMNUM',1)-1;
   ShowAbbott:=ConfigINI.ReadBool('Device','ShowAbbott',true);
   ShowAccu:=ConfigINI.ReadBool('Device','ShowAccu',true);
   DebugMode:=ConfigINI.ReadBool('Device','DebugMode',false);
   DeviceSel.Items.Clear;

   if ShowAbbott then
   begin
      DeviceSel.Items.Add('Xceed [New]');
      DeviceSel.Items.Add('Xceed [Old]');
      DeviceSel.Items.Add('FREEDOM Lite');
      DeviceIndexOffset:=0;
   end
   else DeviceIndexOffset:=3;

   if ShowAccu then
   begin
      DeviceSel.Items.Add('ACCU-CHEK');
   end;
   tempindex:=ConfigINI.ReadInteger('Device','DEVICE',1)-1;
   DeviceSel.ItemIndex:=0;
   if tempindex<DeviceSel.Items.Count then DeviceSel.ItemIndex:= tempindex;
   ConfigINI.WriteInteger('Device','DEVICE',DeviceSel.ItemIndex+1);
   ConfigINI.WriteInteger('Device','COMNUM',COMPORT.ItemIndex+1);
   ConfigINI.WriteBool('Device','ShowAbbott',ShowAbbott);
   ConfigINI.WriteBool('Device','ShowAccu',ShowAccu);
   ConfigINI.WriteBool('Device','DebugMode',DebugMode);
   ConfigINI.UpdateFile;

   DbgModChk.Visible:= DebugMode;
   Accu_log.Visible:= DebugMode;
   DeviceSelChange(nil);
end;

function CheckIdNo(sIdNo: string): Boolean;
const
    IDNIDX : array['A'..'Z'] of byte =
        (1,2,3,4,5,6,7,8,25,9,10,11,12,13,26,14,15,16,17,18,19,20,23,21,22,24);

    IDNTable : array[1..26] of byte = (10,11,12,13,14,15,16,17,18,19,20,21,22,
        23,24,25,26,27,28,29,30,31,32,33,34,35);
var V:integer;
begin
    result := False;
    if length(sIdNo)<>10 then exit;
    if sIdNo[1] in ['A'..'Z'] then
    begin
        V :=
        IDNTable[IDNIDX[sIdNo[1]]] div 10 + (IDNTable[IDNIDX[sIdNo[1]]] mod 10) * 9 +
        (byte(sIdNo[2])-48) * 8 + (byte(sIdNo[3])-48) * 7 + (byte(sIdNo[4])-48) * 6 +
        (byte(sIdNo[5])-48) * 5 + (byte(sIdNo[6])-48) * 4 + (byte(sIdNo[7])-48) * 3 +
        (byte(sIdNo[8])-48) * 2 + byte(sIdNo[9])-48 + byte(sIdNo[10])-48;
        result := (length(sIdNo) = 10) and ((sIdNo[2] = '1') or (sIdNo[2] = '2')) and (V
        div 10 = V / 10);
        //SHOWMESSAGE(BOOLTOSTR(result));
    end
end;

procedure TForm1.Button2Click(Sender: TObject);
var IDNO:string;
begin
  IDNO:=ID.Text;
  if CheckIdNO(IDNO)=true then
  begin
    SIM.SetSuspend:=true;
    PageControl1.TabIndex:=2;
  end
  else
  begin
    showmessage('身分證字號不合法,請檢查是否輸入錯誤');
  end;
end;

procedure ReadXceedOld();
  const CMD:array[0..16] of word=($02,$31,$47,$45,$54,$5F,$45,$56,$45,$4E,$54,$53,$03,$34,$38,$0D,$0A);
var
  Data:pansichar;
  i,j,k:integer;
  Com4Digi:string;
begin
  try
    form1.COMM1.WriteCommData(char($06),1);
    Delay(50);
    form1.COMM1.WriteCommData(char($04),1);
    Delay(50);
    form1.COMM1.WriteCommData(char($05),1);
    Delay(50);
    for k:=0 to 16 do
    begin
      Data:=@CMD[k];
      form1.COMM1.WriteCommData(Data,1);
      Delay(10);
    end;
    //DataLock:=0;
    Delay(100);
    form1.COMM1.WriteCommData(char($04),1);
    Delay(100);
    form1.COMM1.WriteCommData(char($06),1);
    Delay(100);
  except
    form1.COMM1.StopComm;
  end;
end;

procedure ReadXceedNew();
var
  Data:pansichar;
  i,j,k:integer;
  Com4Digi:string;
begin
  try
    form1.COMM1.WriteCommData('$xmem'#13#10,7);
  except
    form1.COMM1.StopComm;
  end;
end;

procedure ReadFreeStyle();
var
  Data:pansichar;
  i,j,k:integer;
  Com4Digi:string;
begin
  try
    form1.COMM1.WriteCommData('mem'#13#10,5);
  except
    form1.COMM1.StopComm;
  end;
end;

procedure WriteACCUCMD(PCMD:pointer;len:integer;delayms:integer);
var
  i:integer;
  CMD:TCMD;
  BUF:PAnsiChar;
  LastACCUBufIndex:integer;
  RetryCount:integer;
  os:integer;
begin
   ACCUBufIndex:=0;
   ACCUDataStr:='';
   for i:=0 to 255 do
   begin
      ACCUDataBuf[i]:=chr(0);
   end;
   for i:=0 to len-1 do
   begin
      RetryCount:=0;
      BUF:=PAnsiChar(PCMD)+i;
      LastACCUBufIndex:=ACCUBufIndex;
      form1.COMM1.WriteCommData(BUF,1);             //PAnsiChar(BUF)
      sleep(delayms);
      while (ACCUBufIndex=LastACCUBufIndex) and (RetryCount<10) do
      begin
        RetryCount:=RetryCount+1;
        sleep(delayms);
        application.ProcessMessages;
      end;
   end;
   form1.ACCU_Log.Lines.Add(ACCUDataStr);
   form1.ACCU_Log.Lines.Add(ACCUDataBuf);

   // 61 31 31 31 42 09 30 39 39 09 31 30 30 33 09 31 36 30 34 30 32 09 30 30 30 30 30 30 30 30 09 09 35 44
   if (ACCUBufIndex=34)and (ACCUDataBuf[0]=chr($61)) then        //只有<10筆
   begin
      ACCUDataValue:=inttostr((ord(ACCUDataBuf[6])-$30)*100 + (ord(ACCUDataBuf[7])-$30)*10 +(ord(ACCUDataBuf[8])-$30));
      ACCUDataTime:= ACCUDataBuf[10]+ACCUDataBuf[11]+':'+ACCUDataBuf[12]+ACCUDataBuf[13];
      ACCUDataDate:= ACCUDataBuf[15]+ACCUDataBuf[16]+'/'+ACCUDataBuf[17]+ACCUDataBuf[18]+'/'+ACCUDataBuf[19]+ACCUDataBuf[20];
      ACCUDataList.Add(ACCUDataDate+','+ACCUDataTime+','+ACCUDataValue);
      if ACCUDataList.Count=ACCUDataCount then ACCUFinally:=true;
   end;
{
   if (ACCUBufIndex=36)and (ACCUDataBuf[0]=chr($61)) then        //只有10~99筆
   begin
      os:=2;
      ACCUDataValue:=inttostr((ord(ACCUDataBuf[6+os])-$30)*100 + (ord(ACCUDataBuf[7+os])-$30)*10 +(ord(ACCUDataBuf[8+os])-$30));
      ACCUDataTime:= ACCUDataBuf[10+os]+ACCUDataBuf[11+os]+':'+ACCUDataBuf[12+os]+ACCUDataBuf[13+os];
      ACCUDataDate:= ACCUDataBuf[15+os]+ACCUDataBuf[16+os]+'/'+ACCUDataBuf[17+os]+ACCUDataBuf[18+os]+'/'+ACCUDataBuf[19+os]+ACCUDataBuf[20+os];
      ACCUDataList.Add(ACCUDataDate+','+ACCUDataTime+','+ACCUDataValue);
      if ACCUDataList.Count=ACCUDataCount then ACCUFinally:=true;
   end;
   if (ACCUBufIndex=38)and (ACCUDataBuf[0]=chr($61)) then        //只有100~999筆
   begin
      os:=4;
      ACCUDataValue:=inttostr((ord(ACCUDataBuf[6+os])-$30)*100 + (ord(ACCUDataBuf[7+os])-$30)*10 +(ord(ACCUDataBuf[8+os])-$30));
      ACCUDataTime:= ACCUDataBuf[10+os]+ACCUDataBuf[11+os]+':'+ACCUDataBuf[12+os]+ACCUDataBuf[13+os];
      ACCUDataDate:= ACCUDataBuf[15+os]+ACCUDataBuf[16+os]+'/'+ACCUDataBuf[17+os]+ACCUDataBuf[18+os]+'/'+ACCUDataBuf[19+os]+ACCUDataBuf[20+os];
      ACCUDataList.Add(ACCUDataDate+','+ACCUDataTime+','+ACCUDataValue);
      if ACCUDataList.Count=ACCUDataCount then ACCUFinally:=true;
   end;
}
   if ACCUFinally then
   begin
      //showmessage(ACCUDataList.DelimitedText);
      form1.ShowACCUData;
   end;
end;

function GenReadCmdforACCu(index:integer;out len:integer):TCMD;
var
  dig:array[0..2]of byte;
  offset:integer;
begin
   if (index>=0)and(index<10) then len:=7;
   if (index>=10)and(index<100) then len:=9;
   if (index>=100)and(index<1000) then len:=11;
   dig[0]:=index mod 10+$30;
   dig[1]:=(index div 10) mod 10+$30;
   dig[2]:=(index div 100) mod 10+$30;
   setlength(result,len);

   offset:=0;
   result[offset]:=$61;  offset:=offset+1;
   result[offset]:=$09;  offset:=offset+1;
   if index>=100 then
   begin
      result[offset]:=dig[2];  offset:=offset+1;
   end;
   if index>=10 then
   begin
      result[offset]:=dig[1];  offset:=offset+1;
   end;
   if index>=0 then
   begin
      result[offset]:=dig[0];  offset:=offset+1;
   end;
   result[offset]:=$09;  offset:=offset+1;
   if index>=100 then
   begin
      result[offset]:=dig[2];  offset:=offset+1;
   end;
   if index>=10 then
   begin
      result[offset]:=dig[1];  offset:=offset+1;
   end;
   if index>=0 then
   begin
      result[offset]:=dig[0];  offset:=offset+1;
   end;
   result[offset]:=$0D;  offset:=offset+1;
   result[offset]:=$06;
end;

procedure ReadACCU();
const
  CLEAR_STATUS      :array[0..2] of byte = ($0B,$0D,$06);
  READ_MEMORY_LEN   :array[0..2] of byte = ($60,$0D,$06);
  READ_MEMORY       :array[0..6] of byte = ($61,$09,$30,$09,$30,$0D,$06);

  SEND_BYTEDELAY    = 10;
  CMDDELAY          = 20;
var
  i,j:integer;
  READ_MEMORY_CNT   :array[0..6] of byte;
  len:integer;
  READ_CMD:TCMD;
begin
  form1.ACCU_Log.Lines.Clear;
  ACCUDataCount:=-1;
  ACCUFinally:=false;
  ACCUDataList:=TStringlist.Create;
  ACCUDataList.Delimiter:=#13;
  try
     form1.ACCU_Log.Lines.Add('------ 初始化 ------');
     WriteACCUCMD(@CLEAR_STATUS,3,SEND_BYTEDELAY);
     sleep(CMDDELAY);

     form1.ACCU_Log.Lines.Add('----- 讀取筆數 -----');
     WriteACCUCMD(@READ_MEMORY_LEN,3,SEND_BYTEDELAY);
     sleep(CMDDELAY);
     application.ProcessMessages;

    if ACCUDataCount=0 then
    begin
      form1.Progress.Position:=100;
      FinishFlag:=true;
      showmessage('血糖機目前無資料，請按 "OK" 回到上一頁');
      CloseFlag:=true;
      form1.close;
      exit;
    end;
    form1.ACCU_Log.Lines.Add('----- 讀取資料 -----');
    form1.Progress.Position:=0;
    form1.Progress.Max:=ACCUDataCount;
    for i:=1 to ACCUDataCount do
    begin
       form1.Progress.Position:=i;

       for j:=0 to 6 do
          READ_MEMORY_CNT[j]:= READ_MEMORY[j];
       READ_MEMORY_CNT[2]:=$30+i;
       READ_MEMORY_CNT[4]:=$30+i;
       WriteACCUCMD(@READ_MEMORY_CNT,7,SEND_BYTEDELAY);
       READ_CMD:=GenReadCmdforACCu(i,len);

       //showmessage(pchar(READ_CMD));
       //WriteACCUCMD(READ_CMD,len,SEND_BYTEDELAY);
       sleep(CMDDELAY);
    end;
  except
    form1.COMM1.StopComm;
  end;
end;


procedure TForm1.Button3Click(Sender: TObject);
var j:integer;
begin
  PageControl1.TabIndex:=3;
  AbbottData.Clear;
  AbbottData.Lines.Delimiter:=',';
  AbbottDataList:='';
  Progress.Position:=0;
  Progress.Max:=600;
  application.ProcessMessages;
  //MD.RowCount:=2;
  COMM1.StopComm;
  COMM1.CommName:='\\.\'+COMPORT.Items.Strings[COMPORT.ItemIndex];
  case DeviceSel.ItemIndex+DeviceIndexOffset of
    0: COMM1.BaudRate:=19200;
    1: COMM1.BaudRate:=9600;
    2: COMM1.BaudRate:=19200;
    3: COMM1.BaudRate:=9600;
  end;
  case DeviceSel.ItemIndex+DeviceIndexOffset of
    0: COMM1.ReadIntervalTimeout:=1000;
    1: COMM1.ReadIntervalTimeout:=1000;
    2: COMM1.ReadIntervalTimeout:=1000;
    3: COMM1.ReadIntervalTimeout:=5;
  end;
  try
    COMM1.StartComm;
  except
    showmessage('系統無此連接埠，請確定傳輸線接妥與連接埠設定正確');
    form1.Close;
  end;
  Delay(500);

  case DeviceSel.ItemIndex+DeviceIndexOffset of
    0:
      ReadXceedNew();
    1:
      ReadXceedOld();
    2:
      ReadFreeStyle();
    3:
      ReadACCU();
  end;

  for j:=0 to 600 do
  begin
    Delay(50);
    Progress.Position:=Progress.Position+1;
    application.ProcessMessages;
  end;

  if CloseFlag=true then  COMM1.StopComm
  else
  begin
    case DeviceSel.ItemIndex+DeviceIndexOffset of
    0,1,2:
      showmessage('無法讀取資料，請確認裝置正確連接並顯示出PC');
    3:
      showmessage('無法讀取資料，請確認血糖機是否開啟紅外線');
    end;
    CloseFlag:=true;
    form1.Close;
  end;

end;

procedure DeleteDataXceedNew(out COM:TComm);
begin
  COM.WriteCommData('$colz,1'#13#10,9);
  sleep(100);
  COM.WriteCommData('$colz,2'#13#10,9);
end;

procedure DeleteDataFreeStyle(out COM:TComm);
begin
  COM.WriteCommData('$colz,1'#13#10,9);
  sleep(100);
  COM.WriteCommData('$colz,2'#13#10,9);
end;

procedure DeleteDataACCU(out COM:TComm);
const
  CLEAR_STATUS      :array[0..2] of byte = ($0B,$0D,$06);
  CLEAR_MEMORY      :array[0..2] of byte = ($52,$0D,$06);
  SEND_BYTEDELAY    = 5;
  CMDDELAY          = 10;
begin
     WriteACCUCMD(@CLEAR_STATUS,3,SEND_BYTEDELAY);
     sleep(CMDDELAY);
     WriteACCUCMD(@CLEAR_MEMORY,3,SEND_BYTEDELAY);
     sleep(CMDDELAY);
end;

procedure DeleteDataXceedOld(out COM:TComm);
const CMD:array[0..9] of word=($02,$31,$21,$32,$45,$03,$43,$43,$0D,$0A);
var
  i:integer;
  Data:pansichar;
begin
  COM.WriteCommData(char($06),1);
  sleep(50);
  COM.WriteCommData(char($04),1);
  sleep(50);
  COM.WriteCommData(char($05),1);
  sleep(50);
  for i:=0 to 16 do
  begin
     Data:=@CMD[i];
     COM.WriteCommData(Data,1);
     sleep(10);
  end;
  sleep(50);
  COM.WriteCommData(char($04),1);
  sleep(50);
  COM.WriteCommData(char($06),1);
  sleep(50);
end;

procedure TForm1.Comm1ReceiveData(Sender: TObject; Buffer: Pointer;
  BufferLength: Word);
var
  BUF:pchar;
  TrimBUF:string;
  BINData:array of byte;
  TempStr:string;
  i:integer;
begin
   BINData:=Buffer;
   BUF:=pchar(BINData);
   TrimBUF:=BUF;
   TrimBUF:=trim(TrimBUF);
   BUF:=pchar(TrimBUF);
  case DeviceSel.ItemIndex+DeviceIndexOffset of
    0:
    begin
       AbbottDataList:=AbbottDataList+pchar(BINData);
       if pos('END', AbbottDataList)>0 then
       begin
          AbbottData.Lines.Text:=copy(AbbottDataList,52,pos('END',AbbottDataList)-59);
          Progress.Position:=500;
          application.ProcessMessages;
          ShowXceedDataNew;
          AbbottDataList:='';
       end;
    end;
    1:
    begin
      if (BufferLength<30)and(pos('END_OF_DATA', BUF)=0) then exit;
      //======================================讀取資料筆數
      if pos('END_OF_DATA', BUF)>0 then
      begin
        TComm(Sender).WriteCommData(char($06),1);
        ShowXceedDataOld;
      end
      else
      begin
        AbbottData.Lines.Add(BUF);
        TComm(Sender).WriteCommData(char($06),1);
      end;
    end;
    2:
    begin
       AbbottDataList:=AbbottDataList+pchar(BINData);
       if pos('END', AbbottDataList)>0 then
       begin
          AbbottData.Lines.Text:=copy(AbbottDataList,55,pos('END',AbbottDataList)-66);
          Progress.Position:=500;
          application.ProcessMessages;
          ShowFreeStyleData;
          AbbottDataList:='';
       end;
    end;
    3:
    begin
       for i:=0 to length(BUF)-1 do
       begin
          ACCUDataBuf[ACCUBufIndex]:=BUF[i];
          ACCUDataStr:=ACCUDataStr+inttohex(ord(ACCUDataBuf[ACCUBufIndex]),2)+' ';
          ACCUBufIndex:=ACCUBufIndex+1;
       end;
       // 60 30 35 09 30 30 35 09 35 42 資料筆數
       if (ACCUBufIndex=10)and (ACCUDataBuf[0]=chr($60)) then
       begin
          ACCUDataCount:=(ord(ACCUDataBuf[4])-$30)*100 + (ord(ACCUDataBuf[5])-$30)*10 +(ord(ACCUDataBuf[6])-$30);
          if ACCUDataCount=0 then
          begin
             Progress.Position:=500;
             application.ProcessMessages;
          end;
       end;
    end;
  end;
  application.ProcessMessages;
end;

function GetMonth(M:string):string;
const
  MonthStr:array[1..12] of string=('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var
  i:integer;
begin
  for i:=1 to 12 do
    if M=MonthStr[i] then
    begin
      result:=inttostr(i);
      if i<10 then result:='0'+result;
    end;
end;

procedure TForm1.ShowXceedDataNew;
var
 i:integer;
 DataCount:integer;
 rowIndex:integer;
 DateStr,TimeStr,yy,mm,dd:string;
 XML:XMLRecord;
 AbbottRow:TStringlist;
 GluValueStr:string;
 T:string;
begin
    rowIndex:=1;
    MD.Cells[0,0]:='日期';
    MD.Cells[1,0]:='時間';
    MD.Cells[2,0]:='血糖';
    AbbottRow:=TStringlist.Create;
    AbbottRow.Delimiter:=#13;
    XML.IDNO:=ID.Text;
    DataCount:=0;
    for i:=0 to AbbottData.Lines.Count-1 do
    begin
       T:=AbbottData.Lines.Strings[i];
       if (T[25]='G')and (T[29]='0')  then
       begin
          MD.RowCount:=rowIndex+1;
          //Memo2.Lines.Add(AbbottData.Lines.Strings[i]);
          yy:=copy(AbbottData.Lines.Strings[i],14,4);
          mm:=GetMonth(copy(AbbottData.Lines.Strings[i],6,3));
          dd:=copy(AbbottData.Lines.Strings[i],11,2);

          if (T[30]='0') then
          begin
           DateStr:=yy+'/'+mm+'/'+dd;
           TimeStr:=copy(AbbottData.Lines.Strings[i],19,5);
          end;
          if (T[30]='2') then      //當時間沒設定的時候..量測的時間應該要清空
          begin
            DateStr:='';
            TimeStr:='';
          end;
          MD.Cells[0,rowIndex]:=trim(DateStr);
          MD.Cells[1,rowIndex]:=trim(TimeStr);

          GluValueStr:=copy(AbbottData.Lines.Strings[i],0,3);
          if pos('LO',GluValueStr)>0 then
          begin
             MD.Cells[2,rowIndex]:='20';
          end
          else if pos('HI',GluValueStr)>0 then
          begin
             MD.Cells[2,rowIndex]:='500';
          end
          else
          begin
            MD.Cells[2,rowIndex]:=trim(inttostr(strtoint(GluValueStr)));
          end;
          AbbottRow.Add(MD.Cells[0,rowIndex]+','+MD.Cells[1,rowIndex]+','+MD.Cells[2,rowIndex]);
          //MeasureData.Rows[i].Strings[rowIndex]:=AbbottData.Lines.Strings[i];
          rowIndex:=rowIndex+1;
          DataCount:=DataCount+1;
       end;
    end;

    if DataCount=0 then  //(CreateXML(XML)='') and (AbbottData.Lines.Count=0)and
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        showmessage('血糖機目前無資料，請按 "OK" 回到上一頁');
        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end
    else
    begin
        if FinishFlag=false then
        begin
          FinishFlag:=true;
          if Application.MessageBox('血糖計資料讀取完成，請問要刪除血糖計資料嗎?','您好',mb_iconquestion+MB_YESNO)=IDYES then
          begin
            DeleteDataXceedNew(Comm1);
            showmessage('血糖計資料已刪除，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');
          end
          else
            showmessage('血糖計資料讀取完成，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');

          XML.MeasureData:=AbbottRow;
          CreateXML(XML);

          CloseFlag:=true;
          form1.close;
          exit;
        end;
    end;

end;

procedure TForm1.ShowFreeStyleData;
var
 i:integer;
 DataCount:integer;
 rowIndex:integer;
 TimeStr,yy,mm,dd:string;
 XML:XMLRecord;
 AbbottRow:TStringlist;
 T,GluValueStr:string;
begin
    rowIndex:=1;
    DataCount:=0;
    MD.Cells[0,0]:='日期';
    MD.Cells[1,0]:='時間';
    MD.Cells[2,0]:='血糖';
    AbbottRow:=TStringlist.Create;
    AbbottRow.Delimiter:=#13;
    XML.IDNO:=ID.Text;
    for i:=0 to AbbottData.Lines.Count-1 do
    begin
       T:=AbbottData.Lines.Strings[i];
       if (T[25]='0')and (T[29]='x')  then
       begin
          MD.RowCount:=rowIndex+1;
          //Memo2.Lines.Add(AbbottData.Lines.Strings[i]);
          yy:=copy(AbbottData.Lines.Strings[i],14,4);
          mm:=GetMonth(copy(AbbottData.Lines.Strings[i],6,3));
          dd:=copy(AbbottData.Lines.Strings[i],11,2);
          TimeStr:=yy+'/'+mm+'/'+dd;
          MD.Cells[0,rowIndex]:=trim(TimeStr);
          MD.Cells[1,rowIndex]:=trim(copy(AbbottData.Lines.Strings[i],19,5));

          GluValueStr:=copy(AbbottData.Lines.Strings[i],0,3);
          if pos('LO',GluValueStr)>0 then
          begin
             MD.Cells[2,rowIndex]:='20';
          end
          else if pos('HI',GluValueStr)>0 then
          begin
             MD.Cells[2,rowIndex]:='500';
          end
          else
          begin
            MD.Cells[2,rowIndex]:=trim(inttostr(strtoint(GluValueStr)));
          end;

          AbbottRow.Add(MD.Cells[0,rowIndex]+','+MD.Cells[1,rowIndex]+','+MD.Cells[2,rowIndex]);
          //MeasureData.Rows[i].Strings[rowIndex]:=AbbottData.Lines.Strings[i];
          rowIndex:=rowIndex+1;
          DataCount:=DataCount+1;
       end;
    end;

    if DataCount=0 then  //(CreateXML(XML)='') and (AbbottData.Lines.Count=0)and
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        showmessage('血糖機目前無資料，請按 "OK" 回到上一頁');
        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end
    else
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        if Application.MessageBox('血糖計資料讀取完成，請問要刪除血糖計資料嗎?','您好',mb_iconquestion+MB_YESNO)=IDYES then
        begin
          DeleteDataFreeStyle(Comm1);
          showmessage('血糖計資料已刪除，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');
        end
        else
          showmessage('血糖計資料讀取完成，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');

        XML.MeasureData:=AbbottRow;
        CreateXML(XML);

        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end;

end;


procedure TForm1.ShowXceedDataOld;
var
 i:integer;
 rowIndex:integer;
 TimeStr,yy,mm,dd:string;
 XML:XMLRecord;
 AbbottRow:TStringlist;
 DataCount:integer;
begin
    rowIndex:=1;
    MD.Cells[0,0]:='日期';
    MD.Cells[1,0]:='時間';
    MD.Cells[2,0]:='血糖';
    AbbottRow:=TStringlist.Create;
    AbbottRow.Delimiter:=#13;
    XML.IDNO:=ID.Text;
    DataCount:=0;
    for i:=0 to AbbottData.Lines.Count-1 do
    begin
       if (AbbottData.Lines.Strings[i][3]='1')and (AbbottData.Lines.Strings[i][27]<>'1')  then
       begin
          MD.RowCount:=rowIndex+1;
          //Memo2.Lines.Add(AbbottData.Lines.Strings[i]);
          yy:=copy(AbbottData.Lines.Strings[i],7,4);
          mm:=copy(AbbottData.Lines.Strings[i],11,2);
          dd:=copy(AbbottData.Lines.Strings[i],13,2);
          TimeStr:=yy+'/'+mm+'/'+dd;

          if AbbottData.Lines.Strings[i][5]='1' then
          begin
            MD.Cells[0,rowIndex]:=trim(TimeStr);
            MD.Cells[1,rowIndex]:=trim(copy(AbbottData.Lines.Strings[i],16,6));
          end
          else
          begin
            MD.Cells[0,rowIndex]:='';
            MD.Cells[1,rowIndex]:='';
          end;
          MD.Cells[2,rowIndex]:=trim(inttostr(strtoint(copy(AbbottData.Lines.Strings[i],22,5))));

          AbbottRow.Add(MD.Cells[0,rowIndex]+','+MD.Cells[1,rowIndex]+','+MD.Cells[2,rowIndex]);
          //MeasureData.Rows[i].Strings[rowIndex]:=AbbottData.Lines.Strings[i];
          rowIndex:=rowIndex+1;
          DataCount:=DataCount+1;
       end;
    end;

    if DataCount=0 then  //(CreateXML(XML)='') and (AbbottData.Lines.Count=0)and
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        showmessage('血糖機目前無資料，請按 "OK" 回到上一頁');
        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end
    else
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        if Application.MessageBox('血糖計資料讀取完成，請問要刪除血糖計資料嗎?','您好',mb_iconquestion+MB_YESNO)=IDYES then
        begin
          DeleteDataXceedOld(Comm1);
          showmessage('血糖計資料已刪除，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');
        end
        else
          showmessage('血糖計資料讀取完成，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');

        XML.MeasureData:=AbbottRow;
        CreateXML(XML);

        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end;

end;

procedure TForm1.ShowACCUData;
var
 i:integer;
 DataCount:integer;
 rowIndex:integer;
 DateStr,TimeStr,yy,mm,dd:string;
 XML:XMLRecord;
 ACCURow:TStringlist;
 GluValueStr:string;
 T:string;
begin
    rowIndex:=1;
    MD.Cells[0,0]:='日期';
    MD.Cells[1,0]:='時間';
    MD.Cells[2,0]:='血糖';
    ACCURow:=TStringlist.Create;
    ACCURow.Delimiter:=',';
    XML.IDNO:=ID.Text;
    DataCount:=0;
    for i:=0 to ACCUDatalist.Count-1 do
    begin
       ACCURow.DelimitedText:=ACCUDatalist.Strings[i];
       MD.RowCount:=rowIndex+1;
       MD.Cells[0,rowIndex]:=trim(ACCURow[0]);
       MD.Cells[1,rowIndex]:=trim(ACCURow[1]);
       MD.Cells[2,rowIndex]:=trim(ACCURow[2]);
       rowIndex:=rowIndex+1;
       DataCount:=DataCount+1;
    end;

    if DataCount=0 then  //(CreateXML(XML)='') and (AbbottData.Lines.Count=0)and
    begin
      if FinishFlag=false then
      begin
        FinishFlag:=true;
        showmessage('血糖機目前無資料，請按 "OK" 回到上一頁');
        CloseFlag:=true;
        form1.close;
        exit;
      end;
    end
    else
    begin
        if FinishFlag=false then
        begin
          FinishFlag:=true;
          if Application.MessageBox('血糖計資料讀取完成，請問要刪除血糖計資料嗎?','您好',mb_iconquestion+MB_YESNO)=IDYES then
          begin
            DeleteDataACCU(Comm1);
            showmessage('血糖計資料已刪除，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');
          end
          else
            showmessage('血糖計資料讀取完成，請按 "OK" 回到上一頁'+#13+'，並按下"資料列表"做上傳動作');

          XML.MeasureData:=ACCUDatalist;
          CreateXML(XML);

          CloseFlag:=true;
          if DebugMode=false then form1.close;
          exit;
        end;
    end;

end;

procedure TForm1.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
  CloseFlag:=true;
  SIM.SetSuspend:=true;
  canclose:=true;
end;

procedure TForm1.COMPORTChange(Sender: TObject);
begin
   DataPath:=ExtractFilePath(Application.ExeName);
   ConfigINI:=tinifile.create(DataPath+'App.ini');
   ConfigINI.WriteInteger('Device','COMNUM',COMPORT.ItemIndex+1);
   ConfigINI.UpdateFile;
end;

procedure TForm1.DeviceSelChange(Sender: TObject);
begin
   DataPath:=ExtractFilePath(Application.ExeName);
   ConfigINI:=tinifile.create(DataPath+'App.ini');
   ConfigINI.WriteInteger('Device','DEVICE',DeviceSel.ItemIndex+1);
   ConfigINI.UpdateFile;
  if DeviceSel.ItemIndex+DeviceIndexOffset=2 then
  begin
    Image2.Visible:=false;
    Image3.Visible:=true;
    Image4.Visible:=false;
  end
  else if DeviceSel.ItemIndex+DeviceIndexOffset=3 then
  begin
    Image2.Visible:=false;
    Image3.Visible:=false;
    Image4.Visible:=true;
  end
  else
  begin
    Image2.Visible:=true;
    Image3.Visible:=false;
    Image4.Visible:=false;
  end;
end;

procedure TForm1.DbgModChkClick(Sender: TObject);
begin
  ACCU_Log.Visible:= DbgModChk.Checked;
end;

end.
