unit main;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Menus, ComCtrls, OverbyteIcsEmulVT, OverbyteIcsTnEmulVT,
  OverbyteIcsTnScript, StdCtrls, ExtCtrls, osutil, Grids, ValEdit, Math,
  Buttons, ToolWin, ActnMan, ActnCtrls, ActnMenus, XPStyleActnCtrls,
  ActnList, StdActns, uLkJSON, GIFImage;

type
  TForm1 = class(TForm)
    pgcMain: TPageControl;
    tsInfo: TTabSheet;
    tnOrionScreen: TTnScript;
    tsScriptEdit: TTabSheet;
    btnConnect: TButton;
    btnFunc1: TButton;
    btnFunc2: TButton;
    btnFunc3: TButton;
    btnFunc4: TButton;
    btnFunc5: TButton;
    lblX: TLabel;
    lblY: TLabel;
    pnl1: TPanel;
    tsLogging: TTabSheet;
    mmoDebugOutput: TMemo;
    tmrCursorPos: TTimer;
    tmrMainTimer: TTimer;
    tmrScreenChecking: TTimer;
    tmrLineChecking: TTimer;
    tmrChecking: TTimer;
    tmrDelay: TTimer;
    dlgSaveMacro: TSaveDialog;
    dlgOpenMacro: TOpenDialog;
    dlgSaveScreenshot: TSaveDialog;
    dlgOpenData: TOpenDialog;
    dlgSaveLog: TSaveDialog;
    rlProcessResults: TValueListEditor;
    lstImportData: TListBox;
    rlOrionMacro: TValueListEditor;
    mmoScriptEdit: TMemo;
    btnMoveRight: TSpeedButton;
    btnMoveLeft: TSpeedButton;
    btnResListClear: TButton;
    btnSaveResList: TButton;
    mmMainMenu: TMainMenu;
    File1: TMenuItem;
    Edit1: TMenuItem;
    Help1: TMenuItem;
    Settings1: TMenuItem;
    Open1: TMenuItem;
    Openscript1: TMenuItem;
    N1: TMenuItem;
    Exit1: TMenuItem;
    N2: TMenuItem;
    SaveScript1: TMenuItem;
    Contents1: TMenuItem;
    N3: TMenuItem;
    AboutOrionStudio1: TMenuItem;
    btnClearLog: TButton;
    btnSaveLog: TButton;
    lblDebugX: TLabel;
    btnTest: TButton;
    btnPlayMacro: TBitBtn;
    btnStopMacro: TBitBtn;
    dlgSaveMMOLog: TSaveDialog;
    btnTestButton: TButton;
    procedure FormCreate(Sender: TObject);
    procedure tmrCursorPosTimer(Sender: TObject);
    procedure btnConnectClick(Sender: TObject);
    procedure tnOrionScreenSessionConnected(Sender: TObject);
    procedure tnOrionScreenSessionClosed(Sender: TObject);
    procedure btnFunc1Click(Sender: TObject);
    procedure btnFunc2Click(Sender: TObject);
    procedure btnFunc3Click(Sender: TObject);
    procedure btnFunc5Click(Sender: TObject);
    procedure btnFunc4Click(Sender: TObject);
    procedure Exit1Click(Sender: TObject);
    procedure btnClearLogClick(Sender: TObject);
    procedure btnSaveLogClick(Sender: TObject);
    procedure tmrMainTimerTimer(Sender: TObject);
    procedure tmrScreenCheckingTimer(Sender: TObject);
    procedure tmrLineCheckingTimer(Sender: TObject);
    procedure tmrCheckingTimer(Sender: TObject);
    procedure tmrDelayTimer(Sender: TObject);

    procedure SaveTnToGif(filename:string);
    procedure SaveTnToTxt(filename:string);
    procedure Openscript1Click(Sender: TObject);
    procedure SaveScript1Click(Sender: TObject);
    procedure Open1Click(Sender: TObject);
    procedure btnMoveLeftClick(Sender: TObject);
    procedure btnMoveRightClick(Sender: TObject);
    procedure btnTestClick(Sender: TObject);
    procedure mmoDebugOutputClick(Sender: TObject);
    procedure btnPlayMacroClick(Sender: TObject);
    procedure btnStopMacroClick(Sender: TObject);
    procedure btnTestButtonClick(Sender: TObject);
    procedure btnResListClearClick(Sender: TObject);
    procedure btnSaveResListClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    InParams : TStringList;      // Input Parameters
    InvoiceData : TList;         // List of Invoices on Screen
    function getBIMInvoiceNumber() : Integer;
    procedure SendClear;
    procedure FailAutoProcess(reason:string);
  end;

var
  Form1: TForm1;
  Recording: Boolean;   // Global Recording
  Processing: Boolean;  // Global Processing
  Waiting: Boolean;     // Global Waiting
  WaitTime: Integer;    // Global WaitTime
  CurMacLine : Integer; // Current Macro Line
  Counter : Integer;    // The Macro Line Counter
  ImportCounter:Integer;// The current store processor line
  LineChecking : Boolean;//Checking a Line for Text
  ScreenChecking : Boolean;//Checking a Line for Text
  Checking: Boolean;    // Checking Screen Status
  CheckAttempts : Integer; // Number of times we've tried checking before failure.
  CheckData : String;   // The Data we're verifying.
  StoreFailed : Boolean;// If a store failed
  OutString : String;   // Output Buffer
  storedir : string;
  outdate : string;
  autoprocess : boolean;
  scriptfile : string;
  datafile : string;
  errormsg : string;
  failtype : integer;   //failure code: 0=unknown,1=failed,2=passed,3=critical

implementation

{$R *.dfm}
procedure TForm1.SendClear;
begin
  tnOrionScreen.SendStr(chr($0C));
end;


procedure TForm1.SaveTnToGif(filename:string);
var
  DCDesk: HDC; // hDC of telnet screen
  bmp: TBitmap;
begin
  bmp := TBitmap.Create;
  bmp.Height := tnOrionScreen.Height;
  bmp.Width := tnOrionScreen.Width-16;
  DCDesk := GetWindowDC(tnOrionScreen.Handle);
  BitBlt(bmp.Canvas.Handle, 0, 0, tnOrionScreen.Width, tnOrionScreen.Height,
         DCDesk, 0, 0, SRCCOPY);
  with TGIFImage.Create do
    try
      Assign(bmp);
      SaveToFile(filename);
    finally
      Free;
    end;
  ReleaseDC(GetDesktopWindow, DCDesk);
  bmp.Free;
  Inc(Counter);
end;

procedure TForm1.SaveTnToTxt(filename:string);
var
  i : integer;
begin
  with TStringList.Create do
    try
      for i := tnOrionScreen.Rows - 1 downto 0  do
      begin
        Add(Copy(tnOrionScreen.Screen.flines[i].txt,0,80));
      end;
      SaveToFile(filename);
    finally
      Free;
    end;
  Inc(counter);    
end;


function TForm1.getBIMInvoiceNumber() : Integer;
var
  i : integer;
  s : string;
  dp : ^TOrionInvoice;
begin
  mmoDebugOutput.Lines.add('=== Reading Screen Data ===');
  InvoiceData.Clear;
  for i := 9 to 19 do
  begin
    s := tnOrionScreen.Screen.Lines[i].txt;
    if Trim(s) <> '' then
    begin
      New(dp);
      dp.invoiceno := Trim(Copy(s,5,7));
      dp.date := Trim(Copy(s,16,8));
      dp.natnum := Trim(Copy(s,27,6));
      dp.tx := Trim(Copy(s,40,2));
      dp.total := StrToReal(Trim(Copy(s,53,9)));
      if ((Length(dp.invoiceno) > 6) and (dp.tx <> 'CN')) then
        InvoiceData.Add(dp);
    end;
  end;
  Result := 0;
  if (InvoiceData.Count > 0) then
  begin
    InvoiceData.Sort(compareInvoices);
    Result := StrToInt(TOrionInvoice(InvoiceData.Items[0]^).invoiceno);
  end;
end;

procedure TForm1.FailAutoProcess(reason:string);
begin
  Writeln(reason);
  Application.Terminate;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  if (ParamCount > 1) then begin
    autoprocess := true;
    scriptfile := ParamStr(1);
    datafile := ParamStr(2);
    mmoDebugOutput.Lines.Add('Processing Script: '+scriptfile);
    mmoDebugOutput.Lines.Add('Processing Data File: '+datafile);
    if FileExists(ExtractFileDir(ParamStr(0))+'\Scripts\'+scriptfile) then
      rlOrionMacro.Strings.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Scripts\'+scriptfile)
    else FailAutoProcess('Could not find Specified Script File');
    if FileExists(ExtractFileDir(ParamStr(0))+'\Data\'+datafile) then
      lstImportData.Items.LoadFromFile(ExtractFileDir(ParamStr(0))+'\Data\'+datafile)
    else FailAutoProcess('Could not find Specified Data File');
    btnPlayMacroClick(nil);
  end;
  storedir := ExtractFileDir(ParamStr(0))+'\Data\lostdata';
  InvoiceData := TList.Create;
  tnOrionScreen.Clear;
  tnOrionScreen.AddEvent(1,'F12 TO CANCEL',#27'[24~',[efIgnoreCase],nil);
end;

procedure TForm1.tmrCursorPosTimer(Sender: TObject);
begin
  lblX.Caption := 'X: '+IntToStr(tnOrionScreen.Screen.FCol);
  lblY.Caption := 'Y: '+IntToStr(tnOrionScreen.Screen.FRow);
end;

procedure TForm1.btnConnectClick(Sender: TObject);
begin
  if not tnOrionScreen.IsConnected then
  begin
    tnOrionScreen.Clear;
    tnOrionScreen.Connect;
  end else
  begin
    tnOrionScreen.Clear;
    tnOrionScreen.Disconnect;
  end;
end;

procedure TForm1.tnOrionScreenSessionConnected(Sender: TObject);
begin
  btnConnect.Caption := 'Disconnect';
  mmoDebugOutput.Lines.add('=== ORION CONNECTION ESTABLISHED ===');
end;

procedure TForm1.tnOrionScreenSessionClosed(Sender: TObject);
begin
  btnConnect.Caption := 'Connect';
  FocusControl(tnOrionScreen);
  mmoDebugOutput.Lines.add('=== ORION CONNECTION TERMINATED ===');
end;

procedure TForm1.btnFunc1Click(Sender: TObject);
begin
  tnOrionScreen.SendStr(#27'OR');       //sending [F3]
  FocusControl(tnOrionScreen);
end;

procedure TForm1.btnFunc2Click(Sender: TObject);
begin
  tnOrionScreen.SendStr(#27'[19~');     //sending [F8]
  FocusControl(tnOrionScreen);
end;

procedure TForm1.btnFunc3Click(Sender: TObject);
begin
  tnOrionScreen.SendStr(#27'[24~');     //sending [F12]
  FocusControl(tnOrionScreen);
end;

procedure TForm1.btnFunc4Click(Sender: TObject);
begin
  tnOrionScreen.SendStr(#9);            //sending [Field Plus(tab)]
  FocusControl(tnOrionScreen);
end;

procedure TForm1.btnFunc5Click(Sender: TObject);
begin
  tnOrionScreen.SendStr(#27#9);         //sending [Field Minus]
  FocusControl(tnOrionScreen);
end;


procedure TForm1.Exit1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TForm1.btnClearLogClick(Sender: TObject);
begin
  mmoDebugOutput.Clear;
end;

procedure TForm1.btnSaveLogClick(Sender: TObject);
begin
  if dlgSaveMMOLog.Execute then
    mmoDebugOutput.Lines.SaveToFile(dlgSaveMMOLog.Files[0]);
end;

procedure TForm1.tmrMainTimerTimer(Sender: TObject);
var
  cmd,tmperr,output,jumpto,failcmd,incust : string;
  i,copyx,copyy,copycount:integer;
  foundjump:boolean;
  storenum : string;
  logdir : string;
begin
  tmperr := '';
  DateTimeToString(outdate,'yyyymmddhhnnss',now);
  if (tnOrionScreen.IsConnected) then
  if (ImportCounter < lstImportData.Items.Count) then
  begin
    storenum := Trim(Parse(lstImportData.items[ImportCounter],',',1));
    storedir := ExtractFileDir(ParamStr(0))+'\Data\'+storenum;
    logdir := ExtractFileDir(ParamStr(0))+'\Log';
    if (storenum <> '') then begin
      if not DirectoryExists(storedir) then begin
        MkDir(storedir);
      end;
    end;
    lstImportData.ItemIndex := ImportCounter;
    If (CurMacLine < rlOrionMacro.Strings.count) then
    begin
      cmd := Parse(Parse(rlOrionMacro.Strings[CurMacLine],'=',1),',',1);
      tmperr := Parse(Parse(rlOrionMacro.Strings[CurMacLine],'=',1),',',2);
      if length(tmperr) > 0 then errormsg := tmperr;
      output := Parse(rlOrionMacro.Strings[CurMacLine],'=',2);
      if cmd = '[out]' then
      begin
        tnOrionScreen.SendStr(ParseString(output,lstImportData.items[ImportCounter]));
        inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      end else
      if cmd ='[clear]' then
      begin
        tnOrionScreen.SendStr(#12);
        Inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      end else
      if cmd ='[wait]' then
      begin
        if not Waiting then
        begin
          WaitTime := strtoint(output);
          Waiting := true;
          tmrDelay.Interval := WaitTime;
          tmrDelay.Enabled := true;
        end;
      end else
      if cmd ='[screenchk]' then
      begin
        if not Checking then
        begin
          CheckAttempts := 0;
          CheckData := output;
          Checking := true;
          tmrChecking.Enabled := true;
        end;
      end else
      if cmd ='[linecheck]' then
      begin
        if not LineChecking then
        begin
          CheckAttempts := 0;
          CheckData := output;
          LineChecking := true;
          tmrLineChecking.Enabled := true;
        end;
      end else
      if cmd ='[screenread]' then
      begin
        incust := Parse(output,',',1);
        copyx := StrToInt(Parse(output,',',2));
        copyy := StrToInt(Parse(output,',',3));
        copycount := StrToInt(Parse(output,',',4));
        if (incust = '{custa}') then
        begin
          custa := Trim(Copy(tnOrionScreen.Screen.Lines[copyy].Txt,copyx,copycount));
          mmoDebugOutput.Lines.Add('CUSTA: '+custa);
        end;
      end else
      if cmd ='[readinvoicenum]' then
      begin
        incust := Parse(output,',',1);
        if (incust = '{custa}') then
        begin
          i := getBIMInvoiceNumber();
          if (i > 0) then
          begin
            custa := IntToStr(i);
            mmoDebugOutput.Lines.Add('CUSTA: '+custa);
            rlProcessResults.InsertRow(Parse(lstImportData.items[ImportCounter],',',1),custa,true);
          end else
          begin
            rlProcessResults.InsertRow(Parse(lstImportData.items[ImportCounter],',',1),'N/A',true);
          end;
        end;
        Inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      end else
      if cmd ='[checkempty]' then
      begin
        incust := Parse(output,',',1);
        if (incust = '{custa}') then
        begin
          if (Trim(custa) = '') then StoreFailed := True;
        end;
        Inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      end else
      if cmd ='[screencheck]' then
      begin
        if not ScreenChecking then
        begin
          CheckAttempts := 0;
          CheckData := output;
          ScreenChecking := true;
          tmrScreenChecking.Enabled := true;
        end;
      end else
      if cmd ='[jumpnofail]' then
      begin
        if (not StoreFailed) then begin
          //Writeln('[{"'+storenum+'":"Failed","FailMsg":"'+errormsg+'"}]');
          foundjump := false;
          jumpto := Parse(output,',',1);
          failcmd := Parse(output,',',2);
          mmoDebugOutput.Lines.add('No Failure, Jumping to: '+jumpto);
          for i := 0 to rlOrionMacro.Strings.Count - 1 do
          begin
            if (Parse(rlOrionMacro.Strings[i],'=',1) = '['+jumpto+']') then
            begin
              mmoDebugOutput.Lines.Add(jumpto);
              foundjump := true;
              CurMacLine := i;
              rlProcessResults.InsertRow(Parse(lstImportData.items[ImportCounter],',',1),'No Fail',true);
              StoreFailed := false;
              tnOrionScreen.SendStr(ParseString(failcmd,''));
              break;
            end;
          end;
          if not foundjump then begin
            mmoDebugOutput.Lines.Add('ERROR IN SCRIPT: JUMP TARGET "'+jumpto+'" NOT FOUND!');
            tmrMainTimer.Enabled := false;
          end;
        end else begin
          StoreFailed := false;
          Inc(CurMacLine);
          SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
        end;
      end else
      if cmd ='[jumpfail]' then
      begin
        if (StoreFailed) then begin
          Writeln('[{"'+storenum+'":"Failed","FailMsg":"'+errormsg+'"}]');
          foundjump := false;
          jumpto := Parse(output,',',1);
          failcmd := Parse(output,',',2);
          mmoDebugOutput.Lines.add('Failed, Jumping to: '+jumpto);
          for i := 0 to rlOrionMacro.Strings.Count - 1 do
          begin
            if (Parse(rlOrionMacro.Strings[i],'=',1) = '['+jumpto+']') then
            begin
              mmoDebugOutput.Lines.Add(jumpto);
              foundjump := true;
              CurMacLine := i;
              rlProcessResults.InsertRow(Parse(lstImportData.items[ImportCounter],',',1),'Failed',true);
              tnOrionScreen.SendStr(ParseString(failcmd,''));
              break;
            end;
          end;
          if not foundjump then begin
            mmoDebugOutput.Lines.Add('ERROR IN SCRIPT: JUMP TARGET "'+jumpto+'" NOT FOUND!');
          end;
        end else begin
          Inc(CurMacLine);
          SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
        end;
      end else
      begin
        //a comment or jumpspot, ignore and continue
        inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      end;
    end else
    begin
      if not storefailed then
      Writeln('[{"'+storenum+'":"Success"}]');
      StoreFailed := false;
      for i := 0 to rlOrionMacro.Strings.Count - 1 do
      begin
        CurMacLine := 9;
        if (Parse(rlOrionMacro.Strings[i],'=',1) = '[begin]') then
        begin
          CurMacLine := i;
          StoreFailed := false;
          break;
        end;
      end;
      mmoDebugOutput.Lines.Add('Completed Processing Store: '+parse(lstImportData.items[ImportCounter],',',1));
      //OutString := concat(OutString,inttostr(counter));
      Writeln(outstring);
      Counter := 0;
      StoreFailed := false;
      Inc(ImportCounter);
      if (ImportCounter < lstImportData.Count) then
      begin
        mmoDebugOutput.Lines.Add('Processing Store: '+parse(lstImportData.items[ImportCounter],',',1));
      end;
//      tmrMainTimer.Enabled := false;
    end;
  end else
  begin
    logdir := ExtractFileDir(ParamStr(0))+'\Log';
    DateTimeToString(outdate,'yyyymmddhhnnss',now);
    mmoDebugOutput.Lines.SaveToFile(logdir+'\OS'+outdate+'.log');
    tmrMainTimer.Enabled := false;
    if autoprocess then Application.Terminate;
  end;
end;

procedure TForm1.tmrScreenCheckingTimer(Sender: TObject);
var
  foundstring : boolean;
  i : integer;
  sString : string;
begin
  if (ScreenChecking) then
  begin
    foundstring := false;
//    sString := Parse(CheckData,',',2);
    sString := Parse(ParseString(CheckData,lstImportData.items[ImportCounter]),',',1);
    mmoDebugOutput.Lines.Add(Format('Checking screen for text: "%s"',[sString]));
    for i := 0 to tnOrionScreen.Rows - 1 do
    begin
      if (Pos(sString,tnOrionScreen.Screen.Lines[i].Txt)>0) then
      begin
        Inc(CurMacLine);
        SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
        ScreenChecking := false;
        foundstring := true;
        mmoDebugOutput.lines.add('Found String!');
        tmrScreenChecking.enabled := False;
      end;
    end;
    if not foundstring then
    begin
      mmoDebugOutput.lines.add('Failure to find String on Screen! Line: '+IntToStr(CurMacLine));
      ScreenChecking := false;
      tmrScreenChecking.enabled := False;
      StoreFailed := true;
      Inc(CurMacLine);
      SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
    end;
  end;
end;

procedure TForm1.tmrLineCheckingTimer(Sender: TObject);
var
  iLine : integer;
  sString : string;
begin
  if (LineChecking) then
  begin
    iLine := StrToInt(Parse(CheckData,',',1));
//    sString := Parse(CheckData,',',2);
    sString := Parse(ParseString(CheckData,lstImportData.items[ImportCounter]),',',2);
    mmoDebugOutput.Lines.Add(Format('Checking line [%d] for text: "%s"',[iLine,sString]));
    if (Pos(sString,tnOrionScreen.Screen.Lines[iLine].Txt)>0) then
    begin
      Inc(CurMacLine);
      SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      LineChecking := false;
      mmoDebugOutput.lines.add('Found String!');
      tmrLineChecking.enabled := False;
    end else
    begin
      mmoDebugOutput.lines.add('Failure to find String! Line: '+IntToStr(CurMacLine));
      LineChecking := false;
      tmrLineChecking.enabled := False;
      StoreFailed := true;
      Inc(CurMacLine);
      SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
    end;
  end;
end;

procedure TForm1.tmrCheckingTimer(Sender: TObject);
var
  ScreenX,ScreenY,CheckX,CheckY:Integer;
  Failed : boolean;
begin
  Failed := false;
  ScreenX := tnOrionScreen.Screen.FCol;
  ScreenY := tnOrionScreen.Screen.FRow;
  if (Checking) and (CheckAttempts < 3) then
  begin
    CheckX := StrToInt(Parse(CheckData,',',1));
    CheckY := StrToInt(Parse(CheckData,',',2));
    if (ScreenX <> CheckX) then Failed := true;
    if (ScreenY <> CheckY) then Failed := true;
    Inc(CheckAttempts);
    if (Failed) and (CheckAttempts = 3) then
    begin
      mmoDebugOutput.Lines.add('!!!!!!!! FAILED CHECK ON LINE '+IntToStr(CurMacLine)+' !!!!!!!!');
//      tmrDelay.Enabled := false;
//      tmrMainTimer.Enabled := false;
      tmrChecking.Enabled := false;
//      tmrCursorPos.Enabled := false;
      Checking := false;
      Waiting := false;
      Inc(CurMacLine);
      SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
      StoreFailed := true;
    end;
    if not Failed then
    begin
      Checking := false;
      tmrChecking.enabled := false;
      Inc(CurMacLine);
      SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
    end;
  end;
end;

procedure TForm1.tmrDelayTimer(Sender: TObject);
begin
  if (Waiting) then
  begin
    Waiting := false;
    tmrDelay.Enabled := false;
    inc(CurMacLine);
    SaveTnToTxt(storedir+'\Debug'+'_'+outdate+'_'+inttostr(Counter)+'.txt');
  end;
end;

procedure TForm1.Openscript1Click(Sender: TObject);
begin
  if (dlgOpenMacro.Execute) then
  begin
    rlOrionMacro.Strings.LoadFromFile(dlgOpenMacro.Files[0]);
    pgcMain.TabIndex := 1;
    mmoScriptEdit.Lines := rlOrionMacro.Strings;
  end;
end;

procedure TForm1.SaveScript1Click(Sender: TObject);
begin
  if (dlgSaveMacro.Execute) then
  begin
    rlOrionMacro.Strings.SaveToFile(dlgSaveMacro.Files[0]);
  end;
end;

procedure TForm1.Open1Click(Sender: TObject);
begin
  if (dlgOpenData.Execute) then
    lstImportData.Items.LoadFromFile(dlgOpenData.FileName);
end;

procedure TForm1.btnMoveLeftClick(Sender: TObject);
begin
  mmoScriptEdit.Lines := rlOrionMacro.Strings;
end;

procedure TForm1.btnMoveRightClick(Sender: TObject);
begin
  rlOrionMacro.Strings := mmoScriptEdit.Lines;
end;

procedure TForm1.btnTestClick(Sender: TObject);
var
  i: Integer;
begin
  SaveTnToTxt(ExtractFileDir(ParamStr(0))+'\test.txt');
  for i := 0 to tnOrionScreen.Rows - 1 do
    mmoDebugOutput.Lines.Add(Format('%.2d',[i])+': '+Copy(tnOrionScreen.Screen.Lines[i].Txt,0,80));
  pgcMain.TabIndex := 2;
end;

procedure TForm1.mmoDebugOutputClick(Sender: TObject);
var
  Row,Col : Integer;
begin
  Row := SendMessage(mmoDebugOutput.Handle, EM_LINEFROMCHAR, mmoDebugOutput.SelStart, 0);
  Col := mmoDebugOutput.SelStart - SendMessage(mmoDebugOutput.Handle, EM_LINEINDEX, Row, 0);
  lblDebugX.Caption := 'X: '+IntToStr(Col-4);
end;

procedure TForm1.btnPlayMacroClick(Sender: TObject);
begin
  Recording := false;
  mmoDebugOutput.Lines.add('=== BEGINNING SCRIPT PROCESSING ===');
  ImportCounter := 0;
  CurMacLine := 0;
  Counter := 0;
  tnOrionScreen.Disconnect;
  tnOrionScreen.Connect;
  tmrMainTimer.Enabled := true;
  if (ImportCounter < lstImportData.Count) then
    mmoDebugOutput.Lines.Add('Processing Store: '+parse(lstImportData.items[ImportCounter],',',1));
end;

procedure TForm1.btnStopMacroClick(Sender: TObject);
begin
  Recording := false;
  mmoDebugOutput.Lines.add('=== ENDING SCRIPT PROCESSING ===');
  ImportCounter := 0;
  CurMacLine := 0;
  Counter := 0;
  tmrMainTimer.Enabled := false;
end;

procedure TForm1.btnTestButtonClick(Sender: TObject);
begin
  SendClear;
end;

procedure TForm1.btnResListClearClick(Sender: TObject);
begin
  rlProcessResults.Strings.Clear;
end;

procedure TForm1.btnSaveResListClick(Sender: TObject);
begin
  if (dlgSaveLog.Execute) then
    rlProcessResults.Strings.SaveToFile(dlgSaveLog.Files[0]);
end;

end.
