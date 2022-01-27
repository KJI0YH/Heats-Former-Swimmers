unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ActiveX, ComObj, ActnList, Menus, Grids, ExtCtrls, ComCtrls, StdCtrls,
  AxCtrls, OleCtrls, VCF1, OleServer, ExcelXP, OleCtnrs, ToolWin,
  uProgram;

type
  TfrmMain = class(TForm)
    pnlAll: TPanel;
    pnlButtons: TPanel;
    pnlProgram: TPanel;
    pnlStarted: TPanel;
    pnlApplicat: TPanel;
    pnlBTNApplicat: TPanel;
    pnlBTNStarted: TPanel;
    pnlHeader: TPanel;
    pnlProgress: TPanel;
    pnlErrors: TPanel;
    pnlFinishBTN: TPanel;    

    ActionList: TActionList;
    actCheckAdd: TAction;
    actEditItem: TAction;
    actSaveProgram: TAction;
    actOpenProgram: TAction;
    actOpenTechnical: TAction;
    actSaveApplicat: TAction;
    actOpenApplicat: TAction;
    actSaveSarted: TAction;
    actCreateStarted: TAction;
    actCreateApplicat: TAction;
    actDeleteItem: TAction;
    actSaveFinish: TAction;
    acrCreateFinish: TAction;
    actOpenStarted: TAction;

    SaveProgram: TSaveDialog;
    OpenProgram: TOpenDialog;
    SaveXL: TSaveDialog;
    OpenApplicat: TOpenDialog;
    OpenTechnical: TOpenDialog;

    tcTabs: TPageControl;
    tsProgramFields: TTabSheet;
    tsApplicat: TTabSheet;
    tcErrors: TTabSheet;
    tsFinish: TTabSheet;

    splVertical: TSplitter;
    splHorizont: TSplitter;

    leOrder: TLabeledEdit;
    lblMeters: TLabel;
    lblStyle: TLabel;
    lblSex: TLabel;
    lblFromYear: TLabel;
    lblToYear: TLabel;
    lblProgress: TLabel;

    btnAdd: TButton;
    btnEdit: TButton;
    btnOpen: TButton;
    btnSave: TButton;
    btnDelete: TButton;
    btnOpenTechnical: TButton;
    btnOpenApplicat: TButton;
    btnSaveStarted: TButton;
    btnSaveApplicat: TButton;
    btnCreateStarted: TButton;
    btnCreateApplicat: TButton;
    btnCreateFinish: TButton;
    btnSaveFinish: TButton;
    btnOpenStarted: TButton;

    cbMeters: TComboBox;
    cbSex: TComboBox;
    cbStyle: TComboBox;

    edtMinYear: TEdit;
    edtMaxYear: TEdit;

    lvProgram: TListView;
    lvErrors: TListView;

    gbApplicat: TGroupBox;
    gbStarted: TGroupBox;
    gbFinish: TGroupBox;

    OleApplicat: TOleContainer;
    OleStarted: TOleContainer;
    OleFinish: TOleContainer;

    pbProgress: TProgressBar;

    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    N4: TMenuItem;
    N5: TMenuItem;
    N6: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    N11: TMenuItem;
    Excel1: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    Excel2: TMenuItem;

    procedure FormCreate(Sender: TObject);
    procedure actCheckAddExecute(Sender: TObject);
    procedure actEditItemExecute(Sender: TObject);
    procedure actDeleteItemExecute(Sender: TObject);
    procedure actSaveProgramExecute(Sender: TObject);
    procedure actOpenProgramExecute(Sender: TObject);
    procedure actOpenTechnicalExecute(Sender: TObject);
    procedure actSaveApplicatExecute(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure actOpenApplicatExecute(Sender: TObject);
    procedure actSaveSartedExecute(Sender: TObject);
    procedure actCreateStartedExecute(Sender: TObject);
    procedure actCreateApplicatExecute(Sender: TObject);
    procedure actOpenStartedExecute(Sender: TObject);
    procedure acrCreateFinishExecute(Sender: TObject);
    procedure actSaveFinishExecute(Sender: TObject);
    procedure lvProgramCompare(Sender: TObject; Item1, Item2: ComCtrls.TListItem; Data: Integer; var Compare: Integer);
    procedure lvProgramSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
    procedure tcTabsChange(Sender: TObject);
    procedure FormResize(Sender: TObject);

    Procedure ClearFields;
    Procedure AddDistance(Var ListView: TListView; Order, Sex, fromYear, toYear, Meters, Style: String);
    Procedure LoadList(Var ListView: TListView; ProgramHead: uProgram.PProgram);

    Function IsExcelExist: Boolean;
    Function IsExcelRun: Boolean;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

Const

  //messages
  strProgramUnSave = 'Программа соревнований не сохранена';
  strProgramLoad = 'Программа соревнований загружена';
  strProgramUnLoad = 'Программа соревнований не загружена';
  strProgramSave = 'Программа соревнований сохранена';

  strTechnicalLoad = 'Загрузка технических протоколов';
  strTechnicalProgress = 'Загружено технических (';

  strApplicatCreate = 'Создание заявочного протокола';
  strApplicatCreated = 'Заявочный протокол создан';
  strApplicatSave = 'Заявочный протокол сохранён';
  strApplicatSaving = 'Сохранение заявочного протокола';
  strApplicatUnSave = 'Заявочный протокол не сохранён';
  strApplicatLoad = 'Загрузка заявочного протокола';
  strApplicatLoaded = 'Заявочный протокол загружен';

  strStartedCreate = 'Создание стартового протокола';
  strStartedCreated = 'Стартовый протокол создан';
  strStartedSave = 'Стартовый протокол сохранён';
  strStartesUnSave = 'Стартовый протокол не сохранён';
  strStartedLoad = 'Загрузка стартового протокола';

  strFinishCreate = 'Создание итогового протокола';
  strFinishCreated = 'Итоговый протокол создан';
  strFinishSaved = 'Итоговый протокол сохранён';
  strFinishSave = 'Сохранение итогового протокола';
  strFinishLoad = 'Загрузка финишного протокола';
  strFinishLoaded = 'Финишный протокол загружен';

var
  frmMain: TfrmMain;
  Order: Integer = 1;
  FileCount: Integer = 0;
  Opened: Integer = 0;

implementation

uses uSwimmer, uExcel, uTechnical, uApplicat, uError, uFinish;

{$R *.dfm}

{*******************
*                  *
* General settings *
*                  *
*******************}

procedure TfrmMain.FormCreate(Sender: TObject);
Var
  Meters: uProgram.TMeters;
  Styles: uProgram.TStyles;
  Sex: uProgram.TSex;
begin
  {creating items for combo boxes}

  //meters combobox
  For Meters:=Low(uProgram.SMeters) to High(uProgram.SMeters) do
    frmMain.cbMeters.Items.Add(uProgram.SMeters[Meters]);

  //styles combobox
  For Styles:=Low(uProgram.SStyles) to High(uProgram.SStyles) do
    frmMain.cbStyle.Items.Add(uProgram.SStyles[Styles]);

  //sex combobox
  For Sex:=Low(uProgram.SSex) to High(uProgram.SSex) do
    frmMain.cbSex.Items.Add(uProgram.SSex[Sex]);

  //setup initial order state
  frmMain.leOrder.Text:=IntToStr(Order);
  lblProgress.Font.Color:=clRed;
  lblProgress.Caption:=uMain.strProgramUnLoad;
end;

//custom sort for list view
Function CustomSort(Item1, Item2: TListItem; Data: Integer): Integer; stdcall;
Var
  N1, N2: Cardinal;
Begin
  N1:=StrToInt(Item1.Caption);
  N2:=StrToInt(Item2.Caption);
  If N1>N2 then
    Result:=-1
  Else IF N1<N2 then
    Result:=1
  Else
    Result:=0;
End;

procedure TfrmMain.FormResize(Sender: TObject);
begin
  pnlProgress.Width:=pnlHeader.Width-lblProgress.Width-10;
end;

//checking excel server on pc
Function TfrmMain.IsExcelExist: Boolean;
Var
  ClassID: TCLSID;
  Rez: HRESULT;
Begin
  Result:=True;
  Rez:=CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  If Rez<>S_OK then
  Begin
    Result:=False;
    MessageDlg('Сервер Excel не установлен на компьютере', mtERROR, [mbOk], 0);
  End;
End;

//check active object
Function TfrmMain.IsExcelRun: Boolean;
Var
  ClassID: TCLSID;
  Unknown: IUnknown;
Begin
  Result:=False;
  ClassID:=ProgIDToClassID('Excel.Application');
  If GetActiveObject(ClassID, nil, Unknown)=S_OK then
  Begin
    Result:=True;
    MessageDlg('Excel запущен',mtWarning,[mbOK],0);
  End;
End;

{**********************************
*                                 *
* Program of the competition part *
*                                 *
**********************************}

//comparing for custom list view sort
procedure TfrmMain.lvProgramCompare(Sender: TObject; Item1, Item2: ComCtrls.TListItem; Data: Integer; var Compare: Integer);
Var
  N1, N2: Integer;
begin
  N1:=StrToInt(Item1.Caption);
  N2:=StrToInt(Item2.Caption);
  If N1>N2 then
    Compare:=1
  Else If N1<N2 then
    Compare:=-1
  Else
    Compare:=0;
end;

//loading list to a list view
Procedure TfrmMain.LoadList(Var ListView: TListView; ProgramHead: uProgram.PProgram);
Var
  Current: PProgram;
  SMinYear, SMaxYear: String;
Begin
  Current:=ProgramHead^.Next;
  uMain.Order:=1;

  ListView.Items.BeginUpdate;
  ListView.Items.Clear;
  While Current<>Nil do
  Begin
    With Current^ do
    Begin

      Case maxYear of
        0: SMaxYear:='';
        -1: SMaxYear:=Copy(SYear, Pos('-',SYear)+1, Length(SYear));
        Else SMaxYear:=IntToStr(maxYear);
      End;

      Case minYear of
        0: SMinYear:='';
        -1: SMinYear:=Copy(SYear, 1, Pos('-',SYear)-1);
        Else SMinYear:=IntToStr(minYear);
      End;

      frmMain.AddDistance(ListView, IntToStr(uMain.Order), SSex[Sex], SMinYear, SMaxYear, SMeters[Meters], SStyles[Style]);
      Current:=Current^.Next;
      Inc(uMain.Order);
    End;
  End;
  ListView.Items.EndUpdate;
End;

//edit and delete buttons checking
procedure TfrmMain.lvProgramSelectItem(Sender: TObject; Item: TListItem; Selected: Boolean);
begin
  If Selected then
  Begin
      frmMain.btnEdit.Enabled:=True;
      frmMain.btnDelete.Enabled:=True;
  End
  Else
  Begin
      frmMain.btnEdit.Enabled:=False;
      frmMain.btnDelete.Enabled:=False;
  End;
end;

//lock/unlock add button
procedure TfrmMain.actCheckAddExecute(Sender: TObject);
begin
  If (frmMain.leOrder.Text<>'') and (frmMain.cbSex.Text<>'') and (frmMain.cbMeters.Text<>'') and
     (frmMain.cbStyle.Text<>'') and ((frmMain.edtMinYear.Text<>'') or (frmMain.edtMaxYear.Text<>'')) then
    frmMain.btnAdd.Enabled:=True
  Else
    frmMain.btnAdd.Enabled:=False;
end;

//adding distance to a program list
Procedure TfrmMain.AddDistance(Var ListView: TListView; Order, Sex, fromYear, toYear, Meters, Style: String);
Var
  I: Integer;
  OrdNum: Integer;
Begin

  //calculate next order
  OrdNum:=StrToInt(Order);
  If OrdNum>ListView.Items.Count then
    OrdNum:=ListView.Items.Count+1;

  //reorder distances
  For I:=OrdNum to lvProgram.Items.Count do
    lvProgram.Items.Item[I-1].Caption:=IntToStr(I);

  //delete item if selected (edit item)
  If frmMain.btnEdit.Enabled then
  Begin
    ListView.Selected.Delete;
    frmMain.btnEdit.Enabled:=False;
    frmMain.btnDelete.Enabled:=False;
  End;

  ListView.Items.BeginUpdate;

  //adding item to a list view
  With ListView.Items.Add do
  Begin
    Caption:=IntToStr(OrdNum);
    SubItems.Add(Sex);

    If fromYear='' then
      SubItems.Add(toYear)
    Else If toYear='' then
      SubItems.Add(fromYear)
    Else
      SubItems.Add(fromYear+'-'+toYear);

    SubItems.Add(Meters);
    SubItems.Add(Style);
  End;

  //sorting items by order
  ListView.SortType:=stData;
  ListView.SortType:=stNone;
  ListView.Items.EndUpdate;
  frmMain.ClearFields;
  frmMain.actCheckAddExecute(frmMain.btnAdd);
  lblProgress.Font.Color:=clBlue;
  lblProgress.Caption:=uMain.strProgramUnSave;
End;

//add distance
procedure TfrmMain.btnAddClick(Sender: TObject);
begin
  frmMain.AddDistance(frmMain.lvProgram,leOrder.Text,cbSex.Text,edtMinYear.Text, edtMaxYear.Text,cbMeters.Text,cbStyle.Text);
end;

//edit distance
procedure TfrmMain.actEditItemExecute(Sender: TObject);
begin

  //order field
  leOrder.Text:=lvProgram.Selected.Caption;

  //sex field
  cbSex.ItemIndex:=uProgram.GetSexInd(lvProgram.Selected.SubItems[0]);

  //meters field
  cbMeters.ItemIndex:=uProgram.GetMetersInd(lvProgram.Selected.SubItems[2]);

  //styles field
  cbStyle.ItemIndex:=GetStyleInd(lvProgram.Selected.SubItems[3]);

  //years field
  If Pos('-', lvProgram.Selected.SubItems[1])=0 then
  Begin
    edtMinYear.Text:=lvProgram.Selected.SubItems[1];
    edtMaxYear.Text:='';
  End
  Else
  Begin
    edtMinYear.Text:=Copy(lvProgram.Selected.SubItems[1],1,Pos('-',lvProgram.Selected.SubItems[1])-1);
    edtMaxYear.Text:=Copy(lvProgram.Selected.SubItems[1],Pos('-',lvProgram.Selected.SubItems[1])+1, Length(lvProgram.Selected.SubItems[1]));
  End;

  frmMain.actCheckAddExecute(Sender);
end;

//deleting item from list view
procedure TfrmMain.actDeleteItemExecute(Sender: TObject);
Var
  OrdNum, I: Integer;
begin
  If lvProgram.Selected=Nil then Exit;
  OrdNum:=StrToInt(lvProgram.Selected.Caption);
  btnDelete.Enabled:=False;
  btnEdit.Enabled:=False;

  //reorder distances in list view
  For I:=OrdNum to lvProgram.Items.Count do
    lvProgram.Items.Item[I-1].Caption:=IntToStr(I-1);

  lvProgram.Selected.Delete;

  uMain.Order:=lvProgram.Items.Count+1;
  leOrder.Text:=IntToStr(uMain.Order);

  If lvProgram.Items.Count<>0 then
  Begin
    lblProgress.Font.Color:=clBlue;
    lblProgress.Caption:=uMain.strProgramUnSave;
  End
  Else
  Begin
    lblProgress.Font.Color:=clRed;
    lblProgress.Caption:=uMain.strProgramUnload;
  End;
end;

//creating fields after adding distance
Procedure TfrmMain.ClearFields;
Begin
  With frmMain do
  Begin
    Order:=lvProgram.Items.Count;
    leOrder.Text:=IntToStr(Order+1);
    If cbSex.ItemIndex>=(cbSex.Items.Count-1) then
      cbSex.ItemIndex:=0
    Else
      cbSex.ItemIndex:=cbSex.ItemIndex+1;
  End;
End;

//open existing program and loading it to a list view
procedure TfrmMain.actOpenProgramExecute(Sender: TObject);
begin
  If not OpenProgram.Execute then Exit;

  //open program and creating program linked list
  uProgram.OpenProgram(OpenProgram.FileName, uProgram.ProgramHead, uProgram.DistList);

  //loading program into a list view
  frmMain.LoadList(frmMain.lvProgram, uProgram.ProgramHead);

  lblProgress.Font.Color:=clGreen;
  lblProgress.Caption:=uMain.strProgramLoad;
end;

//saving program of the competition
procedure TfrmMain.actSaveProgramExecute(Sender: TObject);
begin
  If not SaveProgram.Execute then Exit;

  //creating program list
  uProgram.CreateProgramList(uProgram.ProgramHead, lvProgram, uProgram.DistList);

  //save created program list
  uProgram.SaveProgram(SaveProgram.FileName,uProgram.ProgramHead);
  lblProgress.Font.Color:=clGreen;
  lblProgress.Caption:=uMain.strProgramSave;
end;

{*************************
*                        *
* Applicat protocol part *
*                        *
*************************}

//opening multiple technical files
procedure TfrmMain.actOpenTechnicalExecute(Sender: TObject);
Var
  I: Integer;
begin
  OpenTechnical.Title:=uMain.strTechnicalLoad;
  If (uProgram.ProgramHead=Nil) then
    uProgram.CreateProgramList(uProgram.ProgramHead,frmMain.lvProgram,uProgram.DistList);
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (uProgram.ProgramHead^.Next=Nil) or (not OpenTechnical.Execute) then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  pbProgress.Step:=pbProgress.Max div OpenTechnical.Files.Count;

  uMain.FileCount:=uMain.FileCount+OpenTechnical.Files.Count;
  lblProgress.Caption:=uMain.strTechnicalProgress+IntToStr(uMain.Opened)+' из '+IntToStr(uMain.FileCount)+')';

  //process all techical files
  For I:=0 to OpenTechnical.Files.Count-1 do
  Begin
    uTechnical.ReadTechnical(OpenTechnical.Files[I],uProgram.DistList);
    pbProgress.Position:=pbProgress.Step*(I+1);
    Inc(uMain.Opened);
    lblProgress.Caption:=uMain.strTechnicalProgress+IntToStr(uMain.Opened)+' из '+IntToStr(uMain.FileCount)+')';
  End;
  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;
end;

//creating applicate protocol
procedure TfrmMain.actCreateApplicatExecute(Sender: TObject);
Var
  XL: Variant;
begin
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  lblProgress.Caption:=uMain.strApplicatCreate;

  OleApplicat.CreateObject('Excel.Sheet', True);
  XL:=OleApplicat.OleObject;
  uExcel.CreateApplicat(uProgram.ProgramHead, XL);

  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;

  XL:=Unassigned;
  lblProgress.Caption:=uMain.strApplicatCreated;
end;

//saving applicat protocol as Excel table
procedure TfrmMain.actSaveApplicatExecute(Sender: TObject);
begin
  SaveXL.Title:=uMain.strApplicatSaving;
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (OleApplicat.State=osLoaded) or (not SaveXL.Execute) then Exit;

  OleApplicat.OleObject.SaveAs(SaveXL.FileName);
  lblProgress.Caption:=uMain.strApplicatSave;
end;

{************************
*                       *
* Started protocol part *
*                       *
************************}

//open and read applicat protocol
procedure TfrmMain.actOpenApplicatExecute(Sender: TObject);
begin
  OpenApplicat.Title:=strApplicatLoad;
  If (uProgram.ProgramHead=Nil) then
    uProgram.CreateProgramList(uProgram.ProgramHead,frmMain.lvProgram,uProgram.DistList);
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (uProgram.ProgramHead^.Next=Nil) or not OpenApplicat.Execute then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  lblProgress.Caption:=uMain.strApplicatLoad;

  uProgram.CreateProgramList(uProgram.ProgramHead, frmMain.lvProgram, uProgram.DistList);
  uApplicat.ReadApplicat(OpenApplicat.FileName, uProgram.ProgramHead);

  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;

  lblProgress.Caption:=uMain.strApplicatLoaded;
end;

//save started protocol as Excel table
procedure TfrmMain.actSaveSartedExecute(Sender: TObject);
begin
  SaveXL.Title:=uMain.strStartedSave;
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (OleStarted.State=osLoaded) or (not SaveXL.Execute) then Exit;
  OleStarted.OleObject.SaveAs(SaveXL.FileName);
end;

//creating started protocol
procedure TfrmMain.actCreateStartedExecute(Sender: TObject);
Var
  XL: Variant;
begin
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  lblProgress.Caption:=uMain.strStartedCreate;

  OleStarted.CreateObject('Excel.Sheet', True);
  XL:=OleStarted.OleObject;
  uExcel.CreateStarted(uProgram.ProgramHead, XL);

  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;

  XL:=Unassigned;

  lblProgress.Caption:=uMain.strStartedCreated;
end;

{***************
*              *
* Error report *
*              *
***************}

//show errors on list view
procedure TfrmMain.tcTabsChange(Sender: TObject);
Var
  Current: uError.PErrorList;
begin
  If (uError.ErrorList<>Nil) and (tcTabs.TabIndex=2) then
  Begin
    Current:=uError.ErrorList^.Next;
    lvErrors.Items.BeginUpdate;
    lvErrors.Items.Clear;
    While Current<>Nil do
    Begin
      With lvErrors.Items.Add do
      Begin
        Caption:=uError.SErrorLevel[Current^.ErrorLevel];
        SubItems.Add(Current^.ErrorInfo);
        SubItems.Add(Current^.FileName);
      End;
      Current:=Current^.Next;
    End;
    lvErrors.Items.EndUpdate;
  End;
end;

{******************
*                 *
* Finish protocol *
*                 *
******************}

//open started protocol with finish results
procedure TfrmMain.actOpenStartedExecute(Sender: TObject);
begin
  OpenApplicat.Title:=strFinishLoad;
  If (uProgram.ProgramHead=Nil) then
    uProgram.CreateProgramList(uProgram.ProgramHead,frmMain.lvProgram,uProgram.DistList);
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (uProgram.ProgramHead^.Next=Nil) or not OpenApplicat.Execute then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  lblProgress.Caption:=uMain.strFinishLoad;

  uFinish.ReadStarted(OpenApplicat.FileName,uProgram.ProgramHead);

  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;

  lblProgress.Caption:=uMain.strFinishLoaded;
end;

//creating finish protocol
procedure TfrmMain.acrCreateFinishExecute(Sender: TObject);
Var
  XL: Variant;
begin
  If (uProgram.ProgramHead=Nil) then
    uProgram.CreateProgramList(uProgram.ProgramHead,frmMain.lvProgram,uProgram.DistList);
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (uProgram.ProgramHead^.Next=Nil) then Exit;

  pbProgress.Min:=0;
  pbProgress.Max:=1000;

  lblProgress.Caption:=uMain.strFinishCreate;

  OleFinish.CreateObject('Excel.Sheet', True);
  XL:=OleFinish.OleObject;

  uExcel.CreateFinish(uProgram.ProgramHead,XL);

  pbProgress.Position:=pbProgress.Max;
  Sleep(300);
  pbProgress.Position:=pbProgress.Min;

  XL:=Unassigned;
  lblProgress.Caption:=uMain.strFinishCreated;
end;

//saving finish protocol as Excel table
procedure TfrmMain.actSaveFinishExecute(Sender: TObject);
begin
  SaveXL.Title:=uMain.strFinishSave;
  If (not IsExcelExist) or (uProgram.ProgramHead=Nil) or (OleFinish.State=osLoaded) or(uFinish.SummaryList=Nil) or (not SaveXL.Execute) then Exit;
  OleFinish.OleObject.SaveAs(SaveXL.FileName);
end;

end.













