unit uProgram;
{the unit responsible for
the competition program}

interface

Uses
  ComCtrls, SysUtils,
  uSwimmer;

Const

  //search columns
  cOrder = '№';
  cName = 'Ф.И.';
  cYear = 'ГОД РОЖД';
  cCity = 'ГОРОД';

Type

  //subtypes structures
  TSYear = String[128];
  TSex = (Male, Female);
  TMeters = (m25, m50, m100, m200, m400, m800, m1500);
  TStyles = (Fly, Back, Breast, Free, Medley);

  //program list structure
  PProgram = ^TProgram;
  TProgram = Record
    Meters: TMeters;
    Style: TStyles;
    Sex: TSex;
    minYear, maxYear: Integer;
    SYear: TSYear;
    Swimmers: PSwimmer;
    Next: PProgram;
  End;

  //file of program list
  FProgram = File Of TProgram;

  //list of unique distances
  TDistItem = Record
    Meters: TMeters;
    Style: TStyles;
    Col: Integer;
  End;
  TDistList = Array of TDistItem;
  TDistExist = Array [TMeters] of Array [TStyles] of Boolean;

  //string names of subtypes structure
  TSMeters = Array [TMeters] of String;
  TSStyles = Array [TStyles] of String;
  TSSex = Array [TSex] of String;

Const

  //string values of the types
  SMeters: TSMeters = ('25м','50м','100м','200м','400м','800м','1500м');
  SStyles: TSStyles = ('Баттерфляй','На спине','Брасс','Вольный стиль','Комплекс');
  SSex: TSSex = ('Мужчины', 'Женщины');

Var
  ProgramHead: PProgram = Nil;
  DistList: TDistList;
  {ProgramHead - head of the program list\
  DistList - list of the unique distations}

Procedure CreateProgramList(Var ProgramHead: PProgram; ListView: TListView; Var DistList: TDistList);
Procedure CreateDistList(Var DistList: TDistList; Head: PProgram);
Procedure SaveProgram(FileName: String; ProgramHead: PProgram);
Procedure OpenProgram (FileName: String; Var ProgramHead: PProgram; Var DistList: TDistList);
Function GetSexInd(Str: String): Integer;
Function GetMetersInd(Str: String): Integer;
Function GetStyleInd(Str: String): Integer;
Function GetDistInd(Meters: TMeters; Style: TStyles): Integer;
Procedure ClearSwimmerList(Var SwimmerHead: PSwimmer);
                    
implementation

uses uError;

{*******************
* local procedures *
********************}

//clearing swimmer list
Procedure ClearSwimmerList(Var SwimmerHead: PSwimmer);
Var
  Deleted: PSwimmer;
Begin
  If SwimmerHead=Nil then Exit;
  While SwimmerHead^.Next<>Nil do
  Begin
    Deleted:=SwimmerHead^.Next;
    SwimmerHead^.Next:=SwimmerHead^.Next^.Next;
    Dispose(Deleted);
  End;
  Dispose(SwimmerHead);
  SwimmerHead:=Nil;
End;

//clearing program list
Procedure ClearProgramList(Var ProgramHead: PProgram);
Var
  Deleted: PProgram;
Begin
  If ProgramHead=Nil then Exit;
  While ProgramHead^.Next<>Nil do
  Begin
    Deleted:=ProgramHead^.Next;
    ProgramHead^.Next:=ProgramHead^.Next^.Next;
    ClearSwimmerList(ProgramHead^.Swimmers);
    Dispose(Deleted);
  End;
  Dispose(ProgramHead);
  ProgramHead:=Nil;
End;

//reading years from string
Procedure ReadYears(Var minYear, maxYear: Integer; SYear: String);
Var
  Buffer: String;
  Err: Integer;
Begin

  //determining if only one year in range
  If Pos('-',SYear)=0 then
  Begin
    minYear:=StrToInt(SYear);
    maxYear:=0;
  End
  Else
  Begin

    //determining the lower part of the range
    Buffer:=Copy(SYear,1,Pos('-',SYear)-1);
    Val(Buffer, minYear, Err);
    If Err<>0 then
      minYear:=-1;

    //determining the highest part of the range
    Buffer:=Copy(SYear, Pos('-',SYear)+1, Length(SYear));
    Val(Buffer, maxYear, Err);
    If Err<>0 then
      maxYear:=-1;
  End;
End;

{***********************
* interface procedures *
***********************}

//opening existing program
Procedure OpenProgram (FileName: String; Var ProgramHead: PProgram; Var DistList: TDistList);
Var
  Current: PProgram;
  F: FProgram;
  Item: TProgram;
Begin
  If not FileExists(FileName) then Exit;
  AssignFile(F, FileName);
  Reset(F);

  //clearing program list before creating
  ClearProgramList(ProgramHead);

  //initialization program head fields
  New(ProgramHead);
  ProgramHead^.Swimmers:=Nil;
  ProgramHead^.Next:=Nil;
  Current:=ProgramHead;

  //creating program list
  While not EoF(F) do
  Begin
    New(Current^.Next);
    Current:=Current^.Next;

    //reading data from a file
    Read(F,Item);

    //assign distance fields
    With Current^ do
    Begin
      Meters:=Item.Meters;
      Style:=Item.Style;
      Sex:=Item.Sex;
      maxYear:=Item.maxYear;
      minYear:=Item.minYear;
      SYear:=Item.SYear;
      Swimmers:=Nil;
      Next:=Nil;
    End;
  End;

  //creating unique distation list
  CreateDistList(DistList, ProgramHead);
  uSwimmer.CreateUnsortList(uSwimmer.UnSortList,Length(DistList));
  CloseFile(F);
End;

//saving program to a file
Procedure SaveProgram(FileName: String; ProgramHead: PProgram);
Var
  F: FProgram;
  Current: PProgram;
Begin
  AssignFile(F, FileName);
  Rewrite(F);
  Current:=ProgramHead^.Next;
  While Current<>Nil do
  Begin

    //writing data to a file
    Write(F, Current^);
    Current:=Current^.Next;
  End;
  CloseFile(F);
End;

//creating a program list (from list view)
Procedure CreateProgramList(Var ProgramHead: PProgram; ListView: TListView; Var DistList: TDistList);
Var
  I: Integer;
  Current: PProgram;
Begin

  //clear list structure
  ClearProgramList(ProgramHead);

  //initialization program head of the competition
  New(ProgramHead);
  ProgramHead^.Swimmers:=Nil;
  ProgramHead^.Next:=Nil;
  Current:=ProgramHead;

  //store item fields
  For I:=0 to ListView.Items.Count-1 do
  Begin
    New(Current^.Next);
    Current:=Current^.Next;
    With Current^ do
    Begin
      Meters:=TMeters(GetMetersInd(ListView.Items.Item[I].SubItems[2]));
      Style:=TStyles(GetStyleInd(ListView.Items.Item[I].SubItems[3]));
      Sex:=TSex(GetSexInd(ListView.Items.Item[I].SubItems[0]));
      SYear:=ListView.Items.Item[I].SubItems[1];
      ReadYears(minYear, maxYear, SYear);
      Swimmers:=Nil;
      Next:=Nil;
    End; //with end
  End; //for end

  //creating unique distation list
  CreateDistList(DistList,ProgramHead);

  //creating unsort list
  uSwimmer.CreateUnsortList(uSwimmer.UnSortList,Length(DistList));
End;

//determine unique distations
Procedure CreateDistList(Var DistList: TDistList; Head: PProgram);
Var
  Unique: TDistExist;
  DistCount: Integer;
  I: TMeters;
  J: TStyles;
Begin
  If (Head=Nil) or (Head^.Next=Nil) then Exit;

  //initialize all distances
  For I:=Low(Unique) to High(Unique) do
    For J:=Low(Unique[I]) to High(Unique[I]) do
      Unique[I,J]:=False;

  DistCount:=0;
  Head:=Head^.Next;
  While Head<>Nil do
  Begin

    //finding first occurence of the distation
    If not Unique[Head^.Meters,Head^.Style] then
    Begin
      Unique[Head^.Meters,Head^.Style]:=True;
      Inc(DistCount);
      SetLength(DistList,DistCount);
      DistList[DistCount-1].Meters:=Head^.Meters;
      DistList[DistCount-1].Style:=Head^.Style;
      DistList[DistCount-1].Col:=uError.NOTFOUND;
    End;
    Head:=Head^.Next;
  End;
End;

//getting sex index in array
Function GetSexInd(Str: String): Integer;
Var
  Sex: TSex;
Begin
  Sex:=Low(TSex);
  While Str<>SSex[Sex] do
    Inc(Sex);
  Result:=Ord(Sex);
End;

//getting meters index in array
Function GetMetersInd(Str: String): Integer;
Var
  Meters: TMeters;
Begin
  Meters:=Low(TMeters);
  While Str<>SMeters[Meters] do
    Inc(Meters);
  Result:=Ord(Meters);
End;

//getting style index in array
Function GetStyleInd(Str: String): Integer;
Var
  Style: TStyles;
Begin
  Style:=Low(TStyles);
  While Str<>SStyles[Style] do
    Inc(Style);
  Result:=Ord(Style);
End;

//getting distation index in DistList array
Function GetDistInd(Meters: TMeters; Style: TStyles): Integer;
Var
  I: Integer;
Begin
  Result:=-1;
  For I:=Low(uProgram.DistList) to High(uProgram.DistList) do
    If (DistList[I].Meters=Meters) and (DistList[I].Style=Style) then
    Begin
      Result:=I;
      Exit;
    End;
End;

end.
