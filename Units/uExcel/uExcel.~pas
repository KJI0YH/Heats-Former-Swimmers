unit uExcel;
{the unit responsible for
the Excel representation}

interface

Uses
  uProgram, uSwimmer, uFinish, SysUtils, ComObj, Variants;

Const

  //width of the standart columns
  OrderColSize=6;
  NameColSize=25;
  YearColSize=6;
  CityColSize=24;
  TimeColSize=14;
  PointColSize=14;

  strHeat = 'Заплыв ';
  strDSQ = '00.00';

  strOrder = '№ п.п.';
  strName='Ф.И. спортсмена';
  strYear='Год рожд.';
  strCity='Город';
  strTime='Время';
  strPoint='Очки';

Procedure CreateApplicat(Head: PProgram; XL: Variant);
Procedure CreateStarted(Head: PPRogram; XL: Variant);
Function FindKeyWord(Name: String; FData: OLEVariant; RowCount, ColCount: Integer; ByRows: Boolean): Integer;
Function CreateDistHeader(Sex: TSex; SYear: TSYear; Meters: TMeters; Style: TStyles): String;
Procedure CreateFinish(Head: PProgram; XL: Variant);

implementation

uses uMain, uTechnical, uError;

{*******************
* local procedures *
********************}

//applying constants (column size, alignment, time format)
Procedure ApplyingConst(XL: Variant; WS: Integer);
Var
  I: Integer;
Begin

  //constant size
  XL.WorkSheets[WS].Columns[1].ColumnWidth:=OrderColSize;
  XL.WorkSheets[WS].Columns[2].ColumnWidth:=NameColSize;
  XL.WorkSheets[WS].Columns[3].ColumnWidth:=YearColSize;
  XL.WorkSheets[WS].Columns[4].ColumnWidth:=CityColSize;
  XL.WorkSheets[WS].Columns[5].ColumnWidth:=TimeColSize;
  XL.WorkSheets[WS].Columns[6].ColumnWidth:=PointColSize;
  XL.WorkSheets[WS].Columns[7].ColumnWidth:=PointColSize;
  XL.WorkSheets[WS].Columns[8].ColumnWidth:=PointColSize;


  //time format
  XL.WorkSheets[WS].Columns[5].NumberFormat:='@';

  //alignment
  For I:=1 to 8 do
    XL.WorkSheets[WS].Columns[I].VerticalAlignment:=2;

  XL.WorkSheets[WS].Columns[1].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Columns[2].HorizontalAlignment:=2;
  XL.WorkSheets[WS].Columns[3].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Columns[4].HorizontalAlignment:=2;
  XL.WorkSheets[WS].Columns[5].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Columns[6].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Columns[7].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Columns[8].HorizontalAlignment:=3;

End;

//merging columns and strings in specified range
Procedure AddMergeString(XL: Variant; WS: Integer; Var Row: Integer; Start, Final, Header: String);
Begin
  XL.WorkSheets[WS].Cells[Row, Start]:=Header;

  //merge cells
  XL.WorkSheets[WS].Range[Start+IntToStr(Row)+':'+Final+IntToStr(Row)].Merge;

  //align
  XL.WorkSheets[WS].Rows[Row].VerticalAlignment:=2;
  XL.WorkSheets[WS].Rows[Row].HorizontalAlignment:=3;
  XL.WorkSheets[WS].Rows[Row].Font.Bold:=True;
  Inc(Row);
End;

//adding file header for applicat protocol
Procedure AddHeader(XL: Variant; WS: Integer; Var Row: Integer; Header: String; Cols: Array of String; Dist: Array of String);
Var
  I: Integer;
Begin

  //font
  XL.WorkSheets[WS].Rows[Row].Font.Size:=11;

  //adding merged header
  AddMergeString(XL, WS, Row, 'A', Chr(Ord('A')+Length(Cols)+Length(Dist)-1), Header);

  //align
  XL.WorkSheets[WS].Rows[Row].VerticalAlignment:=2;
  XL.WorkSheets[WS].Rows[Row].HorizontalAlignment:=3;

  //wrap text
  XL.WorkSheets[WS].Rows[Row].WrapText:=True;

  //header text
  For I:=Low(Cols) to High(Cols) do
    XL.WorkSheets[WS].Cells[Row, I+1]:=Cols[I];

  For I:=Low(Dist) to High(Dist) do
    XL.WorkSheets[WS].Cells[Row, Length(Cols)+I+1]:=Dist[I];
  Inc(Row);
End;

//adding swimmer to a Excel table
Procedure AddSwimmer(XL: Variant; WS: Integer; Var Row: Integer; Cols: Array of String; Points: Array of String);
Var
  I: Integer;
Begin
  For I:=Low(Cols) to High(Cols) do
    XL.WorkSheets[WS].Cells[Row, I+1]:=Cols[I];

  For I:=Low(Points) to High(Points) do
    XL.WorkSheets[WS].Cells[Row, Length(Cols)+I+1]:=Points[I];
  Inc(Row);
End;

{***********************
* interface procedures *
***********************}

//create applicat protocol as Excel table
Procedure CreateApplicat(Head: PProgram; XL: Variant);
Var
  Dist: PProgram;
  Swimmer: PSwimmer;
  I, StartRow, CurrRow, SwimmerCount: Integer;
  Header: String;
  UnsortExist: Boolean;
Begin
  XL.WorkSheets[1].Name:='Заявочный протокол';

  //applying constants (column size, alignment, time format)
  ApplyingConst(XL, 1);

  frmMain.pbProgress.Step:=frmMain.pbProgress.Max div (uMain.Order+1);

  Dist:=Head^.Next;
  CurrRow:=1;
  While Dist<>Nil do
  Begin

    //adding distance header
    StartRow:=CurrRow+1;
    Header:=CreateDistHeader(Dist^.Sex,Dist^.SYear,Dist^.Meters,Dist^.Style);
    AddHeader(XL,1,CurrRow,Header,[strOrder,strName,strYear,strCity],[strTime]);
    If (Dist<>Nil) and (Dist^.Swimmers<>Nil) then
      Swimmer:=Dist^.Swimmers^.Next
    Else
      Swimmer:=Nil;
    SwimmerCount:=1;

    //adding swimmers to a Excel table
    While Swimmer<>Nil do
    Begin
      AddSwimmer(XL,1,CurrRow,[IntToStr(SwimmerCount),Swimmer^.Name,Swimmer^.Year,Swimmer^.City,Swimmer^.STime],[]);
      Swimmer:=Swimmer^.Next;
      Inc(SwimmerCount);
    End;

    //drawing cell table lines in Excel
    XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(StartRow+SwimmerCount-1)].Borders.LineStyle:=1;
    XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(StartRow+SwimmerCount-1)].Borders.Weight:=2;

    Dist:=Dist^.Next;
    Inc(CurrRow);
    frmMain.pbProgress.StepIt;
  End;

  //adding unsorted swimmers
  UnsortExist:=False;
  For I:=Low(uSwimmer.UnsortList) to High(uSwimmer.UnsortList) do
    If uSwimmer.UnSortList[I]<>Nil then
    Begin

      If not UnsortExist then
      Begin
        AddMergeString(XL,1,CurrRow,'A','E',uTechnical.SUnsort);
        UnsortExist:=True;
      End;

      StartRow:=CurrRow;

      //adding header for unsort distance
      Header:=uProgram.SMeters[uProgram.DistList[I].Meters]+' '+uProgram.SStyles[uProgram.DistList[I].Style];
      AddMergeString(XL,1,CurrRow,'A','E',Header);
      Swimmer:=uSwimmer.UnSortList[I]^.Next;
      SwimmerCount:=1;

      //adding swimmers to a unsort table
      While Swimmer<>Nil do
      Begin
        AddSwimmer(XL,1,CurrRow,[IntToStr(SwimmerCount),Swimmer^.Name,Swimmer^.Year,Swimmer^.City,Swimmer^.STime],[]);
        Swimmer:=Swimmer^.Next;
        Inc(SwimmerCount);
      End;

      //drawing cell table lines in Excel for unsort table
      XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(CurrRow-1)].Borders.LineStyle:=1;
      XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(CurrRow-1)].Borders.Weight:=2;
    End;
    frmMain.pbProgress.StepIt;
End;

//creating started protocol
Procedure CreateStarted(Head: PPRogram; XL: Variant);
Var
  Dist: PProgram;
  Swimmer: PSwimmer;
  Heat: PHeat;
  CurrRow, CurrHeat, StartRow: Integer;
  Header: String;
Begin
  XL.WorkSheets[1].Name := 'Стартовый протокол';

  //applying constants (column size, alignment, time format)
  ApplyingConst(XL, 1);

  CurrRow:=1;
  CurrHeat:=1;
  StartRow:=CurrRow;

  frmMain.pbProgress.Step:=frmMain.pbProgress.Max div (uMain.Order);

  //reading distances and creating heats
  Dist:=Head^.Next;
  While Dist<>Nil do
  Begin
    Header:=CreateDistHeader(Dist^.Sex,Dist^.SYear,Dist^.Meters,Dist^.Style);
    AddHeader(XL,1,CurrRow,Header,[strOrder,strName,strYear,strCity],[strTime]);
    If (Dist<>Nil) and (Dist^.Swimmers<>Nil) then
      Swimmer:=Dist^.Swimmers^.Next
    Else
      Swimmer:=Nil;

    //creating a heat
    While Swimmer<>Nil do
    Begin
      Heat:=uSwimmer.CreateHeat(Swimmer);
      If Heat^.Next<>Nil then
      Begin
        AddMergeString(Xl,1,CurrRow,'A','E',strHeat+IntToStr(CurrHeat));
        Inc(CurrHeat);
      End;

      //adding swimmer to a Excel table
      While (Heat<>Nil) and (Heat^.Next<>Nil) do
      Begin
        AddSwimmer(XL,1,CurrRow,[IntToStr(Heat^.Next^.Line),Heat^.Next^.Swimmer^.Name,Heat^.Next^.Swimmer^.Year,Heat^.Next^.Swimmer^.City,Heat^.Next^.Swimmer^.STime],[]);
        Heat:=Heat^.Next;
      End;
      uSwimmer.ClearHeatList(Heat);
    End;
    Dist:=Dist^.Next;
    frmMain.pbProgress.StepIt
  End;

  //drawing cell table lines in Excel for started protocol
  XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(CurrRow-1)].Borders.LineStyle:=1;
  XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':E'+IntToStr(CurrRow-1)].Borders.Weight:=2;
End;

//find key word in FData
Function FindKeyWord(Name: String; FData: OLEVariant; RowCount, ColCount: Integer; ByRows: Boolean): Integer;
Var
  I, J: Integer;
  StrValue: String;
Begin
  Result:=uError.NOTFOUND;
  For I:=0 to RowCount-1 do
    For J:=0 to ColCount-1 do
    Begin
      StrValue:=AnsiUpperCase(VarToStr(FData[I+1,J+1]));
      If Pos(AnsiUpperCase(Name),StrValue)<>0 then
      Begin
        If ByRows then
          Result:=I
        Else
          Result:=J;
        Exit;
      End;
    End;
End;

//creating finish protocol
Procedure CreateFinish(Head: PProgram; XL: Variant);
Var
  Dist: PProgram;
  Summary: PSummaryList;
  Swimmer: PSwimmer;
  SwimmerPoint: PSwimmerPoint;
  I, StartRow, CurrRow, SwimmerCount, Points: Integer;
  Header: String;
  strDist, strPoints: Array of String;
Begin
  uFinish.ClearSummary(uFinish.SummaryList);
  
  //creating finish protocol
  XL.WorkSheets[1].Name:='Итоговый протокол';

  //applying constants (column size, alignment, time format)
  ApplyingConst(XL, 1);

  frmMain.pbProgress.Step:=(frmMain.pbProgress.Max div 2) div uMain.Order;

  Dist:=Head^.Next;
  CurrRow:=1;
  While Dist<>Nil do
  Begin
    uFinish.AddSummary(uFinish.SummaryList,Dist);

    //adding distance header
    StartRow:=CurrRow+1;
    Header:=CreateDistHeader(Dist^.Sex,Dist^.SYear,Dist^.Meters,Dist^.Style);
    AddHeader(XL,1,CurrRow,Header,[strOrder,strName,strYear,strCity],[strTime,strPoint]);
    If (Dist<>Nil) and (Dist^.Swimmers<>Nil) then
      Swimmer:=Dist^.Swimmers^.Next
    Else
      Swimmer:=Nil;
    SwimmerCount:=1;

    //adding swimmers to a Excel table
    While Swimmer<>Nil do
    Begin
      If (SwimmerCount>High(uFinish.TPointCount)) or (Swimmer^.STime=strDSQ) then
        Points:=0
      Else
        Points:=uFinish.Points[SwimmerCount];

      If Swimmer^.STime=strDSQ then
        Swimmer^.STime:='диск.';

      AddSwimmer(XL,1,CurrRow,[IntToStr(SwimmerCount),Swimmer^.Name,Swimmer^.Year,Swimmer^.City,Swimmer^.STime],[IntToStr(Points)]);
      uFinish.AddSwimmerPoint(uFinish.SummaryList,Dist,Swimmer,Points);
      Swimmer:=Swimmer^.Next;
      Inc(SwimmerCount);
    End;

    //drawing cell table lines in Excel
    XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':F'+IntToStr(StartRow+SwimmerCount-1)].Borders.LineStyle:=1;
    XL.WorkSheets[1].Range['A'+IntToStr(StartRow)+':F'+IntToStr(StartRow+SwimmerCount-1)].Borders.Weight:=2;

    Dist:=Dist^.Next;
    Inc(CurrRow);
    frmMain.pbProgress.StepIt;
  End;

  frmMain.pbProgress.Position:=frmMain.pbProgress.Max div 2;

  //creating summary list
  XL.Sheets.Add(After:=XL.WorkSheets[1]);
  XL.WorkSheets[2].Name:='Многоборье';
  ApplyingConst(XL, 2);
  CurrRow:=1;

  SetLength(strDist,Length(uProgram.DistList)+1);
  For I:=Low(DistList) to High(DistList) do
    strDist[I]:=uProgram.SMeters[DistList[I].Meters]+' '+uProgram.SStyles[DistList[I].Style];
  strDist[High(strDist)]:=uFinish.strSummary;

  Summary:=uFinish.SummaryList^.Next;
  While Summary<>Nil do
  Begin

    //sort swimmer points
    uFinish.SortPoints(Summary^.SwimmerPoint);

    //adding distance header
    StartRow:=CurrRow+1;
    Header:=uProgram.SSex[Summary^.Sex]+' '+Summary^.SYear+' г.р.';
    AddHeader(XL,2,CurrRow,Header,[strOrder,strName,strYear,strCity],strDist);
    If (Summary<>Nil) and (Summary^.SwimmerPoint<>Nil) then
      SwimmerPoint:=Summary^.SwimmerPoint
    Else
      SwimmerPoint:=Nil;
    SwimmerCount:=1;

    //adding swimmers to a Excel table
    While SwimmerPoint<>Nil do
    Begin

      //creating string values of points
      SetLength(strPoints,Length(SwimmerPoint^.Point)+1);
      Points:=0;
      For I:=Low(SwimmerPoint^.Point) to High(SwimmerPoint^.Point) do
      Begin
        Points:=Points+SwimmerPoint^.Point[I];
        strPoints[I]:=IntToStr(SwimmerPoint^.Point[I]);
      End;
      strPoints[High(strPoints)]:=IntToStr(Points);

      AddSwimmer(XL,2,CurrRow,[IntToStr(SwimmerCount),SwimmerPoint^.Swimmer^.Name,SwimmerPoint^.Swimmer^.Year,SwimmerPoint^.Swimmer^.City],strPoints);
      SwimmerPoint:=SwimmerPoint^.Next;
      Inc(SwimmerCount);
    End;

    //drawing cell table lines in Excel
    XL.WorkSheets[2].Range['A'+IntToStr(StartRow)+':'+Chr(Ord('A')+4+Length(strPoints)-1)+IntToStr(StartRow+SwimmerCount-1)].Borders.LineStyle:=1;
    XL.WorkSheets[2].Range['A'+IntToStr(StartRow)+':G'+IntToStr(StartRow+SwimmerCount-1)].Borders.Weight:=2;

    Summary:=Summary^.Next;
    Inc(CurrRow);
    frmMain.pbProgress.StepIt;
  End;
End;

//creating distation header string
Function CreateDistHeader(Sex: TSex; SYear: TSYear; Meters: TMeters; Style: TStyles): String;
Begin
  Result:=uProgram.SSex[Sex]+' '+SYear+' г.р. '+uProgram.SMeters[Meters]+' '+uProgram.SStyles[Style];
End;

end.

