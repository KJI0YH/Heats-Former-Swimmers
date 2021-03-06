unit uFinish;
{the unit responsible for the
finish protocol}

interface

Uses
  uSwimmer, uProgram, Variants, ComObj, Math;

Type

  TDistPoint = Array of Integer;
  PSwimmerPoint = ^TSwimmerPoint;
  TSwimmerPoint = Record
    Swimmer: PSwimmer;
    Point: TDistPoint;
    FINAPoint: TDistPoint;
    Next: PSwimmerPoint;
  End;

  PSummaryList = ^TSummaryList;
  TSummaryList = Record
    Sex: TSex;
    SYear: TSYear;
    SwimmerPoint: PSwimmerPoint;
    Next: PSummaryList;
  End;

  TPointCount = 1..20;
  TPoints = Array [TPointCount] of Integer;

  TFINAWR = Array [TSex, TMeters, TStyles] of TName;

Const
  Points: TPoints = (50,45,40,36,32,28,25,22,19,16,14,12,10,8,6,5,4,3,2,1);

  strSummary = '?????';

Var
  SummaryList: PSummaryList = Nil;

                          //male
  FINAWR: TFINAWR =      ((//fly       //back       //breast    //free      //medley
                          ('',         '',          '',         '',         ''),         //25m
                          ('21.75',    '22.22',     '24.95',    '20.16',    ''),         //50m
                          ('47.78',    '48.33',     '55.28',    '44.48',    '49.28'),    //100m
                          ('1.48.24',  '1.45.63',   '2.00.16',  '1.39.37',  '1.49.63'),  //200m
                          ('',         '',          '',         '3.32.25',  '3.54.81'),  //400m
                          ('',         '',          '',         '7.23.42',  ''),         //800m
                          ('',         '',          '',         '14.06.88', '')          //1500m
  ),
                          //female
                          (//fly       //back       //breast    //free      //medley
                          ('',         '',          '',         '',         ''),         //25m
                          ('24.38',    '25.27',     '28.56',    '22.93',    ''),         //50m
                          ('54.59',    '54.89',     '1.02.36',  '50.25',    '56.51'),    //100m
                          ('1.59.61',  '1.58.94',   '2.14.57',  '1.50.31',  '2.01.86'),  //200m
                          ('',         '',          '',         '3.53.92',  '4.18.94'),  //400m
                          ('',         '',          '',         '7.59.34',  ''),         //800m
                          ('',         '',          '',         '15.18.01', '')          //1500m
  ));

Procedure ReadStarted(FileName: String; Var ProgramHead: PProgram);
Procedure AddSummary(Var SummaryList: PSummaryList; Dist: PProgram);
Procedure AddSwimmerPoint(Var SummaryList: PSummaryList; Dist: PProgram; Swimmer: PSwimmer; Point, FINAPoint: Integer);
Procedure SortPoints(Var Head: PSwimmerPoint; FINA: Boolean = False);
Procedure ClearSummary(Var SummaryList: PSummaryList);
Function FINAPoints(Sex: TSex; Meters: TMeters; Style: TStyles; STime: TName): Integer;

implementation

uses uError, uTechnical, uExcel, uApplicat, uMain;

{*******************
* local procedures *
*******************}

//check equality of two swimmers
Function Equal(Swimmer1, Swimmer2: PSwimmer): Boolean;
Begin
  Result:=(Swimmer1^.Name=Swimmer2^.Name) and (Swimmer1^.City=Swimmer2^.City);
End;

//get points summary
Function GetPoints(SwimmerPoint: PSwimmerPoint; FINA: Boolean): Integer;
Var
  I: Integer;
Begin
  Result:=0;
  If FINA then
    For I:=Low(SwimmerPoint^.FINAPoint) to High(SwimmerPoint^.FINAPoint) do
      Result:=Result+SwimmerPoint^.FINAPoint[I]
  Else
    For I:=Low(SwimmerPoint^.Point) to High(SwimmerPoint^.Point) do
      Result:=Result+SwimmerPoint^.Point[I];
End;

//clearing swimmerpoint list
Procedure ClearSwimmerPoint(Var Head: PSwimmerPoint);
Var
  Current, Deleted: PSwimmerPoint;
Begin
  If Head=Nil then Exit;
  Current:=Head^.Next;
  While Current<>Nil do
  Begin
    Deleted:=Current;
    Current:=Current^.Next;
    Dispose(Deleted);
  End;
  Head:=Nil;
End;

{***********************
* interface procedures *
***********************}

//clearing summary list
Procedure ClearSummary(Var SummaryList: PSummaryList);
Var
  Current, Deleted: PSummaryList;
Begin
  If SummaryList=Nil then Exit;
  Current:=SummaryList^.Next;
  While Current<>Nil do
  Begin
    Deleted:=Current;
    Current:=Current^.Next;
    Dispose(Deleted);
  End;
  Dispose(SummaryList);
  SummaryList:=Nil;

End;

//reading started protocol with finish results
Procedure ReadStarted(FileName: String; Var ProgramHead: PProgram);
Var
  XL: Variant;
  WorkSheet, FData: OLEVariant;
  CurrentDist: PProgram;
  ColOrder, ColName, ColYear, ColCity, ColTime: Integer;
  RowCount, ColCount, RowDist: Integer;
  HeaderDist: String;
Begin

  //clear swimmer list
  CurrentDist:=ProgramHead^.Next;
  While CurrentDist<>Nil do
  Begin
    uProgram.ClearSwimmerList(CurrentDist^.Swimmers);
    CurrentDist:=CurrentDist^.Next;
  End;

  XL:=CreateOleObject('Excel.Application');

  frmMain.pbProgress.Step:=frmMain.pbProgress.Max div uMain.Order;

  //open book
  XL.Workbooks.Open(FileName);
  XL.DisplayAlerts:=False;

  //get active page
  WorkSheet:=XL.ActiveWorkbook.ActiveSheet;

  //reading data from all used range
  FData:=WorkSheet.UsedRange.Value;

  //define count rows and columns
  RowCount:=WorkSheet.UsedRange.Rows.Count;
  ColCount:=WorkSheet.UsedRange.Columns.Count;

  //determine row and column index with defined data
  ColOrder:=uExcel.FindKeyWord(uProgram.cOrder, FData, RowCount, ColCount, False);
  ColName:=uExcel.FindKeyWord(uProgram.cName, FData, RowCount, ColCount, False);
  ColYear:=uExcel.FindKeyWord(uProgram.cYear, FData, RowCount, ColCount, False);
  ColCity:=uExcel.FindKeyWord(uProgram.cCity, FData, RowCount, ColCount, False);
  ColTime:=uExcel.FindKeyWord(uApplicat.cTime, FData, RowCount, ColCount, False);

  //check for errors
  If not uError.ColError([ColOrder,ColName,ColYear,ColCity,ColTime], FileName, uError.elFatal) then
  Begin

    //reading swimmers from applicat list
    CurrentDist:=ProgramHead^.Next;
    While CurrentDist<>Nil do
    Begin

      //create and find header with distation information
      HeaderDist:=uExcel.CreateDistHeader(CurrentDist^.Sex,CurrentDist^.SYear,CurrentDist^.Meters,CurrentDist^.Style);
      RowDist:=uExcel.FindKeyWord(HeaderDist,FData,RowCount,ColCount,True);

      If not uError.RowError(RowDist,HeaderDist,FileName,uError.elWarning) then
      Begin
        RowDist:=RowDist+3;

        If RowDist<>uError.NOTFOUND then
          While (RowDist<=RowCount) and ((uTechnical.IsNumber(FData[RowDist,ColOrder+1])=0) or (Pos(uExcel.strHeat,VarToStr(FData[RowDist,1]))<>0)) do
          Begin
            If Pos(uExcel.strHeat,VarToStr(FData[RowDist,1]))=0 then
              uSwimmer.AddSwimmer(CurrentDist^.Swimmers, FData[RowDist, ColName+1], FData[RowDist, ColYear+1], FData[RowDist, ColCity+1], FData[RowDist,ColTime+1], uProgram.GetDistInd(CurrentDist^.Meters, CurrentDist^.Style));
            Inc(RowDist);
          End;
      End;
      CurrentDist:=CurrentDist^.Next;
      frmMain.pbProgress.StepIt;
    End;
  End;

  try
    XL.Quit;
  except
  end;
  XL:=Unassigned;
End;

//adding category for summary
Procedure AddSummary(Var SummaryList: PSummaryList; Dist: PProgram);
Var
  Prev, Current: PSummaryList;
  Found: Boolean;
Begin

  //creating header
  If SummaryList=Nil then
  Begin
    New(SummaryList);
    SummaryList^.Sex:=Low(uProgram.TSex);
    SummaryList^.SYear:='';
    SummaryList^.SwimmerPoint:=Nil;
    SummaryList^.Next:=Nil;
  End;

  //try to found category
  Prev:=SummaryList;
  Current:=SummaryList^.Next;
  Found:=False;
  While (Current<>Nil) and not Found do
    If (Current^.Sex<>Dist^.Sex) or (Current^.SYear<>Dist^.SYear) then
    Begin
      Prev:=Current;
      Current:=Current^.Next;
    End
    Else
      Found:=True;

  //creating new category
  If not Found then
  Begin
    New(Prev^.Next);
    Current:=Prev^.Next;
    Current^.Sex:=Dist^.Sex;
    Current^.SYear:=Dist^.SYear;
    Current^.SwimmerPoint:=Nil;
    Current^.Next:=Nil;
  End;
End;

//add swimmer and point for the distation
Procedure AddSwimmerPoint(Var SummaryList: PSummaryList; Dist: PProgram; Swimmer: PSwimmer; Point, FINAPoint: Integer);
Var
  Current: PSummaryList;
  CurrSwimmer: PSwimmerPoint;
  Found: Boolean;
Begin

  //try to find category
  Found:=False;
  Current:=SummaryList^.Next;
  While (Current<>Nil) and not Found do
    If (Current^.Sex<>Dist^.Sex) or (Current^.SYear<>Dist^.SYear) then
      Current:=Current^.Next
    Else
      Found:=True;

  If Found then
  Begin

    //try to found swimmer
    CurrSwimmer:=Current^.SwimmerPoint;
    Found:=False;
    While (CurrSwimmer<>Nil) and not Found do
      If Equal(CurrSwimmer^.Swimmer,Swimmer) then
      Begin
        Found:=True;
        CurrSwimmer^.Point[uProgram.GetDistInd(Dist^.Meters,Dist^.Style)]:=Point;
        CurrSwimmer^.FINAPoint[uProgram.GetDistInd(Dist^.Meters,Dist^.Style)]:=FINAPoint;
      End
      Else
        CurrSwimmer:=CurrSwimmer^.Next;

      //create new swimmer
      If not Found then
      Begin
        New(CurrSwimmer);
        CurrSwimmer^.Swimmer:=Swimmer;
        SetLength(CurrSwimmer^.Point,Length(uProgram.DistList));
        SetLength(CurrSwimmer^.FINAPoint,Length(uProgram.DistList));
        CurrSwimmer^.Point[uProgram.GetDistInd(Dist^.Meters,Dist^.Style)]:=Point;
        CurrSwimmer^.FINAPoint[uProgram.GetDistInd(Dist^.Meters,Dist^.Style)]:=FINAPoint;
        CurrSwimmer^.Next:=Current^.SwimmerPoint;
        Current^.SwimmerPoint:=CurrSwimmer;
      End;
  End;
End;

//sort swimmer points (selection sort)
Procedure SortPoints(Var Head: PSwimmerPoint; FINA: Boolean = False);
Var
  I, J, Current: PSwimmerPoint;
  TempSwimmer: PSwimmer;
  TempPoint, TempPointFINA: TDistPoint;
Begin
  I:=Head;
  While I<>Nil do
  Begin
    Current:=I;
    If I^.Next<>Nil then
    Begin
      J:=I^.Next;
      While J<>Nil do
      Begin
        If GetPoints(J,FINA)>GetPoints(Current,FINA) then
          Current:=J;
        J:=J^.Next;
      End;
    End;

    TempSwimmer:=Current^.Swimmer;
    TempPoint:=Current^.Point;
    TempPointFINA:=Current^.FINAPoint;
    Current^.Swimmer:=I^.Swimmer;
    Current^.Point:=I^.Point;
    Current^.FINAPoint:=I^.FINAPoint;
    I^.Swimmer:=TempSwimmer;
    I^.Point:=TempPoint;
    I^.FINAPoint:=TempPointFINA;

    I:=I^.Next;
  End;
End;

//calculating FINA points
Function FINAPoints(Sex: TSex; Meters: TMeters; Style: TStyles; STime: TName): Integer;
Begin
  try
    Result:=Trunc(Power(uSwimmer.GetMillisec((uSwimmer.ConvertTime(FINAWR[Sex,Meters,Style])))/uSwimmer.GetMillisec(uSwimmer.ConvertTime(STime)),3)*1000);
  except
    Result:=0;
  End;
End;

end.

