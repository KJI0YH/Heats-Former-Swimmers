unit uTechnical;
{the unit responsible for
the technical protocol}

interface

Uses
  SysUtils, Variants, ComObj, VCF1,
  uProgram, uSwimmer;

Const
  SUnsort = 'Спортсмены, не прошедшие сортировку';

Procedure ReadTechnical(FileName: String; Var DistList: TDistList);
Function IsNumber(SNumber: String): Integer;

implementation

uses uMain, uExcel, uError;

{*******************
* local procedures *
********************}

//determine whether the year belong to this range
Function IsYear(Year: TName; minYear, maxYear: Integer): Boolean;
Var
  Number: Integer;
Begin
  Result:=False;
  Number:=StrToInt(Year);
  If (minYear>0) and (maxYear>0) then
    Result:=(Number>=minYear) and (Number<=maxYear)
  Else If (minYear=0) or (maxYear=0) then
    Result:=Number=(minYear or maxYear)
  Else If (minYear=-1) then
    Result:=Number<=maxYear
  Else If (maxYear=-1) then
    Result:=Number>=minYear;
End;

//finding a distance to added swimmer
Function FindDistance(Head: PProgram; Sex: TSex; Meters: TMeters; Style: TStyles; Year: TName): PProgram;
Var
  Current: PProgram;
Begin
  Result:=Nil;
  If Head=Nil then Exit;
  Current:=Head^.Next;
  While (Current<>Nil) and ((Current^.Sex<>Sex) or (Current^.Meters<>Meters) or (Current^.Style<>Style) or (not IsYear(Year,Current^.minYear, Current^.maxYear))) do
    Current:=Current^.Next;
  If Current<>Nil then
    Result:=Current;
End;

//reading swimmers from tecnical list
Procedure ReadSwimmers(FData: OLEVariant; StartRow, RowCount, ColOrder, ColName, ColYear, ColCity: Integer; DistList: TDistList; Sex: TSex);
Var
  I: Integer;
  Distance: PProgram;
Begin
  While (StartRow<=RowCount) and (IsNumber(FData[StartRow,ColOrder+1])=0) do
  Begin
    For I:=Low(DistList) to High(DistList) do
    Begin

      //finding a suitable distance for this swimmer
      Distance:=FindDistance(uProgram.ProgramHead, Sex, DistList[I].Meters, DistList[I].Style, String(FData[StartRow,ColYear+1]));

      //add swimmer to a relevant distance list or to a unsort list
      If (DistList[I].Col<>uError.NOTFOUND) and (Distance<>Nil) then
        uSwimmer.AddSwimmer(Distance^.Swimmers,FData[StartRow,ColName+1],FData[StartRow,ColYear+1],FData[StartRow,ColCity+1], FData[StartRow,DistList[I].Col+1], I)
      Else If (DistList[I].Col<>uError.NOTFOUND) and (String(FData[StartRow,DistList[I].Col+1])<>'') then
        uSwimmer.AddToUnSort(uSwimmer.UnSortList[I],FData[StartRow,ColName+1],FData[StartRow,ColYear+1],FData[StartRow,ColCity+1], FData[StartRow,DistList[I].Col+1], I);
    End;
    Inc(StartRow);
  End;
End;

//determine number
Function IsNumber(SNumber: String): Integer;
Var
  Num: Integer;
Begin
  Val(SNumber,Num,Result);
End;

{***********************
* interface procedures *
************************}

//reading techical list and getting swimmers
Procedure ReadTechnical(FileName: String; Var DistList: TDistList);
Var
  ColOrder, ColName, ColYear, ColCity, RowMale, RowFemale: Integer;
  RowCount, ColCount, I: Integer;
  SDist: String;
  ExcelApp: Variant;
  WorkSheet, FData: OLEVariant;
Begin
  ExcelApp:=CreateOleObject('Excel.Application');
  frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

  //open book
  ExcelApp.Workbooks.Open(FileName);
  ExcelApp.DisplayAlerts:=False;

  //get active page
  WorkSheet:=ExcelApp.ActiveWorkbook.ActiveSheet;
  frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

  //define count rows and columns
  RowCount:=WorkSheet.UsedRange.Rows.Count;
  ColCount:=WorkSheet.UsedRange.Columns.Count;

  //reading data from all used range
  FData:=WorkSheet.UsedRange.Value;
  frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

  //determine row and column index with defined data
  ColOrder:=uExcel.FindKeyWord(uProgram.cOrder, FData, RowCount, ColCount, False);
  ColName:=uExcel.FindKeyWord(uProgram.cName, FData, RowCount, ColCount, False);
  ColYear:=uExcel.FindKeyWord(uProgram.cYear, FData, RowCount, ColCount, False);
  ColCity:=uExcel.FindKeyWord(uProgram.cCity, FData, RowCount, ColCount, False);
  RowMale:=uExcel.FindKeyWord(uProgram.SSex[Male], FData, RowCount, ColCount, True);
  RowFemale:=uExcel.FindKeyWord(uProgram.SSex[Female], FData, RowCount, ColCount, True);
  frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

  //check for errors
  If not uError.ColError([ColOrder, ColName, ColYear, ColCity], FileName, uError.elFatal) then
  Begin

    //define column with distances
    For I:=Low(DistList) to High(DistList) do
    Begin
      SDist:=uProgram.SMeters[DistList[I].Meters]+' '+uProgram.SStyles[DistList[I].Style];
      DistList[I].Col:=uExcel.FindKeyWord(SDist, FData, RowCount, ColCount, False);
      uError.DistError(DistList[I].Col,DistList[I].Meters,DistList[I].Style,FileName,elWarning);
    End;
    frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

    //reading swimmers from technical list
    If not uError.RowError(RowMale, uProgram.SSex[Male], FileName, elWarning) then
      ReadSwimmers(FData,RowMale+2,RowCount,ColOrder,ColName,ColYear,ColCity,DistList,Male);
    frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);

    If not uError.RowError(RowFemale,uProgram.SSex[Female],FileName,elWarning) then
      ReadSwimmers(FData,RowFemale+2,RowCount,ColOrder,ColName,ColYear,ColCity,DistList,Female);
    frmMain.pbProgress.StepBy(frmMain.pbProgress.Step div 7);
  End;

  try
    ExcelApp.Quit;
  except
  end;
  ExcelApp:=Unassigned;
End;

end.

