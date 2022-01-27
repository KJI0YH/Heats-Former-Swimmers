unit uApplicat;
{the unit responsible for
the applicat protocol}

interface

Uses
  Variants, ComObj, Math,
  uProgram;

Const
  cTime = 'ÂÐÅÌß';

Procedure ReadApplicat(FileName: String; ProgramHead: PProgram);

implementation

uses uExcel, uSwimmer, uTechnical, uMain, uError;

{***********************
* interface procedures *
***********************}

//reading applicat protocol from a file
Procedure ReadApplicat(FileName: String; ProgramHead: PProgram);
Var
  ColOrder, ColName, ColYear, ColCity, ColTime: Integer;
  RowCount, ColCount, RowDist: Integer;
  CurrentDist: PProgram;
  HeaderDist: String;
  XL: Variant;
  WorkSheet, FData: OLEVariant;
Begin
  frmMain.pbProgress.Step:=frmMain.pbProgress.Max div (uMain.Order+1);

  XL:=CreateOleObject('Excel.Application');

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

  //try to find unsort string
  RowDist:=uExcel.FindKeyWord(uTechnical.SUnsort, FData, RowCount, ColCount, False);
  If (RowDist<RowCount) and (RowDist<>uError.NOTFOUND) then
    RowCount:=RowDist;

  //determine row and column index with defined data
  ColOrder:=uExcel.FindKeyWord(uProgram.cOrder, FData, RowCount, ColCount, False);
  ColName:=uExcel.FindKeyWord(uProgram.cName, FData, RowCount, ColCount, False);
  ColYear:=uExcel.FindKeyWord(uProgram.cYear, FData, RowCount, ColCount, False);
  ColCity:=uExcel.FindKeyWord(uProgram.cCity, FData, RowCount, ColCount, False);
  ColTime:=uExcel.FindKeyWord(cTime, FData, RowCount, ColCount, False);
  frmMain.pbProgress.StepIt;

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
        RowDist:=RowDist+3;

      If RowDist<>uError.NOTFOUND then
        While (RowDist<=RowCount) and (uTechnical.IsNumber(FData[RowDist,ColOrder+1])=0) do
        Begin
          uSwimmer.AddSwimmer(CurrentDist^.Swimmers, FData[RowDist, ColName+1], FData[RowDist, ColYear+1], FData[RowDist, ColCity+1], FData[RowDist,ColTime+1], uProgram.GetDistInd(CurrentDist^.Meters, CurrentDist^.Style));
          Inc(RowDist);
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

end.






