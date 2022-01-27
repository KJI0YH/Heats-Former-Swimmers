unit uError;
{the unit responsible for
the error messages}

interface

Uses
  uProgram;

Const
  NOTFOUND = -1;

Type
  TErrorLevel = (elFatal, elWarning);
  TErrorInfo = String[255];

  //error list structure
  PErrorList = ^TErrorList;
  TErrorList = Record
    ErrorInfo: TErrorInfo;
    FileName: TErrorInfo;
    ErrorLevel: TErrorLevel;
    Next: PErrorList;
  End;

  TSErrorLevel = Array [TErrorLevel] of String;
  TCol = 0..4;
  TSCol = Array [TCol] of String;

Const

  //string values of types
  SErrorLevel: TSErrorLevel = ('Фатальный', 'Предупреждение');
  SCol: TSCol = ('№ п.п.','Ф.И. спортсмена','Год рожд.','Город','Время');

Var
  ErrorList: PErrorList = Nil;

Function ColError(Ind: Array of Integer; FileName: String; ErrorLevel: TErrorLevel): Boolean;
Function RowError (Ind: Integer; Row, FileName: String; ErrorLevel: TErrorLevel): Boolean;
Function DistError(Ind: Integer; Meters: uProgram.TMeters; Style: TStyles; FileName: String; ErrorLevel: TErrorLevel): Boolean;

implementation

{*******************
* local procedures *
*******************}

//creating header of error list
Procedure CreateHeader(Var Head: PErrorList);
Begin
  If Head=Nil then
  Begin
    New(Head);
    Head^.ErrorInfo:='';
    Head^.FileName:='';
    Head^.ErrorLevel:=elFatal;
    Head^.Next:=Nil;
  End;
End;

//adding item to error list
Procedure AddItem(Var Head: PErrorList; ErrorInfo, FileName: TErrorInfo; ErrorLevel: TErrorLevel);
Var
  Item: PErrorList;
Begin
  CreateHeader(Head);

  New(Item);
  Item^.ErrorInfo:=ErrorInfo;
  Item^.FileName:=FileName;
  Item^.ErrorLevel:=ErrorLevel;

  Item^.Next:=Head^.Next;
  Head^.Next:=Item;
End;

{***********************
* interface procedures *
***********************}

//generate columns error text
Function ColError(Ind: Array of Integer; FileName: String; ErrorLevel: TErrorLevel): Boolean;
Var
  I: Integer;
Begin
  Result:=False;
  For I:=Low(Ind) to High(Ind) do
    If (Ind[I]=NOTFOUND) and (I<=High(TCol)) then
    Begin
      Result:=True;
      AddItem(ErrorList,'Не найден столбец: '+SCol[I],FileName,ErrorLevel);
    End;
End;

//generate row error text
Function RowError (Ind: Integer; Row, FileName: String; ErrorLevel: TErrorLevel): Boolean;
Begin
  Result:=False;
  If (Ind=NOTFOUND) then
  Begin
    Result:=True;
    AddItem(ErrorList,'Не найдена строка: '+Row, FileName,ErrorLevel);
  End;
End;

//generate dist error text
Function DistError(Ind: Integer; Meters: uProgram.TMeters; Style: TStyles; FileName: String; ErrorLevel: TErrorLevel): Boolean;
Begin
  Result:=False;
  If Ind=NOTFOUND then
  Begin
    Result:=True;
    AddItem(ErrorList,'Не найден столбец: '+uProgram.SMeters[Meters]+' '+uProgram.SStyles[Style],FileName,ErrorLevel);
  End;
End;

end.

