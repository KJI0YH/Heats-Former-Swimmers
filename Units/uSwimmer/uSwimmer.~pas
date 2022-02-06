unit uSwimmer;
{the unit responsible for
the swimmers}

interface

Uses
  SysUtils;

Type

  //time view structure
  TTime = Record
    Minute: Integer;
    Second: Integer;
    Hungred: Integer;
  End;

  //subtype structure
  TName = String[255];

  //swimmer list structure
  PSwimmer = ^TSwimmer;
  TSwimmer = Record
    Name: TName;
    Year: TName;
    City: TName;
    STime: TName;
    Time: TTime;
    DistInd: Integer;
    Next: PSwimmer;
  End;

  //subtype for heats part
  TLine6 = 1..6;
  TLineOrder = Array [TLine6] of TLine6;

  //heat list structure
  PHeat = ^THeat;
  THeat = Record
    Line: TLine6;
    Swimmer: PSwimmer;
    Next: PHeat;
  End;

  //unsort list structure
  TUnsortList = Array of PSwimmer;

Const

  //the order of filling in the heat
  LineOrder: TLineOrder = (3,4,2,5,1,6);

Var
  UnSortList: TUnsortList;
  strInfinity: TName = '99.99.99';

Procedure AddSwimmer(Var Head: PSwimmer; Name, Year, City, STime: TName; DistInd: Integer);
Procedure AddToUnSort(Var Head: PSwimmer; Name, Year, City, STime: TName; DistInd: Integer);
Function CreateHeat(Var Head: PSwimmer): PHeat;
Procedure ClearHeatList(Head: PHeat);
Procedure CreateUnsortList(Var UnsortList: TUnsortList; Len: Integer);

implementation

{*******************
* local procedures *
*******************}

//getting number from a string (reverse direction)
Function GetNumber(Var Line: TName): String;
Begin
  Result:='';
  While (Length(Line)>0) and ((Line[Length(Line)]<'0') or (Line[Length(Line)]>'9')) do
    Delete(Line, Length(Line), 1);
  While (Length(Line)>0) and ((Line[Length(Line)]>='0') and (Line[Length(Line)]<='9')) do
  Begin
    Result:=Line[Length(Line)]+Result;
    Delete(Line, Length(Line), 1);
  End;
  If Result='' then
    Result:='0';
End;

//convert TTime format into a integer format
Function GetMillisec(Time: TTime): Integer;
Begin
  Result:=Time.Minute*6000+Time.Second*100+Time.Hungred;
End;

//convert string time to TTime view
Function ConvertTime(Var Time: TName): TTime;
Var
  Err: Integer;
  Number: String;
Begin
  Number:=GetNumber(Time);
  Val(Number, Result.Hungred, Err);
  Number:=GetNumber(Time);
  Val(Number, Result.Second, Err);
  Number:=GetNumber(Time);
  Val(Number, Result.Minute, Err);
  Time:='';
  If Result.Minute<>0 then
    Time:=IntToStr(Result.Minute)+'.';
  Time:=Time+FormatFloat('00',Result.Second)+'.'+FormatFloat('00',Result.Hungred);
End;

//checking duplicates swimmers
Function IsEqual (Swimmer: PSwimmer; Name, Year, City: TName; Time: TTime; DistInd: Integer): Boolean;
Begin
  Result:=False;
  If Swimmer=Nil then Exit;
  While (Swimmer<>Nil) and (GetMillisec(Swimmer^.Time)=GetMillisec(Time)) and (not Result) do
  Begin
    Result:=(Swimmer^.Name=Name) and (Swimmer^.Year=Year) and (Swimmer^.City=City) and (GetMillisec(Swimmer^.Time)=GetMillisec(Time)) and (Swimmer^.DistInd=DistInd);
    Swimmer:=Swimmer^.Next;
  End;
End;

{**********************
* inteface procedures *
**********************}

//adding swimmer to a list
Procedure AddSwimmer(Var Head: PSwimmer; Name, Year, City, STime: TName; DistInd: Integer);
Var
  Time: TTime;
  CurrentTime: Integer;
  Current: PSwimmer;
  Item: PSwimmer;
Begin

  //get time
  If STime='' then Exit;
  Time:=ConvertTime(STime);
  If STime='00.00' then
    Time:=ConvertTime(strInfinity);
  CurrentTime:=GetMillisec(Time);

  //create header
  If Head=Nil then
  Begin
    New(Head);
    Head^.Next:=Nil;
  End;
  Current:=Head;

  //search for a place in the sort order
  While (Current^.Next<>Nil) and (CurrentTime>GetMillisec(Current^.Next^.Time)) do
    Current:=Current^.Next;

  //check duplicates
  If not IsEqual(Current^.Next, Name, Year, City, Time, DistInd) then
  Begin
    New(Item);
    Item^.Name:=Name;
    Item^.Year:=Year;
    Item^.City:=City;
    Item^.STime:=STime;
    Item^.Time:=Time;
    Item^.DistInd:=DistInd;
    Item^.Next:=Current^.Next;
    Current^.Next:=Item;
  End;
End;

//adding swimmers who did not qualify for any group
Procedure AddToUnSort(Var Head: PSwimmer; Name, Year, City, STime: TName; DistInd: Integer);
Var
  Item: PSwimmer;
Begin

  //create header
  If Head=Nil then
  Begin
    New(Head);
    Head^.Next:=Nil;
  End;

  New(Item);
  Item^.Name:=Name;
  Item^.Year:=Year;
  Item^.City:=City;
  Item^.STime:=STime;
  Item^.Time:=ConvertTime(STime);
  Item^.DistInd:=DistInd;
  Item^.Next:=Head^.Next;
  Head^.Next:=Item;
End;

//creating heat from a swimmer list
Function CreateHeat(Var Head: PSwimmer): PHeat;
Var
  Line: TLine6;
  Item, Current: PHeat;
Begin

  //heat header initialization
  New(Result);
  Result^.Line:=1;
  Result^.Swimmer:=Nil;
  Result^.Next:=Nil;

  If Head=Nil then Exit;   

  //insert new swimmer into a heat
  Line:=Low(TLine6);
  While (Head<>Nil) and (Line<=High(TLine6)) do
  Begin
    New(Item);
    Item^.Line:=LineOrder[Line];
    Item^.Swimmer:=Head;
    Item^.Next:=Nil;

    Current:=Result;
    While (Current^.Next<>Nil) and (Item^.Line>Current^.Next^.Line) do
      Current:=Current^.Next;

    Item^.Next:=Current^.Next;
    Current^.Next:=Item;

    Inc(Line);
    Head:=Head^.Next;
  End;
End;

//clear heat list
Procedure ClearHeatList(Head: PHeat);
Var
  Deleted: PHeat;
Begin
  While Head<>Nil do
  Begin
    Deleted:=Head;
    Head:=Head^.Next;
    Dispose(Deleted);
  End;
End;

//create unsort list
Procedure CreateUnsortList(Var UnsortList: TUnsortList; Len: Integer);
Var
  I: Integer;
Begin
  SetLength(UnsortList,Len);
  For I:=Low(UnsortList) to High(UnsortList) do
    UnsortList[I]:=Nil;
End;

end.






