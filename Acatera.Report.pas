unit Acatera.Report;

interface

uses
  Variants, Contnrs, DB, SysUtils, Classes;

type
  TDataType = (dtNumber, dtString);
  TAlign = (alLeft, alRight, alCenter);

  TCellProperty = record
    Size: integer;
    Align: TAlign;
    DataType: TDataType;
    Formatting: string;
  end;
  PCellProperty = ^TCellProperty;

  TCols = array of PCellProperty;

  TRows = array of PCellProperty;

  TCells = array of string;

  TBand = class;
  TBands = class(TObjectList);

  TWorkNotifyEvent = procedure(WorkDone, TotalWork: Integer) of object;
  TReport = class
  private
    fActualRowCount: Integer;
    fBands: TBands;
    fCells: TCells;
    fCols: TCols;
    fRows: TRows;
    fRowCount: integer;
    fColCount: integer;
    fDataSet: TDataSet;
    fParams: TStringList;
    function GetCell(Col, Row: Integer): string;
    procedure SetCell(Col, Row: Integer; Value: string);
    procedure SetCapacity(Cols, Rows: Integer);
    procedure InitializeCols;
    procedure InitializeRows;
    function GetCol(Index: SmallInt): PCellProperty;
    function GetRow(Index: SmallInt): PCellProperty;
    procedure SetCol(Index: SmallInt; const Value: PCellProperty);
    procedure SetBandEvents;
    function FormatCell(X, Y: Integer): string;
    procedure FillParams;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
    constructor Create(Cols, Rows: integer);
    procedure Execute;
    procedure RenderToFile(FileName: string);
    procedure GrowRowCount(Value: integer);
    property Cells[Col, Row: Integer]: string read GetCell write SetCell;
    property Cols[Index: SmallInt]: PCellProperty read GetCol write SetCol; default;
    property Rows[Index: SmallInt]: PCellProperty read GetRow;
  published  
    property Bands: TBands read fBands;
    property ColCount: integer read fColCount;
    property DataSet: TDataSet read fDataSet write fDataSet;
    property RowCount: integer read fRowCount;
  end;

  TBandEvent = procedure(Event: TObject) of object;

  TVarArray = array of Variant;
  
  TAccumulatorOp = (acNone, acAddition, acSubtraction, acMultiply, acDivide, acPower, acRoot);
  
  TAccumulator = class
  strict private
    fDisplayName: string;
    fDriverColumns: TBytes;
    fIsInitialized: boolean;
    fIsMarkedForRendering: Boolean;
    fOperation: TAccumulatorOp;
    fSubtotalColumns: TBytes;
  private  
    fLastCodeValue: Variant; //need it here for the other classes in this unit to access;
    fValues: TVarArray;      //same as above;
    procedure SetSubtotalColumns(const Value: TBytes);
    procedure SetDriverColumns(const Value: TBytes);
  public
    procedure AddValue(ColID: Byte; Value: Variant); 
    function IsInitialized: Boolean;
    procedure Initialized;
    function IsMarkedForRendering: Boolean;
    procedure MarkForRendering;
    procedure UnmarkForRendering;
  published
    property DisplayName: string read fDisplayName write fDisplayName;
    property DriverColumns: TBytes read fDriverColumns write SetDriverColumns;
    property Operation: TAccumulatorOp read fOperation write fOperation;
    property SubtotalColumns: TBytes read fSubtotalColumns write SetSubtotalColumns;
  end;
  
  TAccumulatorList = class(TObjectList);
  
  TOccurence = (atNewPage, atChangeID, atEachRow, atPageEnd);
  
  TBand = class
  strict private
    fOccurence: TOccurence;
    fOnAfterRender: TBandEvent;
    fOnBeforeRender: TBandEvent;
    fOnSubtotalColChange: TBandEvent; 
    fOnWorkProgress: TWorkNotifyEvent;
    fAccumulators: TAccumulatorList;
  protected 
    fDataSet: TDataSet;
    fOwner: TReport;
    function FormatDiplayText(Str: string): string;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
    procedure Execute(Data: TObject); virtual; abstract;
  published 
    property Accumulators: TAccumulatorList read fAccumulators write fAccumulators;
    property Occurence: TOccurence read fOccurence write fOccurence;
    property OnAfterRender: TBandEvent read fOnAfterRender write fOnAfterRender;
    property OnBeforeRender: TBandEvent read fOnBeforeRender write fOnBeforeRender;
    property OnSubtotalColChange: TBandEvent read fOnSubtotalColChange write fOnSubtotalColChange;
    property OnWorkProgress: TWorkNotifyEvent read fOnWorkProgress write fOnWorkProgress;
  end;

  TDetailBand = class(TBand)
  strict private
    procedure InitializeAccumulator(Acc: TAccumulator);
    procedure AddValuesToAccumutator(Acc: TAccumulator);
    function GetCodeValueForAccumulator(Acc: TAccumulator): string;
    procedure AppendParams(Params: TStringList);
  public  
    constructor Create(Owner: TReport);
    procedure Execute(Data: TObject); override;
  end;
  
  TTitleBand = class(TBand)
    constructor Create(Owner: TReport);
    procedure Execute(Data: TObject); override;
  end;

  TSubtitleBand = class(TBand)
    constructor Create(Owner: TReport);
    procedure Execute(Data: TObject); override;
  end;

implementation

uses
  StrUtils, Acatera.Utils, Windows;

{ TReport }

procedure TReport.AfterConstruction;
begin
  fBands := TBands.Create(True);
  fParams := TStringList.Create;
end;

procedure TReport.BeforeDestruction;
begin
  fBands.Free;
  fParams.Free;
end;

constructor TReport.Create(Cols, Rows: integer);
begin
  SetCapacity(Cols, Rows);
end;

procedure TReport.FillParams;
var
  i: integer;
begin
  //Column names
  if (fDataSet <> nil) then begin
    for i := 0 to fDataSet.FieldCount - 1 do
      fParams.Add(Format('%%columnNames[%d]%%=%s', [i, fDataSet.Fields[i].DisplayName]));
  end;
end;

function TReport.FormatCell(X, Y: Integer): string;
var
  Value: double;
begin
  if ((fCols[X].Formatting <> '')) then begin
    if (fCols[X].DataType = dtNumber) then begin
      Value := StrToFloatDef(GetCell(X, Y), MAXDWORD);
      if (Value <> MAXDWORD) then
        Result := Format(fCols[X].Formatting, [Value])
      else 
        Result := GetCell(X, Y);  
    end else 
      Result := Format(fCols[X].Formatting, [GetCell(X, Y)]);
  end else
    Result := GetCell(X, Y);
end;

function TReport.GetCell(Col, Row: Integer): string;
begin
  if (Row > fRowCount - 1) or (Row < 0) or (Col > fColCount - 1) or (Col < 0) then
    Exit;
  Result := fCells[Row * fColCount + Col];
end;

function TReport.GetCol(Index: SmallInt): PCellProperty;
begin
  if ((Index >= 0) and (Index <= fColCount - 1)) then
    Result := fCols[Index];
end;

function TReport.GetRow(Index: SmallInt): PCellProperty;
begin
  if ((Index >= 0) and (Index <= fColCount - 1)) then
    Result := fRows[Index];
end;

procedure TReport.GrowRowCount(Value: integer);
var
  Delta: integer;
  GrowBy: integer;
begin
  if ((fRowCount + Value) > fActualRowCount) then begin
    Delta := fRowCount + Value - fActualRowCount;
    GrowBy := (Round(Delta / 500) + 1) * 500;
    Inc(fActualRowCount, GrowBy);
    Inc(fRowCount, Value);
    SetLength(fCells, fColCount * (fActualRowCount));
    SetLength(fRows, fActualRowCount);
    InitializeRows;
  end else begin
    Inc(fRowCount, Value);
  end;
end;

procedure TReport.Execute;
var
  i: Integer;
begin
  FillParams;
  for i := 0 to fBands.Count - 1 do begin
    if (TBand(fBands[i]).Occurence = atEachRow) then
    begin
      TBand(fBands[i]).Execute(fDataSet);
    end;
  end;
end;

procedure TReport.RenderToFile(FileName: string);
var
  Buffer: string;
  Y: Integer;
  X: Integer;
  Output: TStringList;
begin
  if (Trim(FileName) <> '') then begin
    Output := TStringList.Create;
    try
      Output.Capacity := fRowCount;
      for Y := 0 to fRowCount - 1 do begin
        Buffer := '';
        for X := 0 to fColCount - 1 do begin
          if (fCols[X].Align = alLeft) then
            Buffer := Buffer + RPad(FormatCell(X, Y), fCols[X].Size) + '|'
          else if (fCols[X].Align = alRight) then
            Buffer := Buffer + LPad(FormatCell(X, Y), fCols[X].Size) + '|'
          else
            Buffer := Buffer + CenterStr(FormatCell(X, Y), fCols[X].Size) + '|';
        end;
        Output.Add(Buffer);
      end;
      for Y := 0 to fBands.Count - 1 do
        TBand(fBands[Y]).fDataSet := nil;
      Output.SaveToFile(FileName);
    finally
      Output.Free;
    end;
  end;
end;

procedure TReport.SetBandEvents;
var
  i: SmallInt;
begin
  for i := 0 to fBands.Count - 1 do begin
    if (TBand(fBands[i]).Occurence = atEachRow) then
      if ((i > 0) and (TBand(fBands[i - 1]).Occurence = atNewPage)) then
        TBand(fBands[i]).OnBeforeRender := TBand(fBands[i - 1]).Execute;
  end;
end;

procedure TReport.SetCapacity(Cols, Rows: Integer);
begin
  SetLength(fCells, Rows * Cols);
  SetLength(fCols, Cols);
  SetLength(fRows, Rows);
  fRowCount := Rows;
  fColCount := Cols;
  InitializeCols;
  InitializeRows;
end;

procedure TReport.InitializeCols;
var
  i: integer;
begin
  for i := 0 to fColCount - 1 do begin
    New(fCols[i]);
    fCols[i].Align := alLeft;
    fCols[i].Size := 16;
    fCols[i].DataType := dtString;
  end;
end;

procedure TReport.InitializeRows;
var
  i: integer;
begin
  for i := 0 to fActualRowCount - 1 do begin
    New(fRows[i]);
    fRows[i].Size := 1;
  end;
end;

procedure TReport.SetCell(Col, Row: Integer; Value: string);
begin
  if (Row > fRowCount - 1) or (Row < 0) or (Col > fColCount - 1) or (Col < 0) then
    Exit;

  fCells[Row * fColCount + Col] := Value;
end;

procedure TReport.SetCol(Index: SmallInt; const Value: PCellProperty);
begin
  if ((Index >= 0) and (Index <= fColCount - 1)) then
    fCols[Index] := Value;
end;

{ TBand }

procedure TDetailBand.AppendParams(Params: TStringList);
var
  i: Integer;
begin
  if ((Params <> nil) and (fDataSet <> nil)) then begin
    for i := 0 to fDataSet.FieldCount - 1 do
      Params.Values[Format('%%lastValues[%d]%%', [i])] := fDataSet.Fields[i].AsString;
  end;
end;

constructor TDetailBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  Occurence := atEachRow;
end;

procedure TDetailBand.Execute(Data: TObject);
var
  i, j: Integer;
  ShouldRenderSubtotal: Boolean;
  Acc: TAccumulator;
  RecNo: integer;
begin
  if (Assigned(OnBeforeRender)) then
    OnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TDataSet) then begin
    fDataSet := TDataSet(Data);
    
    RecNo := 0;
    while (not fDataSet.Eof) do begin
      for i := 0 to Accumulators.Count - 1 do begin //Initialize and add values;
        Acc := TAccumulator(Accumulators[i]);
        if (not Acc.IsInitialized) then begin
          InitializeAccumulator(Acc);
        end;
        AddValuesToAccumutator(Acc);
      end;
    
      fOwner.GrowRowCount(1); //Render row;
      for j := 0 to fDataSet.FieldCount - 1 do begin
        fOwner.Cells[j, fOwner.RowCount - 1] := fDataSet.Fields[j].AsString;
      end;

      AppendParams(fOwner.fParams);
      
      fDataSet.Next;
      
      for i := 0 to Accumulators.Count - 1 do begin
        Acc := TAccumulator(Accumulators[i]);
        if (Acc.fLastCodeValue <>  GetCodeValueForAccumulator(Acc)) then begin
          Acc.MarkForRendering;
          ShouldRenderSubtotal := True;
          Acc.fLastCodeValue := GetCodeValueForAccumulator(Acc);
        end;
      end;
      if (Assigned(OnSubtotalColChange) and (ShouldRenderSubtotal)) then begin
        OnSubtotalColChange(Accumulators);
        ShouldRenderSubtotal := False;
      end;
      if (Assigned(OnWorkProgress)) then //OnWorkDone;
        OnWorkProgress(RecNo, fDataSet.RecordCount);
      Inc(RecNo);
    end;

    if (Assigned(OnWorkProgress)) then //OnWorkDone;
      OnWorkProgress(RecNo, fDataSet.RecordCount);
      
    fDataSet := nil;
  end;

  if (Assigned(OnSubtotalColChange)) then //Last subtotal
    OnSubtotalColChange(Accumulators);
      
  if (Assigned(OnAfterRender)) then //End of document, etc
    OnAfterRender(Data);
end;

function TDetailBand.GetCodeValueForAccumulator(Acc: TAccumulator): string;
var
  i: Byte;
begin
  Result := '';
  for i := 0 to Length(Acc.DriverColumns) - 1 do begin
    Result := Result + fDataSet.Fields[Acc.DriverColumns[i]].AsString;
  end;
end;

procedure TDetailBand.AddValuesToAccumutator(Acc: TAccumulator);
var
  i: integer;
begin
  for i := 0 to Length(Acc.SubtotalColumns) - 1 do begin
    Acc.AddValue(i, fDataSet.Fields[Acc.SubtotalColumns[i]].AsVariant);
  end;
end;

procedure TDetailBand.InitializeAccumulator(Acc: TAccumulator);
begin
  Acc.fLastCodeValue := GetCodeValueForAccumulator(Acc);
  Acc.Initialized;
end;

{ TBand }

procedure TBand.AfterConstruction;
begin
  fAccumulators := TAccumulatorList.Create(True);  
end;

procedure TBand.BeforeDestruction;
begin
  fAccumulators.Free;
end;

function TBand.FormatDiplayText(Str: string): string;
var
  i: Integer;
begin
  Result := Trim(Str);
  
  if ((Pos('%', Result) = 0) or (Result = '')) then
    Exit;
  
  for i := 0 to fOwner.fParams.Count - 1 do begin
    Result := StringReplace(Result, fOwner.fParams.Names[i], fOwner.fParams.ValueFromIndex[i], [rfReplaceAll, rfIgnoreCase]);
  end;
end;

{ TTitleBand }

constructor TTitleBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  Occurence := atNewPage;
end;

procedure TTitleBand.Execute(Data: TObject);
var
  i: Integer;
begin
  if (Assigned(OnBeforeRender)) then
    OnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TDataSet) then begin
    fDataSet := TDataSet(Data);
    fOwner.GrowRowCount(1);
    for i := 0 to fDataSet.FieldCount - 1 do begin
      fOwner.Cells[i, fOwner.RowCount - 1] := fDataSet.Fields[i].DisplayName;
    end;
  end;

  if (Assigned(OnAfterRender)) then
    OnAfterRender(Data);
end;

{ TSubtitleBand }

constructor TSubtitleBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  Occurence := atChangeID;
end;

procedure TSubtitleBand.Execute(Data: TObject);
var
  i: Integer;
  Accumulators: TAccumulatorList;
  Acc: TAccumulator;
  acIdx: Integer;
begin
  if (Assigned(OnBeforeRender)) then
    OnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TAccumulatorList) then begin
    Accumulators := TAccumulatorList(Data);
    for i := 0 to Accumulators.Count - 1 do begin
      Acc := TAccumulator(Accumulators[i]);
      if (Acc.IsMarkedForRendering) then begin
        fOwner.GrowRowCount(1);
        fOwner.Cells[0, fOwner.RowCount - 1] := FormatDiplayText(Acc.DisplayName);
        for acIdx := 0 to Length(Acc.SubtotalColumns) - 1 do begin
          fOwner.Cells[Acc.SubtotalColumns[acIdx], fOwner.RowCount - 1] := VarToStr(Acc.fValues[acIdx]);
          Acc.fValues[acIdx] := Null;
        end;
        Acc.UnmarkForRendering;
      end;                         
    end;
  end;

  if (Assigned(OnAfterRender)) then
    OnAfterRender(Data);
end;

{ TAccumulator }

procedure TAccumulator.AddValue(ColID: Byte; Value: Variant);
begin
  if (VarIsNumber(fValues[ColID])) then
    fValues[ColID] := fValues[ColID] + StrToFloatDef(Value, 0)
  else begin
    if (fValues[ColID] = Null) then
      fValues[ColID] := Value
    else  
      fValues[ColID] := fValues[ColID] + Value;
  end;
end;

function TAccumulator.IsInitialized: Boolean;
begin
  Result := fIsInitialized;
end;

function TAccumulator.IsMarkedForRendering: Boolean;
begin
  Result := fIsMarkedForRendering;
end;

procedure TAccumulator.MarkForRendering;
begin
  fIsMarkedForRendering := True;
end;

procedure TAccumulator.Initialized;
begin
  fIsInitialized := True;
end;

procedure TAccumulator.SetDriverColumns(const Value: TBytes);
begin
  SetLength(fDriverColumns, 0);
  fDriverColumns := Value;
end;

procedure TAccumulator.SetSubtotalColumns(const Value: TBytes);
begin
  SetLength(fSubtotalColumns, 0);
  fSubtotalColumns := Value;
  SetLength(fValues, Length(Value));
end;

procedure TAccumulator.UnmarkForRendering;
begin
  fIsMarkedForRendering := False;
end;

end.

