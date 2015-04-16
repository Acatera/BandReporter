unit Acatera.Report;

interface

uses
  Variants, Contnrs, DB;

type
  TDataType = (dtNumber, dtString);
  TAlign = (alLeft, alRight, alCenter);

  // TCell = record
  //   Value: string;
  // end;

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

  TReport = class
  private
    fActualRowCount: Integer;
    fBands: TBands;
    fCells: TCells;
    fCols: TCols;
    fRows: TRows;
    fRowCount: integer;
    fColCount: integer;
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
  public
    fDataSet: TDataSet;
    property Bands: TBands read fBands;
    property Cells[Col, Row: Integer]: string read GetCell write SetCell;
    property ColCount: integer read fColCount;
    property RowCount: integer read fRowCount;
    property Cols[Index: SmallInt]: PCellProperty read GetCol write SetCol; default;
    property Rows[Index: SmallInt]: PCellProperty read GetRow;
    constructor Create(Cols, Rows: integer);
    procedure RenderToFile(FileName: string);
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
    procedure GrowRowCount(Value: integer);
  end;

  TBandEvent = procedure(Event: TObject) of object;
  
  TAccumulatorOp = (acNone, acAddition, acSubtraction, acMultiply, acDivide, acPower, acRoot);
  TAccumulator = class
    fDrivenByCol: SmallInt;
    fLastCodeValue: Variant;
    fIsInitialized: boolean;
    fOp: TAccumulatorOp;
    fValue: Variant;
    fColID: Smallint;
    procedure AddValue(Value: Variant);
  end;
  TAccumulatorList = class(TObjectList);
  
  TOccurence = (atNewPage, atChangeID, atEachRow, atPageEnd);
  TBand = class
    fAccumulators: TAccumulatorList;
    fOwner: TReport;
    fOccurence: TOccurence;
    fDataSet: TDataSet;
    fOnBeforeRender: TBandEvent;
    fOnAfterRender: TBandEvent;
    fOnSubtotalColChange: TBandEvent; 
    procedure Render(Data: TObject); virtual; abstract;
  public
    procedure AfterConstruction; override;
    procedure BeforeDestruction; override;
  end;

  TDetailBand = class(TBand)
    constructor Create(Owner: TReport);
    procedure Render(Data: TObject); override;
  end;
  TTitleBand = class(TBand)
    constructor Create(Owner: TReport);
    procedure Render(Data: TObject); override;
  end;

  TSubtitleBand = class(TBand)
    constructor Create(Owner: TReport);
    procedure Render(Data: TObject); override;
  end;
  
  (*
    Phase 1:
    [x] Render as text file (for now) - temporary;
    [x] Implement default cell values (width, heigth, align, data type);

    Phase 2:
    [x] Implement bands (title, detail, (sub)total);
    
    Phase 3:
    [x] Fix GrowRowCountBy(xyz);
    [x] Implement a way to render the bands in a way that makes them "connected" to eachother (OnBeforeRender?)
    [x] Each band can hold it's own variables, for sum purposes (the AccumulatorList);

    [ ] Each Accumulator should have it's own subtotal band; Think about this... sounds restricting
    
    Phase 4:
    [ ] Implement persistency (dfm file?);

    Phase 5:
    [ ] Cell merging

    Phase 6:
    [ ] Excel export

    Phase 6:
    [ ] Cell font, borders, color;
  *)

implementation

uses
  SysUtils, StrUtils, Acatera.Utils, Windows, Classes;

{ TReport }

procedure TReport.AfterConstruction;
begin
  fBands := TBands.Create(True);
end;

procedure TReport.BeforeDestruction;
begin
  fBands.Free;
end;

constructor TReport.Create(Cols, Rows: integer);
begin
  SetCapacity(Cols, Rows);
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
      for Y := 0 to fBands.Count - 1 do begin
        if (TBand(fBands[Y]).fOccurence = atEachRow) then begin
          TBand(fBands[Y]).Render(fDataSet);
        end;
      end;
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
    if (TBand(fBands[i]).fOccurence = atEachRow) then
      if ((i > 0) and (TBand(fBands[i - 1]).fOccurence = atNewPage)) then
        TBand(fBands[i]).fOnBeforeRender := TBand(fBands[i - 1]).Render;
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

constructor TDetailBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  fOccurence := atEachRow;
end;

procedure TDetailBand.Render(Data: TObject);
var
  i, j: Integer;
  ShouldRenderSubtotal: Boolean;
begin
  if (Assigned(fOnBeforeRender)) then
    fOnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TDataSet) then begin
    fDataSet := TDataSet(Data);
    while (not fDataSet.Eof) do begin
      for i := 0 to fAccumulators.Count - 1 do begin
        if (not TAccumulator(fAccumulators[i]).fIsInitialized) then begin
          TAccumulator(fAccumulators[i]).fLastCodeValue := fDataSet.Fields[TAccumulator(fAccumulators[i]).fDrivenByCol].AsString;
          TAccumulator(fAccumulators[i]).fIsInitialized := True;
        end;
        TAccumulator(fAccumulators[i]).AddValue(fDataSet.Fields[TAccumulator(fAccumulators[i]).fColID].AsVariant);
      end;
    
      fOwner.GrowRowCount(1);
      for j := 0 to fDataSet.FieldCount - 1 do begin
        fOwner.Cells[j, fOwner.RowCount - 1] := fDataSet.Fields[j].AsString;
      end;
      
      fDataSet.Next;
      
      for i := 0 to fAccumulators.Count - 1 do begin
        if (fDataSet.Fields[TAccumulator(fAccumulators[i]).fDrivenByCol].AsString <> TAccumulator(fAccumulators[i]).fLastCodeValue) then begin
          ShouldRenderSubtotal := True;
          TAccumulator(fAccumulators[i]).fLastCodeValue := fDataSet.Fields[TAccumulator(fAccumulators[i]).fDrivenByCol].AsString;
        end;
      end;
      if (Assigned(fOnSubtotalColChange) and (ShouldRenderSubtotal)) then begin
        fOnSubtotalColChange(fAccumulators);
        ShouldRenderSubtotal := False;
      end;
    end;
    fDataSet := nil;
  end;
  
  if (Assigned(fOnSubtotalColChange)) then begin
    fOnSubtotalColChange(fAccumulators);
  end;
      
  if (Assigned(fOnAfterRender)) then
    fOnAfterRender(Data);
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

{ TTitleBand }

constructor TTitleBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  fOccurence := atNewPage;
end;

procedure TTitleBand.Render(Data: TObject);
var
  i: Integer;
begin
  if (Assigned(fOnBeforeRender)) then
    fOnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TDataSet) then begin
    fDataSet := TDataSet(Data);
    fOwner.GrowRowCount(1);
    for i := 0 to fDataSet.FieldCount - 1 do begin
      fOwner.Cells[i, fOwner.RowCount - 1] := fDataSet.Fields[i].DisplayName;
    end;
  end;

  if (Assigned(fOnAfterRender)) then
    fOnAfterRender(Data);
end;

{ TSubtitleBand }

constructor TSubtitleBand.Create(Owner: TReport);
begin
  fOwner := Owner;
  fOccurence := atChangeID;
end;

procedure TSubtitleBand.Render(Data: TObject);
var
  i: Integer;
  Accumulators: TAccumulatorList;
begin
  if (Assigned(fOnBeforeRender)) then
    fOnBeforeRender(Data);
    
  if (Data <> nil) and (Data is TAccumulatorList) then begin
    Accumulators := TAccumulatorList(Data);
    fOwner.GrowRowCount(1);
    for i := 0 to Accumulators.Count - 1 do begin
      fOwner.Cells[0, fOwner.RowCount - 1] := '[Subtotal]';
      fOwner.Cells[TAccumulator(Accumulators[I]).fColID, fOwner.RowCount - 1] := VarToStr(TAccumulator(Accumulators[I]).fValue);
      TAccumulator(Accumulators[I]).fValue := Null;
    end;
  end;

  if (Assigned(fOnAfterRender)) then
    fOnAfterRender(Data);
end;

{ TAccumulator }

procedure TAccumulator.AddValue(Value: Variant);
begin
  if (VarIsNumber(fValue)) then
    fValue := fValue + StrToFloatDef(Value, 0)
  else begin
    if (fValue = Null) then
      fValue := Value
    else  
      fValue := fValue + Value;
  end;
end;

end.

