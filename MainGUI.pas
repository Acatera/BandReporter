unit MainGUI;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DB, ADODB, Acatera.Report, ComCtrls, Gauges, ExtCtrls;

type
  TFMainGUI = class(TForm)
    btnRender: TButton;
    Button: TButton;
    qry1: TADOQuery;
    con1: TADOConnection;
    stat1: TStatusBar;
    Gauge1: TGauge;
    GridPanel1: TGridPanel;
    pnl1: TPanel;
    edt1: TEdit;
    lbl1: TLabel;
    procedure btnRenderClick(Sender: TObject);
    procedure ButtonClick(Sender: TObject);
  private
    procedure UpdateGauge(WorkDone, WorkTotal: integer);
  end;

var
  FMainGUI: TFMainGUI;

implementation

uses
  XLSfile, Acatera.Utils;

{$R *.dfm}

procedure TFMainGUI.btnRenderClick(Sender: TObject);
var
  i: Integer;
  j: Integer;
  Col: TCellProperty;
  DetailBand: TDetailBand;
  TitleBand: TTitleBand;
  SubtotalBand: TSubtitleBand;
  Report: TReport;
  LastTickCount: cardinal;
  QueryTicks: Cardinal;
  ReportExecuteTicks: Cardinal;
  CurrentTicks: Cardinal;
begin
  LastTickCount := GetTickCount;
  Report := TReport.Create(9, 0);
  try
    qry1.Active := True;
    qry1.First;
    QueryTicks := GetTickCount;
    
    Report.DataSet := qry1;
    
    TitleBand := TTitleBand.Create(Report);
    Report.Bands.Add(TitleBand);
    
    DetailBand := TDetailBand.Create(Report);
    DetailBand.Accumulators.Add(TAccumulator.Create);
    with (TAccumulator(DetailBand.Accumulators[0])) do begin
      DisplayName := 'Total doc. %lastValues[3]% / %lastValues[4]%';
      DriverColumns := TBytes.Create(0, 1, 3, 4);
      SubtotalColumns := TBytes.Create(7, 8);
      Operation := acAddition;
    end;
    
    DetailBand.Accumulators.Add(TAccumulator.Create);
    with (TAccumulator(DetailBand.Accumulators[1])) do begin
      DisplayName := 'Total Tip Document';
      Operation := acAddition;
      DriverColumns := TBytes.Create(0, 1);
      SubtotalColumns := TBytes.Create(7, 8);
    end;

    DetailBand.Accumulators.Add(TAccumulator.Create);
    with (TAccumulator(DetailBand.Accumulators[2])) do begin
      DisplayName := 'Total %columnNames[1]%';
      Operation := acAddition;
      DriverColumns := TBytes.Create(0);
      SubtotalColumns := TBytes.Create(7, 8);
    end;
    
    DetailBand.OnBeforeRender := TitleBand.Execute;
    DetailBand.OnWorkProgress := UpdateGauge;

    Report.Bands.Add(DetailBand);
    
    SubtotalBand := TSubtitleBand.Create(Report);
    DetailBand.OnSubtotalColChange := SubtotalBand.Execute;
    Report.Bands.Add(SubtotalBand);
    
    Report.Cols[0].Size := 20;
    Report.Cols[1].Size := 1;
    Report.Cols[2].Size := 20;
    Report.Cols[3].Align := alRight;
    Report.Cols[3].Size := 10;
    Report.Cols[4].Size := 8;
    Report.Cols[5].Size := 20;
    
    Report.Cols[6].Align := alRight;
    Report.Cols[7].Size := 10;
    Report.Cols[6].DataType := dtNumber;
    Report.Cols[6].Formatting := '%.3f';              

    Report.Cols[7].Align := alRight;
    Report.Cols[7].Size := 14;
    Report.Cols[7].DataType := dtNumber;
    Report.Cols[7].Formatting := '%.2f';
    
    Report.Cols[8].Align := alRight;
    Report.Cols[8].Size := 14;
    Report.Cols[8].DataType := dtNumber;
    Report.Cols[8].Formatting := '%.2f';
    
    Report.Execute;
    ReportExecuteTicks := GetTickCount;
    
    Report.RenderToFile(edt1.Text);
  finally
    Report.Free;
  end;
  CurrentTicks := GetTickCount;
  stat1.SimpleText := Format('Query: %d; Execution: %d; Total: %d', [QueryTicks - LastTickCount, ReportExecuteTicks - QueryTicks, CurrentTicks - LastTickCount]);
end;

procedure TFMainGUI.ButtonClick(Sender: TObject);
var
  xls: TXLSfile;
  SetAtribut:TSetOfAtribut;
  LastTick: Cardinal;
  i, j: integer;
  str: string;
begin
  LastTick := GetTickCount;
  for i := 0 to 1000000 do begin
    Str := CenterStr2('a1lex', 10);
  end;

//  xls := TXLSfile.Create(Self);
//  try
//     SetAtribut:=[];
//     xls.FileName := 'D:\Test.xls';
//    
//    for i := 0 to 3000 do
//      for j := 0 to 40 do   
//        xls.AddDoubleCell(1, 1, SetAtribut, 12341234.343);  
//      
//    xls.write;
//  finally
//    xls.Free;
//  end;  
  ShowMessage(IntToStr(GetTickCount - LastTick));
end;

procedure TFMainGUI.UpdateGauge(WorkDone, WorkTotal: integer);
var
  WorkPerPercent: Integer;
begin
  WorkPerPercent := Round(WorkTotal / 100);
  if (Gauge1.MaxValue <> WorkTotal) then
    Gauge1.MaxValue := WorkTotal;
  Gauge1.Progress := WorkDone;
  if (WorkDone mod WorkPerPercent = 0) then
    Application.ProcessMessages;
end;

end.
