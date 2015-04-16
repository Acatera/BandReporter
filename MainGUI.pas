unit MainGUI;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DB, ADODB, Acatera.Report, ComCtrls;

type
  TFMainGUI = class(TForm)
    grd1: TStringGrid;
    btn1: TButton;
    Button: TButton;
    qry1: TADOQuery;
    con1: TADOConnection;
    stat1: TStatusBar;
    procedure btn1Click(Sender: TObject);
    procedure Test(Sender: TObject);
    procedure ButtonClick(Sender: TObject);
  private
    procedure RenderReport(Report: TReport);
  end;

var
  FMainGUI: TFMainGUI;

implementation

uses
  XLSfile, Acatera.Utils;

{$R *.dfm}

procedure TFMainGUI.btn1Click(Sender: TObject);
var
  i: Integer;
  j: Integer;
  Col: TCellProperty;
  DetailBand: TDetailBand;
  TitleBand: TTitleBand;
  SubtotalBand: TSubtitleBand;
  Report: TReport;
  LastTickCount: cardinal;
begin
  LastTickCount := GetTickCount;
  Report := TReport.Create(9, 0);
  try
    qry1.Active := True;
    qry1.First;
    
    Report.fDataSet := qry1;
    
    TitleBand := TTitleBand.Create(Report);
    Report.Bands.Add(TitleBand);
    
    DetailBand := TDetailBand.Create(Report);
    DetailBand.fAccumulators.Add(TAccumulator.Create);
    with (TAccumulator(DetailBand.fAccumulators[0])) do begin
      fOp := acAddition;
      fDrivenByCol := 0;
      fColID := 7;
    end;
    DetailBand.fAccumulators.Add(TAccumulator.Create);
    with (TAccumulator(DetailBand.fAccumulators[1])) do begin
      fOp := acAddition;
      fDrivenByCol := 0;
      fColID := 8;
    end;
    DetailBand.fOnBeforeRender := TitleBand.Render;
    Report.Bands.Add(DetailBand);
    
    SubtotalBand := TSubtitleBand.Create(Report);
    DetailBand.fOnSubtotalColChange := SubtotalBand.Render;
    Report.Bands.Add(SubtotalBand);
    
    Report.Cols[0].Size := 10;
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
    
    RenderReport(Report);
    Report.RenderToFile('Test.txt');
  finally
    Report.Free;
  end;
  stat1.SimpleText := IntToStr(GetTickCount - LastTickCount);
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

procedure TFMainGUI.RenderReport(Report: TReport);
var
  i: Integer;
  j: Integer;
begin
  for i := 0 to Report.RowCount - 1 do
    for j := 0 to Report.ColCount - 1 do
      grd1.Cells[j, i + 1] := Report.Cells[j, i];
end;

procedure TFMainGUI.Test(Sender: TObject);
var
  Test: string;
begin
  Test := 'This is a test';
end;
{
var
   row,col:integer;
   SetAtribut:TSetOfAtribut;
begin
   row:=StrToInt(ERow.text);
   col:=StrToInt(ECol.text);
   SetAtribut:=[];

   if CBShaded.Checked then Include(SetAtribut,acShaded);
   if CBBottomBorder.Checked then Include(SetAtribut,acBottomBorder);
   if CBTopBorder.Checked then Include(SetAtribut,acTopBorder);
   if CBLeftBorder.Checked then Include(SetAtribut,acLeftBorder);
   if CBRightBorder.Checked then Include(SetAtribut,acRightBorder);
   if RBLeftAllign.Checked then Include(SetAtribut,acLeft) else
   if RBRightAllign.Checked then Include(SetAtribut,acRight) else
   if RBCenterAllign.Checked then Include(SetAtribut,acCenter) else
   if RBFillAllign.Checked then Include(SetAtribut,acFill);


   //You can use directly like this: AddWordCell(1,1,[acBottomBorder,acTopBorder],200);
   case RGType.ItemIndex  of
      0: XLSFIle1.AddWordCell(col,row,SetAtribut,StrToInt(Evalue.text));
      1: XLSFIle1.AddDoubleCell(col,row,SetAtribut,StrToFloat(Evalue.text));
      2: XLSFIle1.AddStrCell(col,row,SetAtribut,Evalue.text);
   end;
   ERow.text:='';
   ECol.text:='';
   Evalue.text:='';
end;

procedure TForm1.BclearClick(Sender: TObject);
begin
     XLSfile1.clear; 
end;
}

end.
