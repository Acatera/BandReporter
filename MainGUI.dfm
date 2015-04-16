object FMainGUI: TFMainGUI
  Left = 0
  Top = 0
  Caption = 'FMainGUI'
  ClientHeight = 224
  ClientWidth = 527
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object grd1: TStringGrid
    Left = 0
    Top = 0
    Width = 527
    Height = 169
    Align = alTop
    DefaultColWidth = 70
    DefaultRowHeight = 20
    FixedCols = 0
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine]
    TabOrder = 1
  end
  object btn1: TButton
    Left = 8
    Top = 175
    Width = 75
    Height = 25
    Caption = 'btn1'
    TabOrder = 0
    OnClick = btn1Click
  end
  object Button: TButton
    Left = 89
    Top = 175
    Width = 75
    Height = 25
    Caption = 'Button'
    TabOrder = 2
    OnClick = ButtonClick
  end
  object stat1: TStatusBar
    Left = 0
    Top = 205
    Width = 527
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object qry1: TADOQuery
    Connection = con1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select Gestiune, Tip_Doc, Denfur, Numar, Data, Denumire, Cantita' +
        'te, Pret_achiz, Pret_Aman from del.miv0215 left join del.fur usi' +
        'ng(CodFur)'
      
        '#SELECT CodFC, ContFC, DenFur, NrDFC, DataFC, SumaDB, SumaCR FRO' +
        'M del.fres0315 left join del.fur on codfc = codfur where contfc ' +
        '= 401 and denfur is not null ORDER BY DenFur')
    Left = 408
    Top = 88
  end
  object con1: TADOConnection
    Connected = True
    ConnectionString = 
      'Driver={MySQL ODBC 3.51 Driver};Server=localhost;Port=3306;Optio' +
      'n=%s;Database=mysql;User=root'#39';'
    LoginPrompt = False
    Left = 376
    Top = 32
  end
end
