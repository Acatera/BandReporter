object FMainGUI: TFMainGUI
  Left = 0
  Top = 0
  BorderStyle = bsSingle
  Caption = 'Excel report generator demo'
  ClientHeight = 118
  ClientWidth = 520
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 16
  object Gauge1: TGauge
    AlignWithMargins = True
    Left = 3
    Top = 35
    Width = 514
    Height = 21
    Align = alClient
    Progress = 0
    ExplicitLeft = 8
    ExplicitTop = 0
    ExplicitWidth = 521
    ExplicitHeight = 55
  end
  object stat1: TStatusBar
    Left = 0
    Top = 99
    Width = 520
    Height = 19
    Panels = <>
    SimplePanel = True
  end
  object GridPanel1: TGridPanel
    Left = 0
    Top = 59
    Width = 520
    Height = 40
    Align = alBottom
    ColumnCollection = <
      item
        Value = 50.000000000000000000
      end
      item
        Value = 50.000000000000000000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = btnRender
        Row = 0
      end
      item
        Column = 1
        Control = Button
        Row = 0
      end>
    RowCollection = <
      item
        Value = 100.000000000000000000
      end>
    TabOrder = 1
    object btnRender: TButton
      AlignWithMargins = True
      Left = 4
      Top = 4
      Width = 253
      Height = 32
      Align = alClient
      Caption = 'Generate report'
      TabOrder = 0
      OnClick = btnRenderClick
    end
    object Button: TButton
      AlignWithMargins = True
      Left = 263
      Top = 4
      Width = 253
      Height = 32
      Align = alClient
      Caption = 'Button'
      TabOrder = 1
      OnClick = ButtonClick
    end
  end
  object pnl1: TPanel
    Left = 0
    Top = 0
    Width = 520
    Height = 32
    Align = alTop
    Caption = 'pnl1'
    TabOrder = 2
    object lbl1: TLabel
      AlignWithMargins = True
      Left = 4
      Top = 7
      Width = 147
      Height = 18
      Margins.Top = 6
      Margins.Bottom = 6
      Align = alLeft
      Caption = 'Select exporting location:'
      ExplicitHeight = 16
    end
    object edt1: TEdit
      AlignWithMargins = True
      Left = 157
      Top = 4
      Width = 359
      Height = 24
      Align = alClient
      TabOrder = 0
      Text = 'test.txt'
    end
  end
  object qry1: TADOQuery
    Connection = con1
    CursorType = ctStatic
    Parameters = <>
    SQL.Strings = (
      
        'select Gestiune, Tip_Doc, Denfur, Numar, Data, Denumire, Cantita' +
        'te, Pret_achiz, Pret_Aman from del.miv0215 left join del.fur usi' +
        'ng(CodFur) order by Gestiune, Tip_doc, Numar, Data, Denumire'
      
        '#SELECT CodFC, ContFC, DenFur, NrDFC, DataFC, SumaDB, SumaCR FRO' +
        'M del.fres0315 left join del.fur on codfc = codfur where contfc ' +
        '= 401 and denfur is not null ORDER BY DenFur')
    Left = 136
    Top = 8
  end
  object con1: TADOConnection
    Connected = True
    ConnectionString = 
      'Driver={MySQL ODBC 3.51 Driver};Server=localhost;Port=3306;Optio' +
      'n=%s;Database=mysql;User=root'#39';'
    LoginPrompt = False
    Left = 152
    Top = 8
  end
end
