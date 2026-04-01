object FrmTabela: TFrmTabela
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Transbordos'
  ClientHeight = 521
  ClientWidth = 1087
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  FormStyle = fsStayOnTop
  TextHeight = 15
  object RLTabela: TStringGrid
    Left = 0
    Top = 54
    Width = 1087
    Height = 467
    Align = alClient
    RowCount = 2
    Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goRowMoving, goColMoving, goEditing, goFixedRowClick]
    TabOrder = 0
    ExplicitWidth = 1085
    ExplicitHeight = 459
    ColWidths = (
      64
      64
      64
      64
      64)
    RowHeights = (
      24
      24)
  end
  object ToolBar6: TToolBar
    Left = 0
    Top = 25
    Width = 1087
    Height = 29
    Caption = 'ToolBar3'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 1
    ExplicitWidth = 1085
    object BitBtn13: TBitBtn
      Left = 0
      Top = 0
      Width = 70
      Height = 22
      Action = actExcel
      Caption = 'Exportar'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
  end
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 1087
    Height = 25
    Align = alTop
    Caption = 'Tabela'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 2
    ExplicitWidth = 1085
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 304
    Top = 224
    StyleName = 'Platform Default'
    object actExcel: TAction
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelExecute
    end
  end
end
