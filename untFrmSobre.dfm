object FrmSobre: TFrmSobre
  Left = 0
  Top = 0
  Caption = 'Sobre'
  ClientHeight = 741
  ClientWidth = 934
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  OldCreateOrder = False
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object Panel3: TPanel
    Left = 0
    Top = 0
    Width = 934
    Height = 25
    Align = alTop
    Caption = 'Sobre o programa'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    object DBNavigatorVersao: TDBNavigator
      Left = 1
      Top = 1
      Width = 96
      Height = 23
      DataSource = FrmDataModule.DataSourceColibri
      VisibleButtons = [nbEdit, nbPost, nbCancel, nbRefresh]
      Align = alLeft
      TabOrder = 0
    end
  end
  object MemoInformacoes: TMemo
    Left = 0
    Top = 81
    Width = 934
    Height = 660
    Align = alClient
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clBlack
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 1
    ExplicitTop = 105
    ExplicitHeight = 636
  end
  object PanelEntrada: TPanel
    Left = 0
    Top = 25
    Width = 934
    Height = 56
    Align = alTop
    TabOrder = 2
    object PanelSuperior: TPanel
      Left = 1
      Top = 1
      Width = 932
      Height = 54
      Align = alClient
      TabOrder = 0
      ExplicitTop = -4
      ExplicitHeight = 78
      object PanelVersaoDB: TPanel
        Left = 1
        Top = 1
        Width = 930
        Height = 25
        Align = alTop
        TabOrder = 0
        object Panel6: TPanel
          Left = 1
          Top = 1
          Width = 180
          Height = 23
          Align = alLeft
          Alignment = taLeftJustify
          Caption = ' Vers'#227'o do banco de dados:'
          TabOrder = 0
        end
        object Panel7: TPanel
          Left = 181
          Top = 1
          Width = 748
          Height = 23
          Align = alClient
          TabOrder = 1
          object DBEditVersao: TDBEdit
            Left = 1
            Top = 1
            Width = 746
            Height = 24
            Hint = 'Alterar a vers'#227'o do programa'
            Align = alTop
            DataField = 'versao'
            DataSource = FrmDataModule.DataSourceColibri
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
          end
        end
      end
      object PanelInstalador: TPanel
        Left = 1
        Top = 26
        Width = 930
        Height = 25
        Align = alTop
        TabOrder = 1
        ExplicitTop = 51
        object Panel2: TPanel
          Left = 1
          Top = 1
          Width = 180
          Height = 23
          Align = alLeft
          Alignment = taLeftJustify
          Caption = ' Endere'#231'o do instalador:'
          TabOrder = 0
          object btnInstalador: TSpeedButton
            Left = 156
            Top = 1
            Width = 23
            Height = 21
            Align = alRight
            Glyph.Data = {
              36040000424D3604000000000000360000002800000010000000100000000100
              2000000000000004000000000000000000000000000000000000FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00000000000000
              0000000000000000000000000000000000000000000000000000000000000000
              000000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00000000000000
              0000008080000080800000808000008080000080800000808000008080000080
              80000080800000000000FF00FF00FF00FF00FF00FF00FF00FF000000000000FF
              FF00000000000080800000808000008080000080800000808000008080000080
              8000008080000080800000000000FF00FF00FF00FF00FF00FF0000000000FFFF
              FF0000FFFF000000000000808000008080000080800000808000008080000080
              800000808000008080000080800000000000FF00FF00FF00FF000000000000FF
              FF00FFFFFF0000FFFF0000000000008080000080800000808000008080000080
              80000080800000808000008080000080800000000000FF00FF0000000000FFFF
              FF0000FFFF00FFFFFF0000FFFF00000000000000000000000000000000000000
              00000000000000000000000000000000000000000000000000000000000000FF
              FF00FFFFFF0000FFFF00FFFFFF0000FFFF00FFFFFF0000FFFF00FFFFFF0000FF
              FF0000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF0000000000FFFF
              FF0000FFFF00FFFFFF0000FFFF00FFFFFF0000FFFF00FFFFFF0000FFFF00FFFF
              FF0000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF000000000000FF
              FF00FFFFFF0000FFFF0000000000000000000000000000000000000000000000
              000000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF000000
              00000000000000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00000000000000000000000000FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF000000000000000000FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF0000000000FF00
              FF00FF00FF00FF00FF0000000000FF00FF0000000000FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF000000
              00000000000000000000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
            ParentShowHint = False
            ShowHint = True
            OnClick = btnInstaladorClick
            ExplicitLeft = 152
            ExplicitTop = 3
          end
        end
        object Panel4: TPanel
          Left = 181
          Top = 1
          Width = 748
          Height = 23
          Align = alClient
          TabOrder = 1
          object DBEditInstalador: TDBEdit
            Left = 1
            Top = 1
            Width = 746
            Height = 24
            Hint = 
              'Alterar o endere'#231'o do instalador, caso necess'#225'rio atualiza'#231#227'o au' +
              'tom'#225'tica'
            Align = alTop
            DataField = 'enderecoInstalador'
            DataSource = FrmDataModule.DataSourceColibri
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
          end
        end
      end
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 312
    Top = 216
  end
end
