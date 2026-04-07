object FrmParesObrigatorios: TFrmParesObrigatorios
  Left = 0
  Top = 0
  BorderStyle = bsDialog
  Caption = 'Pares Obrigat'#243'rios de Distribui'#231#227'o'
  ClientHeight = 420
  ClientWidth = 520
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  OldCreateOrder = False
  Position = poMainFormCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 520
    Height = 36
    Align = alTop
    BevelOuter = bvNone
    Caption = 'Pares Obrigat'#243'rios - Distribui'#231#227'o Autom'#225'tica'
    Color = 3355443
    Font.Charset = ANSI_CHARSET
    Font.Color = clWhite
    Font.Height = -13
    Font.Name = 'Arial'
    Font.Style = [fsBold]
    ParentFont = False
    TabOrder = 0
  end
  object LabelInfo: TLabel
    Left = 8
    Top = 44
    Width = 504
    Height = 26
    AutoSize = False
    Caption =
      'Defina os pares de plataformas que devem obrigatoriamente ser at' +
      'endidas pela mesma embarca'#231#227'o.'
    Font.Charset = ANSI_CHARSET
    Font.Color = clGrayText
    Font.Height = -10
    Font.Name = 'Arial'
    Font.Style = []
    ParentFont = False
    WordWrap = True
  end
  object DBGridPares: TDBGrid
    Left = 8
    Top = 76
    Width = 504
    Height = 296
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines,
      dgRowLines, dgTabs, dgAlwaysShowSelection, dgTitleClick,
      dgTitleHotTrack]
    TabOrder = 1
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
    OnCellClick = DBGridParesCellClick
    OnDrawColumnCell = DBGridParesDrawColumnCell
    Columns = <
      item
        Expanded = False
        FieldName = 'PlataformaA'
        Title.Alignment = taCenter
        Title.Caption = 'Plataforma A'
        Width = 180
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'PlataformaB'
        Title.Alignment = taCenter
        Title.Caption = 'Plataforma B'
        Width = 180
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Ativo'
        Title.Alignment = taCenter
        Title.Caption = 'Ativo'
        Width = 60
        Visible = True
      end>
  end
  object PanelBotoes: TPanel
    Left = 0
    Top = 380
    Width = 520
    Height = 40
    Align = alBottom
    BevelOuter = bvNone
    TabOrder = 2
    object BtnAdicionar: TBitBtn
      Left = 8
      Top = 6
      Width = 90
      Height = 28
      Caption = '+ Adicionar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      OnClick = BtnAdicionarClick
    end
    object BtnExcluir: TBitBtn
      Left = 104
      Top = 6
      Width = 80
      Height = 28
      Caption = 'Excluir'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      OnClick = BtnExcluirClick
    end
    object BtnSalvar: TBitBtn
      Left = 200
      Top = 6
      Width = 80
      Height = 28
      Caption = 'Salvar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
      OnClick = BtnSalvarClick
    end
    object BtnCancelar: TBitBtn
      Left = 286
      Top = 6
      Width = 80
      Height = 28
      Caption = 'Cancelar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ParentFont = False
      TabOrder = 3
      OnClick = BtnCancelarClick
    end
    object BtnFechar: TBitBtn
      Left = 424
      Top = 6
      Width = 88
      Height = 28
      Caption = 'Fechar'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Arial'
      Font.Style = []
      ModalResult = mrCancel
      ParentFont = False
      TabOrder = 4
      OnClick = BtnFecharClick
    end
  end
end
