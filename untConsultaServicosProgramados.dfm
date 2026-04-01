object FrmConsultaServicosProgramados: TFrmConsultaServicosProgramados
  Left = 0
  Top = 0
  Caption = 'Servi'#231'os Programados por Per'#237'odo'
  ClientHeight = 414
  ClientWidth = 1356
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 13
  object DBGridServicosProgramados: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 1356
    Height = 341
    Align = alClient
    DataSource = FrmDataModule.DataSourceConsultaServicosProgramados
    ReadOnly = True
    TabOrder = 0
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridServicosProgramadosDrawColumnCell
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    LayoutGrid = ColunasLayout
    EnableZebra = False
    LayoutButton = btnLayout
    ExcelButton = btnExcel
    Columns = <
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataProgramacao'
        Title.Alignment = taCenter
        Title.Caption = 'Data da Programa'#231#227'o'
        Width = 114
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'txtDestino'
        Title.Alignment = taCenter
        Title.Caption = 'Destino'
        Width = 90
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'txtTipoEtapaServico'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo de Etapa de Servico'
        Width = 197
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CampoSelecao'
        Title.Alignment = taCenter
        Title.Caption = 'Servi'#231'o/Ano-Etapa'
        Width = 100
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TextoBreveOP'
        Title.Alignment = taCenter
        Title.Caption = 'Texto Breve da Opera'#231#227'o'
        Width = 250
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TextoBreveOM'
        Title.Alignment = taCenter
        Title.Caption = 'Texto Breve da Ordem'
        Width = 250
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'OrdemManutencao'
        Title.Alignment = taCenter
        Title.Caption = 'Ordem de Manuten'#231#227'o'
        Width = 119
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Operacao'
        Title.Alignment = taCenter
        Title.Caption = 'Opera'#231#227'o'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CentroTrabalhoOP'
        Title.Alignment = taCenter
        Title.Caption = 'Centro Trabalho [OP]'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CriadoPor'
        Title.Alignment = taCenter
        Title.Caption = 'Criado Por'
        Width = 76
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataCriacao'
        Title.Alignment = taCenter
        Title.Caption = 'Data [Cria'#231#227'o]'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'ComputadorCriacao'
        Title.Alignment = taCenter
        Title.Caption = 'Computador [Cria'#231#227'o]'
        Visible = True
      end>
  end
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 1356
    Height = 25
    Align = alTop
    Caption = 'Servi'#231'os Programados por Per'#237'odo'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 1354
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 395
    Width = 1356
    Height = 19
    Panels = <
      item
        Width = 50
      end
      item
        Width = 50
      end
      item
        Width = 50
      end
      item
        Width = 50
      end
      item
        Width = 50
      end>
    ExplicitTop = 387
    ExplicitWidth = 1354
  end
  object ColunasLayout: TStringGrid
    Left = 50
    Top = 153
    Width = 113
    Height = 94
    ColCount = 7
    DefaultRowHeight = 21
    RowCount = 12
    TabOrder = 3
    Visible = False
    RowHeights = (
      21
      21
      21
      21
      21
      21
      21
      21
      21
      21
      21
      21)
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 1356
    Height = 29
    ButtonHeight = 24
    Caption = 'ToolBar1'
    Images = FrmPrincipal.ImageList1
    TabOrder = 4
    ExplicitWidth = 1354
    object Panel8: TPanel
      Left = 0
      Top = 0
      Width = 70
      Height = 24
      Align = alLeft
      Alignment = taLeftJustify
      Caption = '   Data In'#237'cio:'
      TabOrder = 0
    end
    object dataInicio: TDateTimePicker
      Left = 70
      Top = 0
      Width = 95
      Height = 24
      Hint = 'Data In'#237'cio para filtro'
      Align = alLeft
      Date = 42599.000000000000000000
      Time = 0.462823229157947900
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnCloseUp = actProcurarExecute
    end
    object Panel24: TPanel
      Left = 165
      Top = 0
      Width = 70
      Height = 24
      Align = alLeft
      Alignment = taLeftJustify
      Caption = '   Data Fim:'
      TabOrder = 2
    end
    object dataFim: TDateTimePicker
      Left = 235
      Top = 0
      Width = 95
      Height = 24
      Hint = 'Data Fim para filtro'
      Align = alLeft
      Date = 42599.000000000000000000
      Time = 0.462823229157947900
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
      OnCloseUp = actProcurarExecute
    end
    object btnClearFiltro: TToolButton
      Left = 330
      Top = 0
      Hint = 'Limpar filtro'
      Caption = 'btnClearFiltro'
      ImageIndex = 225
      ParentShowHint = False
      ShowHint = True
    end
    object btnLayout: TToolButton
      Left = 353
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'btnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 376
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Caption = 'Exportar'
      ImageIndex = 54
      ParentShowHint = False
      ShowHint = True
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 288
    Top = 304
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      ShortCut = 116
      OnExecute = actProcurarExecute
    end
  end
end
