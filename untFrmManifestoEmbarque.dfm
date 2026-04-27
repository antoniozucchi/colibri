object FrmManifestoEmbarque: TFrmManifestoEmbarque
  Left = 0
  Top = 0
  Caption = 'Manifesto de Embarque de Passageiros'
  ClientHeight = 441
  ClientWidth = 952
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  FormStyle = fsMDIChild
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 15
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 952
    Height = 25
    Align = alTop
    Caption = 'Manifesto de Embarque de Passageiros'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 950
  end
  object StatusBarManifestoEmbarque: TStatusBar
    Left = 0
    Top = 422
    Width = 952
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
    ExplicitTop = 414
    ExplicitWidth = 950
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 108
    Width = 952
    Height = 29
    ButtonHeight = 23
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 3
    ExplicitWidth = 950
    object DBNavigatorManifestoEmbarque: TDBNavigator
      Left = 0
      Top = 0
      Width = 138
      Height = 23
      DataSource = FrmDataModule.DataSourceManifestoEmbarque
      VisibleButtons = [nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh]
      Hints.Strings = (
        'Primeiro'
        'Anterior'
        'Pr'#243'ximo'
        #218'ltimo'
        'Inserir'
        'Excluir'
        'Editar'
        'Gravar'
        'Cancelar'
        'Atualizar'
        'Aplicar atualiza'#231#245'es'
        'Calcelar atualiza'#231#245'es')
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
    object btnClearFiltro: TToolButton
      Left = 138
      Top = 0
      Hint = 'Limpar filtro'
      Caption = 'btnClearFiltro'
      ImageIndex = 225
      ParentShowHint = False
      ShowHint = True
    end
    object btnLayout: TToolButton
      Left = 161
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'btnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 184
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Caption = 'Exportar'
      ImageIndex = 54
      ParentShowHint = False
      ShowHint = True
    end
    object ToolButton1: TToolButton
      Left = 207
      Top = 0
      Action = actExcelImportar
    end
    object BitBtn1: TBitBtn
      Left = 230
      Top = 0
      Width = 90
      Height = 23
      Action = actCriarProgramacao
      Align = alLeft
      Caption = 'Programar'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
    end
    object BitBtn11: TBitBtn
      Left = 320
      Top = 0
      Width = 105
      Height = 23
      Action = actAtualizarExecutante
      Align = alLeft
      Caption = 'Executante'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
    end
    object BitBtnExcluir: TBitBtn
      Left = 425
      Top = 0
      Width = 70
      Height = 23
      Action = actExcluirFiltrados
      Align = alLeft
      Caption = 'Excluir'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
    end
  end
  object DBGridManifestoEmbarque: TFilterDBGrid
    Left = 0
    Top = 137
    Width = 952
    Height = 285
    Align = alClient
    Color = clWhite
    DataSource = FrmDataModule.DataSourceManifestoEmbarque
    TabOrder = 4
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = 'Segoe UI'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridManifestoEmbarqueDrawColumnCell
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    EnableZebra = False
    ProgressBar = FrmPrincipal.ProgressBarPrincipal
    LayoutButton = btnLayout
    ExcelButton = btnExcel
    Columns = <
      item
        Expanded = False
        FieldName = 'TipoEmbarque'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo Embarque'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'EmbarqueDesembarque'
        Title.Alignment = taCenter
        Title.Caption = 'Embarque/Desembarque'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataEmbarque'
        Title.Alignment = taCenter
        Title.Caption = 'Data Embarque'
        Width = 95
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NomeExecutante'
        Title.Alignment = taCenter
        Title.Caption = 'Nome'
        Width = 293
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CodigoSAP'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Origem'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Destino'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CentroCusto'
        Title.Alignment = taCenter
        Title.Caption = 'Centro de Custo'
        Width = 124
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DiagramaRede'
        Title.Alignment = taCenter
        Title.Caption = 'Diagrama de Rede'
        Width = 123
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'OperRede'
        Title.Alignment = taCenter
        Title.Caption = 'Oper. Rede'
        Width = 70
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'ElementoPEP'
        Title.Alignment = taCenter
        Title.Caption = 'Elemento PEP'
        Width = 123
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CPF'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'OutroDocumento'
        Title.Alignment = taCenter
        Title.Caption = 'Passaporte'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'txtTipoEtapaServico'
        Title.Alignment = taCenter
        Title.Caption = 'Tipo Etapa Servi'#231'o'
        Width = 184
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'txtFuncao'
        Title.Alignment = taCenter
        Title.Caption = 'Empresa'
        Width = 144
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Servico'
        Title.Alignment = taCenter
        Title.Caption = 'Servi'#231'o'
        Width = 251
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'txtFuncao'
        Title.Alignment = taCenter
        Title.Caption = 'Fun'#231#227'o'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'HorarioSaida'
        Title.Alignment = taCenter
        Title.Caption = 'Hor'#225'rio Sa'#237'da'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TransporteTerrestre'
        Title.Alignment = taCenter
        Title.Caption = 'Trans. Terrestre'
        Visible = True
      end>
  end
  object Panel3: TPanel
    Left = 0
    Top = 25
    Width = 952
    Height = 25
    Align = alTop
    TabOrder = 1
    ExplicitWidth = 950
    object dataInicio: TDateTimePicker
      Left = 71
      Top = 1
      Width = 95
      Height = 23
      Hint = 'Data In'#237'cio para filtro'
      Align = alLeft
      Date = 42599.000000000000000000
      Time = 0.462823229157947900
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
      OnCloseUp = dataInicioCloseUp
      OnKeyPress = dataInicioKeyPress
    end
    object Panel4: TPanel
      Left = 166
      Top = 1
      Width = 70
      Height = 23
      Align = alLeft
      Alignment = taLeftJustify
      Caption = '   Data Fim:'
      TabOrder = 1
    end
    object dataFim: TDateTimePicker
      Left = 236
      Top = 1
      Width = 95
      Height = 23
      Hint = 'Data Fim para filtro'
      Align = alLeft
      Date = 42599.000000000000000000
      Time = 0.462823229157947900
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
      OnCloseUp = dataFimCloseUp
      OnKeyPress = dataFimKeyPress
    end
    object Panel8: TPanel
      Left = 1
      Top = 1
      Width = 70
      Height = 23
      Align = alLeft
      Alignment = taLeftJustify
      Caption = '   Data In'#237'cio:'
      TabOrder = 3
    end
  end
  object PanelFiltroRapido: TPanel
    Left = 0
    Top = 50
    Width = 952
    Height = 58
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    ExplicitWidth = 950
    object GroupBoxTipoEmbarque: TGroupBox
      Left = 0
      Top = 0
      Width = 430
      Height = 58
      Align = alLeft
      Caption = 'Filtro R'#225'pido - Tipo Embarque'
      TabOrder = 0
      object clbTipoEmbarque: TCheckListBox
        Left = 2
        Top = 17
        Width = 426
        Height = 39
        Align = alClient
        BorderStyle = bsNone
        Columns = 2
        ItemHeight = 17
        Items.Strings = (
          'FIXO'
          'SOV'
          'BATE-VOLTA'
          'A'#201'REO')
        TabOrder = 0
        OnClick = clbFiltroRapidoClickCheck
      end
    end
    object GroupBoxSentido: TGroupBox
      Left = 430
      Top = 0
      Width = 250
      Height = 58
      Align = alLeft
      Caption = 'Filtro R'#225'pido - Movimento'
      TabOrder = 1
      object clbEmbarqueDesembarque: TCheckListBox
        Left = 2
        Top = 17
        Width = 246
        Height = 39
        Align = alClient
        BorderStyle = bsNone
        Columns = 2
        ItemHeight = 17
        Items.Strings = (
          'Embarque'
          'Desembarque')
        TabOrder = 0
        OnClick = clbFiltroRapidoClickCheck
      end
    end
    object PanelFiltroRapidoAcoes: TPanel
      Left = 680
      Top = 0
      Width = 272
      Height = 58
      Align = alClient
      BevelOuter = bvNone
      TabOrder = 2
      ExplicitWidth = 270
      object btnRestaurarFiltrosRapidos: TButton
        Left = 12
        Top = 16
        Width = 145
        Height = 25
        Caption = 'Restaurar R'#225'pidos'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
        OnClick = btnRestaurarFiltrosRapidosClick
      end
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 360
    Top = 192
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
    object actExcelImportar: TAction
      Caption = 'Importar'
      Hint = 'Importar dados do Excel'
      ImageIndex = 241
      OnExecute = actExcelImportarExecute
    end
    object actAtualizarExecutante: TAction
      Caption = 'Executante'
      ImageIndex = 32
      OnExecute = actAtualizarExecutanteExecute
    end
    object actCriarProgramacao: TAction
      Caption = 'Programar'
      ImageIndex = 0
      OnExecute = actCriarProgramacaoExecute
    end
    object actExcluirFiltrados: TAction
      Caption = 'Excluir'
      Hint = 'Excluir todos os registros filtrados do manifesto'
      ImageIndex = 324
      OnExecute = actExcluirFiltradosExecute
    end
  end
end

