object FrmConsultaExecutantesProgramados: TFrmConsultaExecutantesProgramados
  Left = 0
  Top = 0
  Caption = 'Executantes Programados por Per'#237'odo'
  ClientHeight = 628
  ClientWidth = 1090
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
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 1090
    Height = 25
    Align = alTop
    Caption = 'Executantes Programados por Per'#237'odo'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 1088
  end
  object Panel3: TPanel
    Left = 0
    Top = 25
    Width = 1090
    Height = 25
    Align = alTop
    TabOrder = 1
    ExplicitWidth = 1088
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
  object PageControl1: TPageControl
    Left = 0
    Top = 50
    Width = 1090
    Height = 578
    ActivePage = TabSheet4
    Align = alClient
    TabOrder = 2
    OnChange = PageControl1Change
    ExplicitWidth = 1088
    ExplicitHeight = 570
    object TabSheet1: TTabSheet
      Caption = 'Executantes Programados'
      object DBGridExecutantesProgramados: TFilterDBGrid
        Left = 0
        Top = 54
        Width = 1082
        Height = 477
        Align = alClient
        DataSource = FrmDataModule.DataSourceConsultaExecutantesProgramados
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGridExecutantesProgramadosCellClick
        OnDrawColumnCell = DBGridExecutantesProgramadosDrawColumnCell
        SelectAllButton = btnSelAll
        SelectClearButton = btnSelClear
        ClearFilterButton = btnClearFiltroExecutante
        SearchAction = actProcurarProgramacaoExecutante
        LayoutGrid = ColunasLayoutExecutanteProgramado
        EnableZebra = False
        FieldSelection = 'booleanRecolhimento'
        ProgressBar = FrmPrincipal.ProgressBarPrincipal
        LayoutButton = btnLayoutExecutante
        ExcelButton = btnExcelExecutante
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'idProgramacaoExecutante'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'ID Executante'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Transbordo'
            Title.Alignment = taCenter
            Title.Caption = 'Transb.'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_GrupoTransbordo'
            Title.Alignment = taCenter
            Title.Caption = 'Grupo Transbordo'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_TransbordoAereo'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Transb. A'#233'reo'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_PrimeiraOrigemTransbordo'
            Title.Alignment = taCenter
            Title.Caption = '1'#186' Origem Transb.'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_PrimeiroDestinoTransbordo'
            Title.Alignment = taCenter
            Title.Caption = '1'#186' Destino Transb.'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_SeqTransbordo'
            Title.Alignment = taCenter
            Title.Caption = 'Seq. Transb.'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Status'
            Title.Alignment = taCenter
            Title.Caption = 'Status'
            Width = 103
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CodigoSAP'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo SAP'
            Width = 86
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'booleanRecolhimento'
            Title.Alignment = taCenter
            Title.Caption = 'Recolhimento'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Origem'
            Title.Alignment = taCenter
            Width = 80
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'txtDestino'
            Title.Alignment = taCenter
            Title.Caption = 'Destino'
            Width = 80
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RecolherPara'
            Title.Alignment = taCenter
            Title.Caption = 'Recolher Para'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataProgramacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Ida'
            Width = 96
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DataVolta'
            Title.Alignment = taCenter
            Title.Caption = 'Data Volta'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_HoraIda'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Ida'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_HoraVolta'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Volta'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Tipo'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo RT'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Modal'
            Title.Alignment = taCenter
            Title.Caption = 'Modal'
            Width = 58
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Classe'
            Title.Alignment = taCenter
            Title.Caption = 'Classe'
            Width = 83
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Mensagem'
            Title.Alignment = taCenter
            Title.Caption = 'Mensagem'
            Width = 120
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'StatusProgramacao'
            Title.Alignment = taCenter
            Title.Caption = 'Status [Programa'#231#227'o]'
            Width = 114
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumExecutantes'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Exec.'
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'txtTipoEtapaServico'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo de Etapa de Servico'
            Width = 156
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'NomeExecutante'
            Title.Alignment = taCenter
            Title.Caption = 'Nome do Executante'
            Width = 168
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumAprovados'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Apro.'
            Width = 50
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumCancelados'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Canc.'
            Width = 50
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Funcao'
            Title.Alignment = taCenter
            Title.Caption = 'Fun'#231#227'o'
            Width = 128
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Empresa'
            Title.Alignment = taCenter
            Width = 177
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'MotivoProgramacao'
            Title.Alignment = taCenter
            Title.Caption = 'Motivo [Programa'#231#227'o]'
            Width = 170
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Documento'
            Title.Alignment = taCenter
            Title.Caption = 'CPF'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'OutroDocumento'
            Title.Alignment = taCenter
            Title.Caption = 'Passaporte'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Kit_PS'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'InseridoProgramacaoTransporte'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'TM'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RequisitantePT'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Req. PT'
            Width = 56
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CriadoPor'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Criado Por'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataCriacao'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Data [Cria'#231#227'o]'
            Width = 87
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'ComputadorCriacao'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Computador [Cria'#231#227'o]'
            Width = 114
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'AvaliadoPorProgramacao'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Avalia'#231#227'o [Programa'#231#227'o]'
            Width = 134
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataAvaliacaoProgramacao'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Data [Programa'#231#227'o]'
            Width = 114
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'ComputadorProgramacao'
            ReadOnly = True
            Title.Alignment = taCenter
            Title.Caption = 'Computador [Programa'#231#227'o]'
            Width = 151
            Visible = True
          end>
      end
      object StatusBarExecutantes: TStatusBar
        Left = 0
        Top = 531
        Width = 1082
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
        ExplicitTop = 523
        ExplicitWidth = 1080
      end
      object RLTemporario: TStringGrid
        Left = 128
        Top = 288
        Width = 179
        Height = 97
        FixedCols = 0
        FixedRows = 0
        TabOrder = 2
        Visible = False
      end
      object ToolBar1: TToolBar
        Left = 0
        Top = 25
        Width = 1082
        Height = 29
        ButtonHeight = 23
        ButtonWidth = 25
        Caption = 'ToolBar1'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 3
        ExplicitWidth = 1080
        object btnSelAll: TToolButton
          Left = 0
          Top = 0
          Hint = 'Selecionar todos os registros'
          Caption = 'btnSelAll'
          ImageIndex = 232
          ParentShowHint = False
          ShowHint = True
        end
        object btnSelClear: TToolButton
          Left = 25
          Top = 0
          Hint = 'Desmarcar todos os registros'
          Caption = 'btnSelClear'
          ImageIndex = 231
          ParentShowHint = False
          ShowHint = True
        end
        object btnFiltros: TToolButton
          Left = 50
          Top = 0
          Action = actProcurarSomenteTransbordos
          DropdownMenu = PopupMenuFiltros
          ImageIndex = 27
          Style = tbsDropDown
        end
        object ToolButton4: TToolButton
          Left = 94
          Top = 0
          Action = actConfigurarRT
          ParentShowHint = False
          ShowHint = True
        end
        object btnClearFiltroExecutante: TToolButton
          Left = 119
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnClearFiltroExecutante'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object btnLayoutExecutante: TToolButton
          Left = 144
          Top = 0
          Hint = 'Editar layout de colunas'
          Caption = 'btnLayoutExecutante'
          ImageIndex = 196
          ParentShowHint = False
          ShowHint = True
        end
        object btnExcelExecutante: TToolButton
          Left = 169
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'Exportar'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
        object BitBtn3: TBitBtn
          Left = 194
          Top = 0
          Width = 70
          Height = 23
          Action = actImprmir
          Align = alLeft
          Caption = 'Imprimir'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
        object BitBtn11: TBitBtn
          Left = 264
          Top = 0
          Width = 85
          Height = 23
          Action = actPrepararDadosRT
          Align = alLeft
          Caption = 'Prepara'#231#227'o'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 4
        end
        object BitBtn13: TBitBtn
          Left = 349
          Top = 0
          Width = 85
          Height = 23
          Action = actTransbordo
          Align = alLeft
          Caption = 'Transbordos'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 5
        end
        object BitBtn2: TBitBtn
          Left = 434
          Top = 0
          Width = 93
          Height = 23
          Action = actRT_TMIB
          Align = alLeft
          Caption = 'RT Embarque'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
        end
        object BitBtn1: TBitBtn
          Left = 527
          Top = 0
          Width = 104
          Height = 23
          Action = actGerarMultiplasRTs
          Align = alLeft
          Caption = 'Script RT-SAP'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 6
        end
        object BitBtn9: TBitBtn
          Left = 631
          Top = 0
          Width = 74
          Height = 23
          Action = actLimpar
          Align = alLeft
          Caption = 'Limpar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
        end
        object ToolButton11: TToolButton
          Left = 705
          Top = 0
          Width = 8
          Caption = 'ToolButton11'
          ImageIndex = 36
          Style = tbsSeparator
        end
        object CheckBoxOrigemDestino: TCheckBox
          Left = 713
          Top = 0
          Width = 173
          Height = 23
          Caption = 'DESTINO diferente da ORIGEM'
          TabOrder = 1
          OnClick = actProcurarProgramacaoExecutanteExecute
        end
      end
      object ColunasLayoutExecutanteProgramado: TStringGrid
        Left = 44
        Top = 96
        Width = 165
        Height = 129
        ColCount = 7
        RowCount = 42
        TabOrder = 4
        Visible = False
        RowHeights = (
          24
          24
          21
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24)
      end
      object MemoSAP: TMemo
        Left = 344
        Top = 288
        Width = 185
        Height = 89
        Lines.Strings = (
          'MemoSAP')
        TabOrder = 5
        Visible = False
      end
      object Panel2: TPanel
        Left = 0
        Top = 0
        Width = 1082
        Height = 25
        Align = alTop
        Caption = 'Executantes Programados por Per'#237'odo'
        Color = clSilver
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 6
        ExplicitWidth = 1080
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Espelho de Cria'#231#227'o/Cancelamento de RTs'
      ImageIndex = 2
      object ToolBar3: TToolBar
        Left = 0
        Top = 25
        Width = 1082
        Height = 29
        ButtonHeight = 23
        Caption = 'ToolBar3'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 0
        object btnClearFiltroRT: TToolButton
          Left = 0
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnClearFiltroRT'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object btnExcelRT: TToolButton
          Left = 23
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'btnExcelRT'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
        object btnLaypoutRT: TToolButton
          Left = 46
          Top = 0
          Hint = 'Editar layout de colunas'
          Caption = 'btnLaypoutRT'
          ImageIndex = 196
          ParentShowHint = False
          ShowHint = True
        end
        object btnCancelarForcado: TBitBtn
          Left = 69
          Top = 0
          Width = 104
          Height = 23
          Action = actCancelarRTForcado
          Align = alLeft
          Caption = 'Cancelar RT'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
        object BitBtn10: TBitBtn
          Left = 173
          Top = 0
          Width = 104
          Height = 23
          Action = ConciliarRTsOrfas
          Align = alLeft
          Caption = 'Conciliar RTs'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
      end
      object StatusBarGestaoRT: TStatusBar
        Left = 0
        Top = 531
        Width = 1082
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
      end
      object DBGridProgramcaoRT: TFilterDBGrid
        Left = 0
        Top = 54
        Width = 1082
        Height = 477
        Align = alClient
        DataSource = FrmDataModule.DataSourceGestaoRT
        TabOrder = 2
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGridProgramcaoRTDrawColumnCell
        ClearFilterButton = btnClearFiltroRT
        SearchAction = actProcurarProgramacaoRT
        LayoutGrid = ColunasLayoutRT
        EnableZebra = False
        ProgressBar = FrmPrincipal.ProgressBarPrincipal
        LayoutButton = btnLaypoutRT
        ExcelButton = btnExcelRT
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'idProgramacaoRT'
            Title.Alignment = taCenter
            Title.Caption = 'ID Espelho'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'idProgramacaoExecutante'
            Title.Alignment = taCenter
            Title.Caption = 'ID Executante'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CodigoSAP'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo SAP'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Status'
            Title.Alignment = taCenter
            Title.Caption = 'Status'
            Width = 135
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Orfa'
            Title.Alignment = taCenter
            Title.Caption = 'RT Orf'#227
            Width = 57
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_PendenteCancelamento'
            Title.Alignment = taCenter
            Title.Caption = 'Pend. Canc.'
            Width = 66
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_MotivoConciliacao'
            Title.Alignment = taCenter
            Title.Caption = 'Concilia'#231#227'o'
            Width = 90
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'booleanRecolhimento'
            Title.Alignment = taCenter
            Title.Caption = 'Recolhimento'
            Width = 73
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
            FieldName = 'txtDestino'
            Title.Alignment = taCenter
            Title.Caption = 'Destino'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RecolherPara'
            Title.Alignment = taCenter
            Title.Caption = 'Recolher Para'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataIda'
            Title.Alignment = taCenter
            Title.Caption = 'Data Ida'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataVolta'
            Title.Alignment = taCenter
            Title.Caption = 'Data Volta'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_HoraIda'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Ida'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_HoraVolta'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Volta'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Mensagem'
            Title.Alignment = taCenter
            Title.Caption = 'Mensagem'
            Width = 147
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Descricao'
            Title.Alignment = taCenter
            Title.Caption = 'Descri'#231#227'o'
            Width = 229
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Cancelada'
            Title.Alignment = taCenter
            Title.Caption = 'Cancelada'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Cobertura'
            Title.Alignment = taCenter
            Title.Caption = 'Cobertura'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Tipo'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Modal'
            Title.Alignment = taCenter
            Title.Caption = 'Modal'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Classe'
            Title.Alignment = taCenter
            Title.Caption = 'Classe'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_CentroPlan'
            Title.Alignment = taCenter
            Title.Caption = 'Centro Plan.'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_GrpPlan'
            Title.Alignment = taCenter
            Title.Caption = 'Grp Plan.'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_PessoaContato'
            Title.Alignment = taCenter
            Title.Caption = 'Pessoa Contato'
            Width = 106
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_Requisitante'
            Title.Alignment = taCenter
            Title.Caption = 'Requisitante'
            Width = 85
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_TelContato'
            Title.Alignment = taCenter
            Title.Caption = 'Tel. Contato'
            Width = 76
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_CentroCusto'
            Title.Alignment = taCenter
            Title.Caption = 'Centro de Custo'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_DiagramaRede'
            Title.Alignment = taCenter
            Title.Caption = 'Diagrama de Rede'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_OperRede'
            Title.Alignment = taCenter
            Title.Caption = 'Oper. Rede'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'RT_ElementoPEP'
            Title.Alignment = taCenter
            Title.Caption = 'Elemento PEP'
            Visible = True
          end>
      end
      object ColunasLayoutRT: TStringGrid
        Left = 24
        Top = 128
        Width = 138
        Height = 121
        ColCount = 7
        RowCount = 32
        FixedRows = 0
        TabOrder = 3
        Visible = False
        RowHeights = (
          24
          26
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24)
      end
      object Panel1: TPanel
        Left = 0
        Top = 0
        Width = 1082
        Height = 25
        Align = alTop
        Caption = 'Registro de RTs'
        Color = clSilver
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 4
      end
      object RLImpressao: TStringGrid
        Left = 232
        Top = 280
        Width = 320
        Height = 120
        TabOrder = 5
        Visible = False
      end
    end
    object TabSheet4: TTabSheet
      Caption = 'RTs importadas do SAP'
      ImageIndex = 3
      object DBGridRTSapImport: TFilterDBGrid
        Left = 0
        Top = 54
        Width = 1082
        Height = 477
        Align = alClient
        DataSource = FrmDataModule.DataSourceSAPImport
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        ClearFilterButton = btnClearFiltroSAPImport
        SearchAction = actProcurarSAPImport
        LayoutGrid = ColunasLayoutSAPImport
        EnableZebra = False
        ProgressBar = FrmPrincipal.ProgressBarPrincipal
        LayoutButton = btnLayoutSAPImport
        ExcelButton = btnExcelSAPImport
        Columns = <
          item
            Expanded = False
            FieldName = 'DataImportacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Importa'#231#227'o'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'OrigemConsulta'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PeriodoInicio'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'PeriodoFim'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'QMNUM'
            Title.Caption = 'RT'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'QMART'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'IWERK'
            Title.Alignment = taCenter
            Title.Caption = 'Centro'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'INGRP'
            Title.Alignment = taCenter
            Title.Caption = 'Grp. Plan.'
            Width = 85
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Origem'
            Title.Alignment = taCenter
            Width = 76
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'txtDestino'
            Title.Alignment = taCenter
            Title.Caption = 'Destino'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataViagem'
            Title.Alignment = taCenter
            Title.Caption = 'Data Viagem'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'HoraViagem'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Viagem'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'PERNR'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo SAP'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'TipoDoc'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo Doc'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Documento'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Passageiro'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'QMTXT'
            Title.Alignment = taCenter
            Title.Caption = 'Descri'#231#227'o'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Modal'
            Title.Alignment = taCenter
            Title.Caption = 'Modal'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Classe'
            Title.Alignment = taCenter
            Title.Caption = 'Classe'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'StatusItem'
            Title.Alignment = taCenter
            Title.Caption = 'Status Item'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'StatusDescricao'
            Title.Alignment = taCenter
            Title.Caption = 'Status Descri'#231#227'o'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Cancelada'
            Title.Alignment = taCenter
            Title.Caption = 'Cancelada'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ChavePassageiro'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ChaveIda'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ChaveVolta'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ChaveCompleta'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'idProgramacaoRT'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'idProgramacaoExecutante'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'ImportadoColibri'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Observacao'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_Orfa'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RT_PendenteCancelamento'
            Visible = True
          end>
      end
      object ToolBar4: TToolBar
        Left = 0
        Top = 25
        Width = 1082
        Height = 29
        ButtonHeight = 23
        Caption = 'ToolBar3'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 1
        object btnClearFiltroSAPImport: TToolButton
          Left = 0
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnClearFiltroRT'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object btnExcelSAPImport: TToolButton
          Left = 23
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'btnExcelRT'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
        object btnLayoutSAPImport: TToolButton
          Left = 46
          Top = 0
          Hint = 'Editar layout de colunas'
          Caption = 'btnLaypoutRT'
          ImageIndex = 196
          ParentShowHint = False
          ShowHint = True
        end
        object BitBtn7: TBitBtn
          Left = 69
          Top = 0
          Width = 116
          Height = 23
          Action = actImportRTSAP
          Align = alLeft
          Caption = 'Importar RT SAP'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
        object BitBtn12: TBitBtn
          Left = 185
          Top = 0
          Width = 114
          Height = 23
          Action = actHistorizarProgramacao_DadosSAP
          Align = alLeft
          Caption = 'Historizar SAP'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
      end
      object ColunasLayoutSAPImport: TStringGrid
        Left = 32
        Top = 136
        Width = 177
        Height = 129
        ColCount = 7
        RowCount = 32
        FixedRows = 0
        TabOrder = 2
        Visible = False
        RowHeights = (
          24
          26
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24
          24)
      end
      object StatusBarSAPImport: TStatusBar
        Left = 0
        Top = 531
        Width = 1082
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
      end
      object Panel5: TPanel
        Left = 0
        Top = 0
        Width = 1082
        Height = 25
        Align = alTop
        Caption = 'Requisi'#231#245'es de Transporte Importadas do SAP'
        Color = clSilver
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 4
      end
    end
    object TabSheet3: TTabSheet
      Caption = 'Auditoria de Roteiro de Embarques'
      ImageIndex = 3
      object DBGridResumoFrequencia: TFilterDBGrid
        Left = 0
        Top = 54
        Width = 1082
        Height = 477
        Align = alClient
        DataSource = FrmDataModule.DataSourceFrequenciaResumo
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        ClearFilterButton = btnClearFiltroBuscaEmbarque
        SearchAction = actProcurarBuscaEmbarque
        LayoutGrid = RLLayoutBuscaEmbarque
        EnableZebra = False
        ProgressBar = FrmPrincipal.ProgressBarPrincipal
        LayoutButton = btnLauoutBuscaEmbarque
        ExcelButton = btnExcelBuscaEmbarque
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataProgramacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Programa'#231#227'o'
            Width = 103
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Origem'
            Title.Alignment = taCenter
            Width = 92
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'txtDestino'
            Title.Alignment = taCenter
            Title.Caption = 'Destino'
            Width = 91
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CodigoSAP'
            Title.Alignment = taCenter
            Width = 71
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'NomeExecutante'
            Title.Alignment = taCenter
            Title.Caption = 'Nome Executante'
            Width = 258
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Funcao'
            Title.Alignment = taCenter
            Title.Caption = 'Fun'#231#227'o'
            Width = 207
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Empresa'
            Title.Alignment = taCenter
            Width = 140
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'HoraRoteamento'
            Title.Alignment = taCenter
            Title.Caption = 'Hora Rota'
            Width = 99
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'NomeEmbarcacao'
            Title.Alignment = taCenter
            Title.Caption = 'Embarca'#231#227'o'
            Width = 117
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'NomeRota'
            Title.Alignment = taCenter
            Title.Caption = 'Descri'#231#227'o da Rota'
            Width = 111
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RotaSequencia'
            Title.Alignment = taCenter
            Title.Caption = 'Roteiro'
            Visible = True
          end>
      end
      object RLLayoutBuscaEmbarque: TStringGrid
        Left = 40
        Top = 144
        Width = 177
        Height = 129
        ColCount = 7
        RowCount = 11
        FixedRows = 0
        TabOrder = 1
        Visible = False
        RowHeights = (
          24
          26
          24
          24
          24
          24
          24
          24
          24
          24
          24)
      end
      object ToolBar2: TToolBar
        Left = 0
        Top = 25
        Width = 1082
        Height = 29
        ButtonHeight = 23
        Caption = 'ToolBar3'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 2
        object btnClearFiltroBuscaEmbarque: TToolButton
          Left = 0
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnClearFiltroRT'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object btnExcelBuscaEmbarque: TToolButton
          Left = 23
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'btnExcelRT'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
        object btnLauoutBuscaEmbarque: TToolButton
          Left = 46
          Top = 0
          Hint = 'Editar layout de colunas'
          Caption = 'btnLaypoutRT'
          ImageIndex = 196
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton1: TToolButton
          Left = 69
          Top = 0
          Action = actProcurarBuscaEmbarque
        end
      end
      object Panel6: TPanel
        Left = 0
        Top = 0
        Width = 1082
        Height = 25
        Align = alTop
        Caption = 'Requisi'#231#245'es de Transporte Importadas do SAP'
        Color = clSilver
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 3
      end
      object StatusBarFrequenciaResumo: TStatusBar
        Left = 0
        Top = 531
        Width = 1082
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
          end>
      end
    end
    object TabSheet5: TTabSheet
      Caption = 'Simulacao Logistica'
      ImageIndex = 4
      object GridSimulacaoParametros: TStringGrid
        Left = 0
        Top = 54
        Width = 445
        Height = 477
        Align = alLeft
        ColCount = 3
        DefaultRowHeight = 22
        FixedCols = 0
        RowCount = 19
        FixedRows = 1
        TabOrder = 0
      end
      object SplitterSimulacao: TSplitter
        Left = 445
        Top = 54
        Height = 477
      end
      object MemoSimulacaoRelatorio: TMemo
        Left = 448
        Top = 54
        Width = 634
        Height = 477
        Align = alClient
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Consolas'
        Font.Style = []
        ParentFont = False
        ScrollBars = ssVertical
        TabOrder = 1
      end
      object StatusBarSimulacao: TStatusBar
        Left = 0
        Top = 531
        Width = 1082
        Height = 19
        Panels = <>
        SimplePanel = True
      end
      object ToolBarSimulacao: TToolBar
        Left = 0
        Top = 25
        Width = 1082
        Height = 29
        ButtonHeight = 23
        Caption = 'ToolBarSimulacao'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 3
        object btnSimulacaoPadrao: TBitBtn
          Left = 0
          Top = 0
          Width = 120
          Height = 23
          Caption = 'Cenario base'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          OnClick = btnSimulacaoPadraoClick
        end
        object btnSimulacaoRodar: TBitBtn
          Left = 120
          Top = 0
          Width = 128
          Height = 23
          Caption = 'Rodar simulacao'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnClick = btnSimulacaoRodarClick
        end
      end
      object PanelSimulacao: TPanel
        Left = 0
        Top = 0
        Width = 1082
        Height = 25
        Align = alTop
        Caption = 'Simulacao logistica comparativa por cenario'
        Color = clSilver
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 4
      end
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 600
    Top = 152
    StyleName = 'Platform Default'
    object actProcurarProgramacaoExecutante: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      ShortCut = 116
      OnExecute = actProcurarProgramacaoExecutanteExecute
    end
    object actImprmir: TAction
      Caption = 'Imprimir'
      Hint = 'Imprimir lista de assinaturas'
      ImageIndex = 21
      OnExecute = actImprmirExecute
    end
    object actNumDias: TAction
      Caption = 'Total Dias'
      Hint = 'Calcular total de dias dos registros filtrados'
      ImageIndex = 3
      OnExecute = actNumDiasExecute
    end
    object actStatusSELECIONADO: TAction
      Category = 'Status'
      Caption = 'Calcular Status Programa'#231#227'o (Selecionado)'
      Hint = 
        'Calcular "N'#176' Exec.", N'#176' Apro." e "N'#176' Canc." para o registro sele' +
        'cionado'
      ImageIndex = 35
      ShortCut = 119
      OnExecute = actStatusSELECIONADOExecute
    end
    object actStatusTODOS: TAction
      Category = 'Status'
      Caption = 'Calcular Status Programa'#231#227'o (Todos)'
      Hint = 
        'Calcular "N'#176' Exec.", N'#176' Apro." e "N'#176' Canc." para todos os regist' +
        'ros FILTRADOS'
      ImageIndex = 35
      OnExecute = actStatusTODOSExecute
    end
    object actTransbordo: TAction
      Caption = 'Transbordos'
      Hint = 'Filtrar Transbordos'
      ImageIndex = 456
      OnExecute = actTransbordoExecute
    end
    object actMotivo: TAction
      Caption = 'actMotivo'
      OnExecute = actMotivoExecute
    end
    object actGerarMultiplasRTs: TAction
      Category = 'Script SAP'
      Caption = 'Script RT-SAP'
      Hint = 'Gerar RT para todas as programa'#231#245'es filtradas'
      ImageIndex = 463
      OnExecute = actGerarMultiplasRTsExecute
    end
    object actConfigurarRT: TAction
      Category = 'Script SAP'
      Hint = 'Configura'#231#227'o inicial para cria'#231#227'o de RT'
      ImageIndex = 487
      OnExecute = actConfigurarRTExecute
    end
    object actCancelarRTForcado: TAction
      Category = 'Espelho RT'
      Caption = 'Cancelar RT'
      Hint = 'Cancelar todas as RTs filtradas de forma for'#231'ada'
      ImageIndex = 463
      OnExecute = actCancelarRTForcadoExecute
    end
    object actProcurarProgramacaoRT: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarProgramacaoRTExecute
    end
    object actPrepararDadosRT: TAction
      Category = 'Script SAP'
      Caption = 'Prepara'#231#227'o'
      Hint = 'Auto preecnhimento'
      ImageIndex = 4
      OnExecute = actPrepararDadosRTExecute
    end
    object actSelAllSelecao: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'Marcar todos "Sele'#231#227'o"'
      ImageIndex = 232
      OnExecute = actSelAllSelecaoExecute
    end
    object actSelAllRecolhimento: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'Marcar todos "Recolhimento"'
      ImageIndex = 232
      OnExecute = actSelAllRecolhimentoExecute
    end
    object actSelClearSelecao: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'Desmarcar todos "Sele'#231#227'o"'
      ImageIndex = 231
      OnExecute = actSelClearSelecaoExecute
    end
    object actSelClearRecolhimento: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'Desmarcar todos "Recolhimento"'
      ImageIndex = 231
      OnExecute = actSelClearRecolhimentoExecute
    end
    object actProcurarSAPImport: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarSAPImportExecute
    end
    object actImportRTSAP: TAction
      Category = 'Script SAP'
      Caption = 'Importar RT SAP'
      Hint = 'Importar RTs do SAP'
      ImageIndex = 463
      OnExecute = actImportRTSAPExecute
    end
    object actProgramacaoRTxProgramacaoExecutante: TAction
      Category = 'Script SAP'
      Caption = 
        'Carregar dados da tabela "Registros de RT" para a tabela "Execut' +
        'antes Programados"'
      Hint = 
        'Carregar dados da tabela "Registros de RT" para a tabela "Execut' +
        'antes Programados"'
      ImageIndex = 24
      OnExecute = actProgramacaoRTxProgramacaoExecutanteExecute
    end
    object actHistorizarProgramacao_DadosSAP: TAction
      Category = 'Script SAP'
      Caption = 'Historizar SAP'
      Hint = 
        'Carregar o n'#250'mero da RTs importadas do SAP para as tabelas: "Reg' +
        'istro de RTs" e "Executantes Programados"'
      ImageIndex = 463
      OnExecute = actHistorizarProgramacao_DadosSAPExecute
    end
    object ConciliarRTsOrfas: TAction
      Category = 'Espelho RT'
      Caption = 'Conciliar RTs'
      Hint = 'An'#225'lise para conciliar RTs Orf'#227's'
      ImageIndex = 8
      OnExecute = ConciliarRTsOrfasExecute
    end
    object actRT_TMIB: TAction
      Category = 'Script SAP'
      Caption = 'RT Embarque'
      Hint = 'Lista de executantes com RT de Embarque'
      ImageIndex = 452
      OnExecute = actRT_TMIBExecute
    end
    object actProcurarSomenteTransbordos: TAction
      Category = 'Procurar'
      Caption = 'Filtrar transbordos'
      Hint = 'Filtrar transbordos'
      ImageIndex = 456
      OnExecute = actProcurarSomenteTransbordosExecute
    end
    object actLimpar: TAction
      Category = 'Script SAP'
      Caption = 'Limpar'
      Hint = 'Limpar'
      ImageIndex = 43
      OnExecute = actLimparExecute
    end
    object actProcurarBuscaEmbarque: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarBuscaEmbarqueExecute
    end
  end
  object SavePictureDialog1: TSavePictureDialog
    Left = 276
    Top = 231
  end
  object PopupMenuStatus: TPopupMenu
    Left = 608
    Top = 209
    object CalcularStatusExecSelecionado1: TMenuItem
      Action = actStatusSELECIONADO
    end
    object CalcularStatusExec1: TMenuItem
      Action = actStatusTODOS
    end
  end
  object SaveDialog1: TSaveDialog
    DefaultExt = '*.vbs'
    FileName = 'Script.vbs'
    Filter = 'Arquivo de texto (*.vbs)|*.vbs'
    Left = 272
    Top = 176
  end
  object PopupMenuSelAll: TPopupMenu
    Left = 732
    Top = 186
    object MarcartodosRecolhimento1: TMenuItem
      Action = actSelAllRecolhimento
    end
    object MarcartodosSeleo1: TMenuItem
      Action = actSelAllSelecao
    end
  end
  object PopupMenuSelClear: TPopupMenu
    Left = 604
    Top = 322
    object DesmarcartodosRecolhimento1: TMenuItem
      Action = actSelClearRecolhimento
    end
    object DesmarcartodosSeleo1: TMenuItem
      Action = actSelClearSelecao
    end
  end
  object PopupMenuFuncoes: TPopupMenu
    Left = 604
    Top = 266
    object AvaliarnecessidadedecriaodeRTs1: TMenuItem
      Caption = 'Avaliar necessidade de cria'#231#227'o de RTs'
      Hint = 'Avaliar necessidade de cria'#231#227'o de RTs'
      ImageIndex = 2
    end
    object Aplicarregrasautomticas1: TMenuItem
      Caption = 'Aplicar regras autom'#225'ticas'
    end
    object SincronizarProgramaodeExecutantecomasRTsExistentesnoSAP1: TMenuItem
      Caption = 
        'Sincronizar Programa'#231#227'o de Executante com as RTs existentes na p' +
        'rograma'#231#227'o de RT'
    end
    object CarregardadosdatabelaRegistrosdeRTparaatabelaExecutantesProgramados1: TMenuItem
      Action = actProgramacaoRTxProgramacaoExecutante
    end
    object CarregaronmerodaRTsimportadasdoSAPparaastabelasRegistrodeRTseExecutantesProgramados1: TMenuItem
      Action = actHistorizarProgramacao_DadosSAP
    end
    object Preparao1: TMenuItem
      Action = actPrepararDadosRT
    end
    object actReprocessar1: TMenuItem
      Caption = 'Reprocessar chaves do per'#237'odo'
    end
    object actConcliliar1: TMenuItem
      Caption = 'Conciliar '#243'rf'#227's'
    end
    object actAuditarSemExecutante1: TMenuItem
      Caption = 'Auditar sem Executante local'
    end
    object actResumoAuditoriaSemExecutante1: TMenuItem
      Caption = 'consulta-resumo para contar quantos ficaram em cada situa'#231#227'o'
    end
  end
  object PopupMenuFiltros: TPopupMenu
    Left = 852
    Top = 266
    object actProcurarSomenteTransbordos1: TMenuItem
      Action = actProcurarSomenteTransbordos
    end
  end
end
