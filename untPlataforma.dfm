object FrmPlataforma: TFrmPlataforma
  Left = 0
  Top = 0
  Caption = 'Cadastro de Instala'#231#227'o'
  ClientHeight = 403
  ClientWidth = 1180
  Color = clMenu
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  FormStyle = fsMDIChild
  Position = poScreenCenter
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 14
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 1180
    Height = 25
    Align = alTop
    Caption = 'Cadastro de Instala'#231#227'o'
    Color = clGradientInactiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 1178
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 384
    Width = 1180
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
    ExplicitTop = 376
    ExplicitWidth = 1178
  end
  object DBGridPlataformas: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 1180
    Height = 330
    Align = alClient
    DataSource = FrmDataModule.DataSourcePlataforma
    TabOrder = 2
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
    OnCellClick = DBGridPlataformasCellClick
    OnDrawColumnCell = DBGridPlataformasDrawColumnCell
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
        FieldName = 'HoraSaidaOrigem'
        Title.Alignment = taCenter
        Title.Caption = 'Hora Sa'#237'da Origem'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'booleanPlataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Controle'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'RT_Modal'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'booleanTurno2'
        Title.Caption = '2'#186' Turno'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'booleanOrigem'
        Title.Alignment = taCenter
        Title.Caption = 'Origem'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Plataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Instala'#231#227'o'
        Width = 162
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NomeSAP'
        Title.Alignment = taCenter
        Title.Caption = 'SAP'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'booleanNaoCriarRT'
        Title.Alignment = taCenter
        Title.Caption = 'N'#227'o Criar RT'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'booleanProntidao'
        Title.Alignment = taCenter
        Title.Caption = 'Prontid'#227'o'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Salvatagem'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'ElevacaoTopDECK'
        Title.Alignment = taCenter
        Title.Caption = 'Eleva'#231#227'o conv'#233's superior (El) [cm]'
        Width = 200
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'InicioOperacao'
        Title.Alignment = taCenter
        Title.Caption = 'In'#237'cio da opera'#231#227'o'
        Width = 150
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'LaminaAgua'
        Title.Alignment = taCenter
        Title.Caption = 'L'#226'mina d'#39#225'gua [m]'
        Width = 150
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Latitude'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Longitude'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CoordX'
        Title.Alignment = taCenter
        Title.Caption = 'Coord. X'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CoordY'
        Title.Alignment = taCenter
        Title.Caption = 'Coord. Y'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AvaliadoPor'
        Title.Alignment = taCenter
        Title.Caption = 'Avaliado Por'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataAtualizacao'
        Title.Alignment = taCenter
        Title.Caption = 'Data Atualiza'#231#227'o'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Computador'
        ReadOnly = True
        Title.Alignment = taCenter
        Width = 70
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CentroCusto'
        Title.Alignment = taCenter
        Title.Caption = 'Centro de Custo'
        Width = 105
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DiagramaRede'
        Title.Alignment = taCenter
        Title.Caption = 'Diagrama de Rede'
        Width = 99
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'OperRede'
        Title.Alignment = taCenter
        Title.Caption = 'Oper. Rede'
        Width = 73
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'ElementoPEP'
        Title.Alignment = taCenter
        Title.Caption = 'Elemento PEP'
        Width = 97
        Visible = True
      end>
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 1180
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 3
    ExplicitWidth = 1178
    object DBNavigatorCadastro: TDBNavigator
      Left = 0
      Top = 0
      Width = 144
      Height = 22
      DataSource = FrmDataModule.DataSourcePlataforma
      VisibleButtons = [nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh]
      Hints.Strings = (
        'Primeiro'
        'Anterior'
        'Pr'#243'ximo'
        #218'ltimo'
        'Inserir'
        'Deletar'
        'Editar'
        'Carregar'
        'Cancelar'
        'Atualizar')
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
    object ToolButton2: TToolButton
      Left = 144
      Top = 0
      Width = 8
      Caption = 'ToolButton2'
      ImageIndex = 58
      Style = tbsSeparator
    end
    object btnClearFiltro: TToolButton
      Left = 152
      Top = 0
      Hint = 'Limpar filtro'
      Caption = 'btnClearFiltro'
      ImageIndex = 225
      ParentShowHint = False
      ShowHint = True
    end
    object btnLayout: TToolButton
      Left = 175
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'btnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 198
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Caption = 'Exportar'
      ImageIndex = 54
      ParentShowHint = False
      ShowHint = True
    end
    object BitBtn5: TBitBtn
      Left = 221
      Top = 0
      Width = 97
      Height = 22
      Action = actCalcularXY
      Align = alLeft
      Caption = 'Calcular (X,Y)'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
    end
    object DBEdit1: TDBEdit
      Left = 318
      Top = 0
      Width = 121
      Height = 22
      DataField = 'idPlataforma'
      DataSource = FrmDataModule.DataSourcePlataforma
      TabOrder = 1
      Visible = False
    end
  end
  object ColunasLayout: TStringGrid
    Left = 521
    Top = 160
    Width = 248
    Height = 145
    ColCount = 7
    DefaultRowHeight = 21
    RowCount = 24
    TabOrder = 4
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
      21
      21)
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 216
    Top = 194
    StyleName = 'Platform Default'
    object actPanTo: TAction
      Category = 'Ferramentas'
      Caption = 'actPanTo'
      Hint = 'Posicionar (Latitude, Longitude)'
      ImageIndex = 290
    end
    object actImprimir: TAction
      Category = 'Ferramentas'
      Caption = 'actImprimir'
      Hint = 'Imprimir o mapa'
      ImageIndex = 21
    end
    object actSalvar: TAction
      Category = 'Ferramentas'
      Caption = 'actSalvar'
      Hint = 'Salvar mapa como imagem'
      ImageIndex = 264
    end
    object actAtualizarTituo: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actAtualizarTituo'
      OnExecute = actAtualizarTituoExecute
    end
    object actMarcar: TAction
      Category = 'Ferramentas'
      Caption = 'actMarcar'
      Hint = 'Marcador'
      ImageIndex = 112
    end
    object actProcurar: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
    object actProcurarTabuaMare: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actProcurarTabuaMare'
    end
    object actCalcularXY: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Calcular (X,Y)'
      Hint = 'Calcular coordenadas X,Y com base nas Latitudes e Longitudes'
      ImageIndex = 51
      OnExecute = actCalcularXYExecute
    end
    object actZoomFit: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomFit'
      Hint = 'Zoom Fit'
      ImageIndex = 45
    end
    object actPan: TAction
      Category = 'Ferramentas'
      Caption = 'actPan'
      Hint = 'Pan'
      ImageIndex = 41
    end
    object actZoomDinamico: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomDinamico'
      Hint = 'Zoom din'#226'mico'
      ImageIndex = 48
    end
    object actZoomWindow: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomWindow'
      Hint = 'Zoom janela'
      ImageIndex = 46
    end
    object actClearGrafico: TAction
      Category = 'Ferramentas'
      Caption = 'actClearGrafico'
      Hint = 'Limpar o Mapa'
      ImageIndex = 324
    end
    object actMapaMundi: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Mapa Mundi'
    end
    object actBrasil: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Brasil'
    end
    object actLatitudeLongitude: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Latitudes e Longitudes'
    end
    object actCarregarLinhas: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Carregar linhas 2D (x1;y1;x2;y2)...'
      Hint = 
        'Carregar linhas a partir de arquivo *.txt com as coordenadas das' +
        ' linhas no formato: x1;y1;x2;y2 '
    end
    object actAbrirDesenho: TAction
      Category = 'Ferramentas'
      Caption = 'actAbrirDesenho'
      Hint = 'Abrir arquivo de desenho 2D (*.zuchi)...'
      ImageIndex = 20
    end
    object actSalvarDesenhoData: TAction
      Category = 'Ferramentas'
      Caption = 'actSalvarDesenhoData'
      Hint = 'Salvar arquivo como desenho 2D (*.zuchi)...'
      ImageIndex = 462
    end
    object actTabuaMare: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Calcular'
      Hint = 'Consultar altura da mar'#233' agora'
      ImageIndex = 458
    end
  end
  object SavePictureDialog1: TSavePictureDialog
    Filter = 'Metafiles (*.wmf)|*.wmf'
    Left = 152
    Top = 120
  end
  object PopupMenuMapa: TPopupMenu
    Left = 181
    Top = 123
    object MapaMundi1: TMenuItem
      Action = actMapaMundi
    end
    object LatitudeseLongitudes1: TMenuItem
      Action = actLatitudeLongitude
    end
    object Brasil1: TMenuItem
      Action = actBrasil
    end
    object Carregarlinhas2Dx1y1x2y21: TMenuItem
      Action = actCarregarLinhas
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 320
    Top = 128
  end
  object SaveDialog1: TSaveDialog
    Left = 437
    Top = 115
  end
end
