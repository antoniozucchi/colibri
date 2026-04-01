object FrmEmbarcacao: TFrmEmbarcacao
  Left = 0
  Top = 0
  Caption = 'Cadastro de Embarca'#231#245'es'
  ClientHeight = 365
  ClientWidth = 740
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
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 740
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 0
    ExplicitWidth = 738
    object DBNavigator1: TDBNavigator
      Left = 0
      Top = 0
      Width = 138
      Height = 22
      DataSource = FrmDataModule.DataSourceEmbarcacoes
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
  end
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 740
    Height = 25
    Align = alTop
    Caption = 'Cadastro de Embarca'#231#245'es'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 738
  end
  object DBGridEmbarcacao: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 740
    Height = 292
    Align = alClient
    DataSource = FrmDataModule.DataSourceEmbarcacoes
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridEmbarcacaoDrawColumnCell
    OnKeyPress = DBGridEmbarcacaoKeyPress
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    LayoutGrid = ColunasLayout
    EnableZebra = False
    ProgressBar = FrmPrincipal.ProgressBarPrincipal
    LayoutButton = btnLayout
    ExcelButton = btnExcel
    Columns = <
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Distribuicao'
        Title.Alignment = taCenter
        Title.Caption = 'Distribui'#231#227'o'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'UsaBridgeMesmoGrupo'
        Title.Alignment = taCenter
        Title.Caption = 'Bridge'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'OrigemBridge'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NomeEmbarcacao'
        Title.Alignment = taCenter
        Title.Caption = 'Nome Embarca'#231#227'o'
        Width = 150
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'TipoEmbarcacao'
        PickList.Strings = (
          'Embarca'#231#227'o de apoio'
          'Lancha de passageiros'
          'Mergulho')
        Title.Alignment = taCenter
        Title.Caption = 'Tipo de Embarca'#231#227'o'
        Width = 125
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CapacidadePAX'
        Title.Alignment = taCenter
        Title.Caption = 'Capacidade (PAX)'
        Width = 100
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'StatusEmbarcacao'
        PickList.Strings = (
          'Operando'
          'Indispon'#237'vel'
          'Fora de opera'#231#227'o')
        Title.Alignment = taCenter
        Title.Caption = 'Status'
        Width = 120
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Observacao'
        Title.Alignment = taCenter
        Title.Caption = 'Observa'#231#227'o'
        Width = 250
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'VelocidadeLancha'
        Title.Alignment = taCenter
        Title.Caption = 'Velocidade (km/h)'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'TempoTransbordo'
        Title.Alignment = taCenter
        Title.Caption = 'Tempo Transbordo (min.)'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AvaliadoPor'
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Avaliado Por'
        Width = 70
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DataAtualizacao'
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Data Atualiza'#231#227'o'
        Width = 120
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Computador'
        ReadOnly = True
        Title.Alignment = taCenter
        Visible = True
      end>
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 346
    Width = 740
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
    ExplicitTop = 338
    ExplicitWidth = 738
  end
  object ColunasLayout: TStringGrid
    Left = 80
    Top = 166
    Width = 113
    Height = 94
    ColCount = 7
    DefaultRowHeight = 21
    RowCount = 12
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
      21)
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 272
    Top = 216
    StyleName = 'Platform Default'
    object actExcel: TAction
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
    end
    object actProcurar: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
  end
end
