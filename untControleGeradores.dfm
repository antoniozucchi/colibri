object FrmControleGeradores: TFrmControleGeradores
  Left = 0
  Top = 0
  Caption = 'Controle de Geradores'
  ClientHeight = 506
  ClientWidth = 922
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
  object Splitter1: TSplitter
    Left = 0
    Top = 376
    Width = 922
    Height = 3
    Cursor = crVSplit
    Align = alBottom
    ExplicitLeft = 1
    ExplicitTop = 204
    ExplicitWidth = 726
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 922
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 0
    ExplicitWidth = 920
    object DBNavigatorGerador: TDBNavigator
      Left = 0
      Top = 0
      Width = 138
      Height = 22
      DataSource = FrmDataModule.DataSourceGeradores
      VisibleButtons = [nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh]
      Enabled = False
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
    object btnLayput: TToolButton
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
    Width = 922
    Height = 25
    Align = alTop
    Caption = 'Controle de Geradores'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 920
  end
  object DBGridGerador: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 922
    Height = 322
    Align = alClient
    DataSource = FrmDataModule.DataSourceGeradores
    ReadOnly = True
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridGeradorDrawColumnCell
    OnKeyPress = DBGridGeradorKeyPress
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    EnableZebra = False
    LayoutButton = btnLayput
    ExcelButton = btnExcel
    Columns = <
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Status'
        PickList.Strings = (
          'Operando'
          'Fora de opera'#231#227'o')
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NomeGerador'
        Title.Alignment = taCenter
        Title.Caption = 'Gerador'
        Width = 127
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Plataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Local de Instala'#231#227'o'
        Width = 106
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'FrenteServico'
        Title.Alignment = taCenter
        Title.Caption = 'Frente de Servi'#231'o'
        Width = 200
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Tensao'
        Title.Alignment = taCenter
        Title.Caption = 'Tens'#227'o'
        Width = 60
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Potencia'
        Title.Alignment = taCenter
        Title.Caption = 'Pot'#234'ncia'
        Width = 60
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Propriedade'
        Title.Alignment = taCenter
        Title.Caption = 'Propriet'#225'rio'
        Width = 120
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'ValidadeEslingas'
        Title.Alignment = taCenter
        Title.Caption = 'Validade das Eslingas'
        Width = 113
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Notas'
        Title.Alignment = taCenter
        Title.Caption = 'Ocorr'#234'ncias'
        Width = 70
        Visible = True
      end
      item
        Alignment = taCenter
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
        Title.Alignment = taCenter
        Visible = True
      end>
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 487
    Width = 922
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
    ExplicitTop = 479
    ExplicitWidth = 920
  end
  object Panel2: TPanel
    Left = 0
    Top = 379
    Width = 922
    Height = 108
    Align = alBottom
    Caption = 'Panel2'
    TabOrder = 4
    ExplicitTop = 371
    ExplicitWidth = 920
    object PanelTituloPltaformas: TPanel
      Left = 1
      Top = 1
      Width = 920
      Height = 20
      Align = alTop
      Caption = 'Ocorr'#234'ncias'
      Color = clMenuBar
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentBackground = False
      ParentFont = False
      TabOrder = 0
      ExplicitWidth = 918
    end
    object DBMemo1: TDBMemo
      Left = 1
      Top = 21
      Width = 920
      Height = 86
      Align = alClient
      DataField = 'Notas'
      DataSource = FrmDataModule.DataSourceGeradores
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      ExplicitWidth = 918
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 312
    Top = 150
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
  end
end
