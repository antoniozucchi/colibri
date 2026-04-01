object FrmSituacaoEquipamentoAcesso: TFrmSituacaoEquipamentoAcesso
  Left = 0
  Top = 0
  Caption = 'Situa'#231#227'o dos Equipamentos e Acessos das Plataformas'
  ClientHeight = 545
  ClientWidth = 1003
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
    Top = 415
    Width = 1003
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
    Width = 1003
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 0
    ExplicitWidth = 819
    object DBNavigator: TDBNavigator
      Left = 0
      Top = 0
      Width = 132
      Height = 22
      DataSource = FrmDataModule.DataSourcePlataformaControle
      VisibleButtons = [nbEdit, nbPost, nbCancel, nbRefresh]
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
      Left = 132
      Top = 0
      Hint = 'Limpar filtro'
      Caption = 'btnClearFiltro'
      ImageIndex = 225
      ParentShowHint = False
      ShowHint = True
    end
    object btnLayout: TToolButton
      Left = 155
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'btnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 178
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
    Width = 1003
    Height = 25
    Align = alTop
    Caption = 'Situa'#231#227'o dos Equipamentos e Acessos das Plataformas'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 819
  end
  object DBGridSituacaoEquipamentoAcesso: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 1003
    Height = 361
    Align = alClient
    DataSource = FrmDataModule.DataSourcePlataformaControle
    ReadOnly = True
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridSituacaoEquipamentoAcessoDrawColumnCell
    OnKeyPress = DBGridSituacaoEquipamentoAcessoKeyPress
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
        FieldName = 'Plataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Local de Instala'#231#227'o'
        Width = 106
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoGD'
        Title.Alignment = taCenter
        Title.Caption = 'GD'
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CapPrincipal'
        Title.Caption = 'SWL Prin. (t)'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'CapAuxiliar'
        Title.Caption = 'SWL Aux. (t)'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoBCI'
        Title.Alignment = taCenter
        Title.Caption = 'BCI'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoLinhaBCI'
        Title.Alignment = taCenter
        Title.Caption = 'Linha BCI'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoBalsa'
        Title.Alignment = taCenter
        Title.Caption = 'Balsa'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoSurfer'
        Title.Alignment = taCenter
        Title.Caption = 'Surfer'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoSOV'
        Title.Alignment = taCenter
        Title.Caption = 'SOV'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoAqua'
        Title.Alignment = taCenter
        Title.Caption = #193'qua'
        Width = 65
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoDegraus'
        Title.Alignment = taCenter
        Title.Caption = 'Limp.Degraus'
        Width = 72
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataRealizacaoDegraus'
        Title.Alignment = taCenter
        Title.Caption = 'Data Realiza'#231#227'o'
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
    Top = 526
    Width = 1003
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
    ExplicitTop = 327
    ExplicitWidth = 819
  end
  object ColunasLayout: TStringGrid
    Left = 48
    Top = 102
    Width = 113
    Height = 94
    ColCount = 7
    DefaultRowHeight = 21
    RowCount = 14
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
      21)
  end
  object Panel2: TPanel
    Left = 0
    Top = 418
    Width = 1003
    Height = 108
    Align = alBottom
    Caption = 'Panel2'
    TabOrder = 5
    ExplicitTop = 219
    ExplicitWidth = 819
    object PanelTituloNotas: TPanel
      Left = 1
      Top = 1
      Width = 1001
      Height = 20
      Align = alTop
      Caption = 'Observa'#231#245'es'
      Color = clMenuBar
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = [fsBold]
      ParentBackground = False
      ParentFont = False
      TabOrder = 0
      ExplicitWidth = 817
    end
    object DBMemo1: TDBMemo
      Left = 1
      Top = 21
      Width = 1001
      Height = 86
      Align = alClient
      DataField = 'SituacaoNotas'
      DataSource = FrmDataModule.DataSourcePlataformaControle
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -13
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 1
      ExplicitWidth = 817
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
