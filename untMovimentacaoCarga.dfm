object FrmMovimentacaoCarga: TFrmMovimentacaoCarga
  Left = 0
  Top = 0
  Caption = 'Movimenta'#231#227'o de Carga'
  ClientHeight = 421
  ClientWidth = 806
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
    Width = 806
    Height = 25
    Align = alTop
    Caption = 'Movimenta'#231#227'o de Carga'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 804
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 50
    Width = 806
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 1
    ExplicitWidth = 804
    object DBNavigator1: TDBNavigator
      Left = 0
      Top = 0
      Width = 138
      Height = 22
      DataSource = FrmDataModule.DataSourceMovimentacaoCarga
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
    object brnLayout: TToolButton
      Left = 161
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'brnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 184
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Action = actExcel
      ParentShowHint = False
      ShowHint = True
    end
  end
  object DBGridMovimentacaoCarga: TFilterDBGrid
    Left = 0
    Top = 79
    Width = 806
    Height = 323
    Align = alClient
    DataSource = FrmDataModule.DataSourceMovimentacaoCarga
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridMovimentacaoCargaDrawColumnCell
    OnKeyPress = DBGridMovimentacaoCargaKeyPress
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    LayoutGrid = ColunasLayout
    EnableZebra = False
    LayoutButton = brnLayout
    ExcelButton = btnExcel
    Columns = <
      item
        Expanded = False
        FieldName = 'Status'
        PickList.Strings = (
          'Atendido'
          'Pendente'
          'Cancelado')
        Title.Alignment = taCenter
        Width = 128
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataNecessidade'
        Title.Alignment = taCenter
        Title.Caption = 'Data Necessidade'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataAtendimento'
        Title.Alignment = taCenter
        Title.Caption = 'Data Atendimento'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Descricao'
        Title.Alignment = taCenter
        Title.Caption = 'Descri'#231#227'o'
        Width = 215
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Carga'
        Title.Alignment = taCenter
        Title.Caption = 'Carga (kg)'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Origem'
        Title.Alignment = taCenter
        Width = 120
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Destino'
        Title.Alignment = taCenter
        Width = 120
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Responsavel'
        Title.Alignment = taCenter
        Title.Caption = 'Respons'#225'vel'
        Width = 150
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Observacao'
        Title.Alignment = taCenter
        Title.Caption = 'Observa'#231#227'o'
        Width = 200
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'AvaliadoPor'
        ReadOnly = True
        Title.Alignment = taCenter
        Title.Caption = 'Avaliado Por'
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'DataAtualizacao'
        ReadOnly = True
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
        Visible = True
      end>
  end
  object ColunasLayout: TStringGrid
    Left = 53
    Top = 160
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
  object StatusBar1: TStatusBar
    Left = 0
    Top = 402
    Width = 806
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
    ExplicitTop = 394
    ExplicitWidth = 804
  end
  object Panel3: TPanel
    Left = 0
    Top = 25
    Width = 806
    Height = 25
    Align = alTop
    TabOrder = 5
    ExplicitWidth = 804
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
      OnChange = actProcurarExecute
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
      OnChange = actProcurarExecute
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
  object DBDataNecessidade: TDBDateTimePicker
    Left = 360
    Top = 264
    Width = 186
    Height = 21
    Date = 46073.000000000000000000
    Time = 0.399926018515543500
    TabOrder = 6
    Visible = False
    Caption = ''
    DataField = 'DataNecessidade'
    DataSource = FrmDataModule.DataSourceMovimentacaoCarga
  end
  object DBDataAtendimento: TDBDateTimePicker
    Left = 360
    Top = 291
    Width = 186
    Height = 21
    Date = 46073.000000000000000000
    Time = 0.399926018515543500
    TabOrder = 7
    Visible = False
    Caption = ''
    DataField = 'DataAtendimento'
    DataSource = FrmDataModule.DataSourceMovimentacaoCarga
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 312
    Top = 152
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
    object actExcel: TAction
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
    end
  end
end
