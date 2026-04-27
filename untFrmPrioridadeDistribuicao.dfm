object FrmPrioridadeDistribuicao: TFrmPrioridadeDistribuicao
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Defini'#231#227'o de prioridade das plataformas na cria'#231#227'o de rotas'
  ClientHeight = 362
  ClientWidth = 624
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  FormStyle = fsStayOnTop
  OnCreate = FormCreate
  TextHeight = 15
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 624
    Height = 25
    Align = alTop
    Caption = 'Defini'#231#227'o de prioridade das plataformas na cria'#231#227'o de rotas'
    Color = clSilver
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 622
  end
  object GridPanel1: TGridPanel
    Left = 0
    Top = 321
    Width = 624
    Height = 41
    Align = alBottom
    ColumnCollection = <
      item
        Value = 50.000000000000000000
      end
      item
        Value = 50.000000000000000000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = btnSalvar
        Row = 0
      end
      item
        Column = 1
        Control = btnFechar
        Row = 0
      end>
    RowCollection = <
      item
        Value = 100.000000000000000000
      end>
    TabOrder = 1
    ExplicitTop = 313
    ExplicitWidth = 622
    DesignSize = (
      624
      41)
    object btnSalvar: TBitBtn
      Left = 119
      Top = 8
      Width = 75
      Height = 25
      Action = actSalvar
      Anchors = []
      Caption = 'Salvar'
      TabOrder = 0
    end
    object btnFechar: TBitBtn
      Left = 430
      Top = 8
      Width = 75
      Height = 25
      Action = actFechar
      Anchors = []
      Caption = 'Fechar'
      TabOrder = 1
    end
  end
  object DBGrid1: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 624
    Height = 267
    Align = alClient
    DataSource = dsPrioridades
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -12
    TitleFont.Name = 'Segoe UI'
    TitleFont.Style = []
    OnDrawColumnCell = DBGrid1DrawColumnCell
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    EnableZebra = False
    ProgressBar = FrmPrincipal.ProgressBarPrincipal
    ExcelButton = btnExcel
    Columns = <
      item
        Expanded = False
        FieldName = 'PrioridadeDistribuicao'
        Title.Alignment = taCenter
        Title.Caption = 'Prioridade'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Plataforma'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'NomeSAP'
        Title.Alignment = taCenter
        Title.Caption = 'Nome SAP'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'RT_Modal'
        Title.Alignment = taCenter
        Title.Caption = 'RT Modal'
        Visible = True
      end>
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 624
    Height = 29
    ButtonHeight = 25
    Caption = 'ToolBar1'
    Images = FrmPrincipal.ImageList1
    TabOrder = 3
    ExplicitWidth = 622
    DesignSize = (
      624
      29)
    object btnCarregar: TBitBtn
      Left = 0
      Top = 0
      Width = 89
      Height = 25
      Action = actCarregar
      Anchors = []
      Caption = 'Carregar'
      TabOrder = 0
    end
    object btnExcel: TToolButton
      Left = 89
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Caption = 'btnExcel'
      ImageIndex = 54
      ParentShowHint = False
      ShowHint = True
    end
    object btnClearFiltro: TToolButton
      Left = 112
      Top = 0
      Hint = 'Limpar filtro'
      Caption = 'btnClearFiltro'
      ImageIndex = 225
      ParentShowHint = False
      ShowHint = True
    end
    object DateTimePickerData: TDateTimePicker
      Left = 135
      Top = 0
      Width = 94
      Height = 25
      Date = 46119.000000000000000000
      Time = 0.756554259256518000
      TabOrder = 1
    end
  end
  object qryPrioridades: TADOQuery
    BeforePost = qryPrioridadesBeforePost
    Parameters = <>
    Left = 416
    Top = 144
  end
  object dsPrioridades: TDataSource
    DataSet = qryPrioridades
    Left = 432
    Top = 224
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 288
    Top = 128
    StyleName = 'Platform Default'
    object actCarregar: TAction
      Caption = 'Carregar'
      Hint = 'Carregar plataformas'
      ImageIndex = 2
      OnExecute = actCarregarExecute
    end
    object actSalvar: TAction
      Caption = 'Salvar'
      Hint = 'Salvar'
      ImageIndex = 18
      OnExecute = actSalvarExecute
    end
    object actFechar: TAction
      Caption = 'Fechar'
      Hint = 'Fechar'
      ImageIndex = 188
      OnExecute = actFecharExecute
    end
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
  end
end

