object FrmConfigRT: TFrmConfigRT
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Configura'#231#227'o - Requisi'#231#227'o de Transporte'
  ClientHeight = 456
  ClientWidth = 790
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  FormStyle = fsStayOnTop
  OnCreate = FormCreate
  TextHeight = 15
  object PageControl1: TPageControl
    Left = 0
    Top = 25
    Width = 790
    Height = 431
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 788
    ExplicitHeight = 423
    object TabSheet1: TTabSheet
      Caption = 'Configura'#231#227'o'
      object GridPanel2: TGridPanel
        Left = 0
        Top = 29
        Width = 782
        Height = 75
        Align = alTop
        ColumnCollection = <
          item
            Value = 27.374595098842590000
          end
          item
            Value = 72.625404901157410000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = Label3
            Row = 0
          end
          item
            Column = 0
            Control = Label4
            Row = 1
          end
          item
            Column = 1
            Control = DBEdit3
            Row = 0
          end
          item
            Column = 1
            Control = DBEdit4
            Row = 1
          end
          item
            Column = 0
            Control = Label5
            Row = 2
          end
          item
            Column = 1
            Control = DBEdit5
            Row = 2
          end>
        RowCollection = <
          item
            Value = 33.631618661353820000
          end
          item
            Value = 33.631618661353820000
          end
          item
            Value = 32.736762677292360000
          end>
        TabOrder = 0
        DesignSize = (
          782
          75)
        object Label3: TLabel
          Left = 74
          Top = 6
          Width = 68
          Height = 15
          Anchors = []
          Caption = 'Requisitante:'
          ExplicitLeft = 81
          ExplicitTop = 9
        end
        object Label4: TLabel
          Left = 57
          Top = 30
          Width = 101
          Height = 15
          Anchors = []
          Caption = 'Pessoa de Contato:'
          ExplicitLeft = 68
          ExplicitTop = 35
        end
        object DBEdit3: TDBEdit
          Left = 215
          Top = 1
          Width = 566
          Height = 25
          Hint = 'Matr'#237'cula do requisitante'
          Align = alClient
          DataField = 'RT_Requisitante'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          TextHint = 'Matr'#237'cula do requisitante'
          ExplicitHeight = 23
        end
        object DBEdit4: TDBEdit
          Left = 215
          Top = 26
          Width = 566
          Height = 24
          Hint = 'Pessoa de contato'
          Align = alClient
          DataField = 'RT_PessoaContato'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          TextHint = 'Pessoa de contato'
          ExplicitHeight = 23
        end
        object Label5: TLabel
          Left = 53
          Top = 54
          Width = 110
          Height = 15
          Anchors = []
          Caption = 'Telefone de Contato:'
          ExplicitLeft = 114
          ExplicitTop = 62
        end
        object DBEdit5: TDBEdit
          Left = 215
          Top = 50
          Width = 566
          Height = 24
          Hint = 'Telefone de contato'
          Align = alClient
          DataField = 'RT_TelContato'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
          TextHint = 'Telefone de contato'
          ExplicitHeight = 23
        end
      end
      object GridPanel1: TGridPanel
        Left = 0
        Top = 104
        Width = 782
        Height = 105
        Align = alTop
        ColumnCollection = <
          item
            Value = 29.439150426530630000
          end
          item
            Value = 21.563016765138840000
          end
          item
            Value = 28.238468715806740000
          end
          item
            Value = 20.759364092523790000
          end>
        ControlCollection = <
          item
            Column = 0
            Control = Label6
            Row = 1
          end
          item
            Column = 1
            Control = DBEdit6
            Row = 1
          end
          item
            Column = 2
            Control = Label8
            Row = 1
          end
          item
            Column = 3
            Control = DBEdit10
            Row = 1
          end
          item
            Column = 2
            Control = Label11
            Row = 2
          end
          item
            Column = 3
            Control = DBEdit11
            Row = 2
          end
          item
            Column = 0
            Control = Label13
            Row = 0
          end
          item
            Column = 2
            Control = Label14
            Row = 0
          end
          item
            Column = 0
            Control = Label15
            Row = 3
          end
          item
            Column = 1
            Control = DBDateTimePicker2
            Row = 3
          end
          item
            Column = 1
            Control = DBEdit12
            Row = 0
          end
          item
            Column = 3
            Control = DBEdit13
            Row = 0
          end
          item
            Column = 1
            Control = DBEdit1
            Row = 2
          end
          item
            Column = 0
            Control = Label1
            Row = 2
          end
          item
            Column = 2
            Control = DBCheckBox1
            Row = 3
          end>
        RowCollection = <
          item
            Value = 25.151879541090610000
          end
          item
            Value = 24.906079538212240000
          end
          item
            Value = 25.089301358579700000
          end
          item
            Value = 24.852739562117450000
          end>
        TabOrder = 1
        DesignSize = (
          782
          105)
        object Label6: TLabel
          Left = 57
          Top = 32
          Width = 118
          Height = 15
          Anchors = []
          Caption = 'Hora Inicio (2'#186' Turno):'
          ExplicitLeft = 51
        end
        object DBEdit6: TDBEdit
          Left = 231
          Top = 27
          Width = 168
          Height = 26
          Hint = 'Centro de planejamento, exemplo 2620'
          Align = alClient
          DataField = 'HoraInicio_Turno2'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          ExplicitHeight = 23
        end
        object Label8: TLabel
          Left = 451
          Top = 32
          Width = 115
          Height = 15
          Anchors = []
          Caption = 'Hora Volta (2'#186' Turno):'
          ExplicitLeft = 325
        end
        object DBEdit10: TDBEdit
          Left = 619
          Top = 27
          Width = 162
          Height = 26
          Hint = 'Grupo de planejamento, exemplo: T31'
          Align = alClient
          DataField = 'HoraVolta_Turno2'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          ExplicitHeight = 23
        end
        object Label11: TLabel
          Left = 482
          Top = 58
          Width = 54
          Height = 15
          Anchors = []
          Caption = 'Grp. Plan.:'
          ExplicitLeft = 329
        end
        object DBEdit11: TDBEdit
          Left = 619
          Top = 53
          Width = 162
          Height = 25
          Align = alClient
          DataField = 'RT_GrpPlan'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
          ExplicitHeight = 23
        end
        object Label13: TLabel
          Left = 57
          Top = 6
          Width = 118
          Height = 15
          Anchors = []
          Caption = 'Hora Inicio (1'#186' Turno):'
          ExplicitLeft = 1
        end
        object Label14: TLabel
          Left = 451
          Top = 6
          Width = 115
          Height = 15
          Anchors = []
          Caption = 'Hora Volta (1'#186' Turno):'
          ExplicitLeft = 248
        end
        object Label15: TLabel
          Left = 73
          Top = 83
          Width = 85
          Height = 15
          Anchors = []
          Caption = 'Hora Cobertura:'
          ExplicitLeft = 39
        end
        object DBDateTimePicker2: TDBDateTimePicker
          Left = 231
          Top = 78
          Width = 168
          Height = 26
          Align = alClient
          Date = 46075.000000000000000000
          Time = 0.418912650464335500
          Kind = dtkTime
          TabOrder = 3
          Caption = ''
          DataField = 'RT_HoraCobertura'
          DataSource = FrmDataModule.DataSourceConfigRT
        end
        object DBEdit12: TDBEdit
          Left = 231
          Top = 1
          Width = 168
          Height = 26
          Hint = 'Este '#233' o hor'#225'rio padr'#227'o considerado de inicio'
          Align = alClient
          DataField = 'HoraInicio_Turno1'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 4
          ExplicitHeight = 23
        end
        object DBEdit13: TDBEdit
          Left = 619
          Top = 1
          Width = 162
          Height = 26
          Hint = 'Caso seja bate e volta, este '#233' o hor'#225'rio padr'#227'o de volta'
          Align = alClient
          DataField = 'HoraVolta_Turno1'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 5
          ExplicitHeight = 23
        end
        object DBEdit1: TDBEdit
          Left = 231
          Top = 53
          Width = 168
          Height = 25
          Hint = 'Centro de planejamento, exemplo 2620'
          Align = alClient
          DataField = 'RT_CentroPlan'
          DataSource = FrmDataModule.DataSourceConfigRT
          ParentShowHint = False
          ShowHint = True
          TabOrder = 6
          ExplicitHeight = 23
        end
        object Label1: TLabel
          Left = 96
          Top = 58
          Width = 39
          Height = 15
          Anchors = []
          Caption = 'Centro:'
          ExplicitLeft = 52
        end
        object DBCheckBox1: TDBCheckBox
          Left = 460
          Top = 82
          Width = 97
          Height = 17
          Anchors = []
          Caption = 'Salvar'
          DataField = 'RT_Salvar'
          DataSource = FrmDataModule.DataSourceConfigRT
          TabOrder = 7
        end
      end
      object ToolBar1: TToolBar
        Left = 0
        Top = 0
        Width = 782
        Height = 29
        Caption = 'ToolBar1'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        TabOrder = 2
        object DBNavigator1: TDBNavigator
          Left = 0
          Top = 0
          Width = 128
          Height = 22
          DataSource = FrmDataModule.DataSourceConfigRT
          VisibleButtons = [nbEdit, nbPost, nbCancel, nbRefresh]
          Hints.Strings = (
            'First record'
            'Prior record'
            'Next record'
            'Last record'
            'Insert record'
            'Delete record'
            'Editar'
            'Gravar'
            'Cancelar'
            'Atualizar'
            'Aplicar'
            'Cancelar')
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Regras de recolhimento'
      ImageIndex = 1
      object ToolBar3: TToolBar
        Left = 0
        Top = 0
        Width = 782
        Height = 29
        ButtonHeight = 23
        Caption = 'ToolBar3'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 0
        ExplicitWidth = 780
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
      end
      object FilterDBGrid1: TFilterDBGrid
        Left = 0
        Top = 29
        Width = 782
        Height = 353
        Align = alClient
        DataSource = FrmDataModule.DataSourceRegrasRecolhimento
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Segoe UI'
        TitleFont.Style = []
        ClearFilterButton = btnClearFiltroRT
        SearchAction = actProcurar
        LayoutGrid = RLLayout
        EnableZebra = False
        LayoutButton = btnLaypoutRT
        Columns = <
          item
            Expanded = False
            FieldName = 'idRegraRecolhimento'
            Title.Alignment = taCenter
            Title.Caption = 'ID'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Ativa'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Descricao'
            Title.Alignment = taCenter
            Title.Caption = 'Descri'#231#227'o da Regra'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Origem'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'txtDestino'
            Title.Alignment = taCenter
            Title.Caption = 'Destino'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DestinoNaoCriarRT'
            Title.Alignment = taCenter
            Title.Caption = 'Destino N'#227'o Criar RT'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RecolherParaTipo'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo Recolher Para'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RecolherParaValor'
            Title.Alignment = taCenter
            Title.Caption = 'Valor Recolher Para'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Observacao'
            Title.Alignment = taCenter
            Title.Caption = 'Observa'#231#227'o'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'Prioridade'
            Title.Alignment = taCenter
            Visible = True
          end>
      end
      object StatusBarRegrasRecolhimento: TStatusBar
        Left = 0
        Top = 382
        Width = 782
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
        ExplicitTop = 374
        ExplicitWidth = 780
      end
      object RLLayout: TStringGrid
        Left = 384
        Top = 168
        Width = 97
        Height = 120
        ColCount = 7
        RowCount = 10
        FixedRows = 0
        TabOrder = 3
        Visible = False
      end
    end
  end
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 790
    Height = 25
    Align = alTop
    Caption = 'Configura'#231#227'o - Requisi'#231#227'o de Transporte'
    Color = clSilver
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 1
    ExplicitWidth = 788
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 276
    Top = 187
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
  end
end
