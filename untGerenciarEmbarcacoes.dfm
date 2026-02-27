object FrmGerenciarEmbarcacoes: TFrmGerenciarEmbarcacoes
  Left = 0
  Top = 0
  Caption = 'Gerenciar Transporte Mar'#237'timo'
  ClientHeight = 557
  ClientWidth = 1282
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
    Width = 1282
    Height = 25
    Align = alTop
    Caption = 'Gerenciar Transporte Mar'#237'timo'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 1280
  end
  object PageControlPrincipal: TPageControl
    Left = 0
    Top = 50
    Width = 1282
    Height = 507
    ActivePage = TabSheet3
    Align = alClient
    TabOrder = 1
    OnChange = PageControlPrincipalChange
    ExplicitWidth = 1280
    ExplicitHeight = 499
    object TabSheet3: TTabSheet
      Caption = 'Distribui'#231#227'o'
      ImageIndex = 2
      object Splitter2: TSplitter
        Left = 460
        Top = 0
        Height = 479
        ExplicitLeft = 456
        ExplicitTop = 264
        ExplicitHeight = 100
      end
      object PanelLateral: TPanel
        Left = 0
        Top = 0
        Width = 460
        Height = 479
        Align = alLeft
        TabOrder = 0
        ExplicitHeight = 471
        object Splitter1: TSplitter
          Left = 1
          Top = 286
          Width = 458
          Height = 3
          Cursor = crVSplit
          Align = alTop
          ExplicitTop = 249
          ExplicitWidth = 438
        end
        object PanelCadastro: TPanel
          Left = 1
          Top = 1
          Width = 458
          Height = 285
          Align = alTop
          TabOrder = 0
          object DBGridRoteamento: TFilterDBGrid
            Left = 1
            Top = 55
            Width = 456
            Height = 210
            Align = alClient
            DataSource = FrmDataModule.DataSourceRoteamento
            Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
            TabOrder = 0
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
            OnDrawColumnCell = DBGridRoteamentoDrawColumnCell
            OnKeyDown = DBGridRoteamentoKeyDown
            OnTitleClick = DBGridRoteamentoTitleClick
            ClearFilterButton = btnFiltroClearRoteamento
            SearchAction = actProcuraRoteamento
            LayoutGrid = ColunasLayoutRoteamento
            EnableZebra = False
            Columns = <
              item
                Expanded = False
                FieldName = 'NomeRota'
                Title.Alignment = taCenter
                Title.Caption = 'Descri'#231#227'o da Rota'
                Width = 102
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'NomeEmbarcacao'
                Title.Alignment = taCenter
                Title.Caption = 'Embarca'#231#227'o'
                Width = 100
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'HoraRoteamento'
                ReadOnly = True
                Title.Alignment = taCenter
                Title.Caption = 'Hora 1'#176' Origem'
                Width = 83
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'CapacidadePAX'
                Title.Alignment = taCenter
                Title.Caption = 'Capacidade (PAX)'
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
                FieldName = 'linhaCor'
                Title.Alignment = taCenter
                Title.Caption = 'Cor da linha'
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
          object ColorBoxLinha: TColorBox
            Left = 27
            Top = 183
            Width = 82
            Height = 22
            TabOrder = 1
            Visible = False
            OnCloseUp = ColorBoxLinhaCloseUp
            OnMouseLeave = ColorBoxLinhaMouseLeave
          end
          object ToolBar3: TToolBar
            Left = 1
            Top = 26
            Width = 456
            Height = 29
            Caption = 'ToolBar3'
            DrawingStyle = dsGradient
            EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
            Images = FrmPrincipal.ImageList1
            TabOrder = 2
            object btnFiltroClearRoteamento: TToolButton
              Left = 0
              Top = 0
              Hint = 'Limpar filtro'
              Caption = 'btnFiltroClearRoteamento'
              ImageIndex = 225
              ParentShowHint = False
              ShowHint = True
            end
            object btnInserir: TBitBtn
              Left = 23
              Top = 0
              Width = 70
              Height = 22
              Action = actInserirRoteamento
              Caption = 'Inserir'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 0
            end
            object btnSalvar: TBitBtn
              Left = 93
              Top = 0
              Width = 70
              Height = 22
              Action = actEditarRoteamento
              Caption = 'Salvar'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 2
            end
            object btnExcluirRota: TBitBtn
              Left = 163
              Top = 0
              Width = 90
              Height = 22
              Action = actExcluirRoteamento
              Caption = 'Excluir Rota'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 1
            end
            object DBEditCodigoRoteamento: TDBEdit
              Left = 253
              Top = 0
              Width = 48
              Height = 22
              DataField = 'idRoteamento'
              DataSource = FrmDataModule.DataSourceRoteamento
              TabOrder = 3
              Visible = False
              OnChange = actProcuraRotaExecutantesExecute
            end
          end
          object PanelTituloCdastro: TPanel
            Left = 1
            Top = 1
            Width = 456
            Height = 25
            Align = alTop
            Caption = 'Cadastro de Rotas'
            Color = clMenuBar
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentBackground = False
            ParentFont = False
            TabOrder = 3
          end
          object MaskEditHora: TMaskEdit
            Left = 27
            Top = 133
            Width = 79
            Height = 21
            Hint = 'Hor'#225'rio de sa'#237'da da embarca'#231#227'o na origem'
            EditMask = '!90:00;1;_'
            MaxLength = 5
            ParentShowHint = False
            ShowHint = True
            TabOrder = 4
            Text = '06:00'
            Visible = False
            OnChange = MaskEditHoraChange
          end
          object ColunasLayoutRoteamento: TStringGrid
            Left = 168
            Top = 129
            Width = 125
            Height = 76
            ColCount = 7
            DefaultRowHeight = 21
            RowCount = 10
            TabOrder = 5
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
              21)
          end
          object StatusBarRoteamento: TStatusBar
            Left = 1
            Top = 265
            Width = 456
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
        end
        object PanelPltaformas: TPanel
          Left = 1
          Top = 289
          Width = 458
          Height = 189
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 181
          object Splitter4: TSplitter
            Left = 209
            Top = 26
            Height = 162
            ExplicitLeft = 194
            ExplicitTop = 32
            ExplicitHeight = 353
          end
          object PanelTituloPltaformas: TPanel
            Left = 1
            Top = 1
            Width = 456
            Height = 25
            Align = alTop
            Caption = 'Roteiro da Embarca'#231#227'o'
            Color = clMenuBar
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentBackground = False
            ParentFont = False
            TabOrder = 0
          end
          object Panel4: TPanel
            Left = 1
            Top = 26
            Width = 208
            Height = 162
            Align = alLeft
            Caption = 'PanelSelecaoPlataforma'
            TabOrder = 1
            ExplicitHeight = 154
            object ToolBar4: TToolBar
              Left = 1
              Top = 1
              Width = 206
              Height = 29
              ButtonHeight = 23
              Caption = 'ToolBar4'
              DrawingStyle = dsGradient
              EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
              HotTrackColor = clYellow
              Images = FrmPrincipal.ImageList1
              TabOrder = 0
              object ToolButton1: TToolButton
                Left = 0
                Top = 0
                Action = actExcelDestinos
                ParentShowHint = False
                ShowHint = True
              end
              object btnPendentes: TToolButton
                Left = 23
                Top = 0
                Hint = 'Visualizar somente os destinos pendentes'
                Caption = 'Pendentes'
                Down = True
                ImageIndex = 459
                ParentShowHint = False
                ShowHint = True
                Style = tbsCheck
                OnClick = actFiltroDestinosExecute
              end
              object btnTotal: TToolButton
                Left = 46
                Top = 0
                Hint = 'Total de passageiros (Pendentes ou N'#227'o)'
                Caption = 'btnTotal'
                Down = True
                ImageIndex = 458
                ParentShowHint = False
                ShowHint = True
                Style = tbsCheck
                OnClick = actFiltroDestinosExecute
              end
              object ToolButton2: TToolButton
                Left = 69
                Top = 0
                Hint = 'Atualizar'
                Caption = 'Atualizar'
                ImageIndex = 238
                ParentShowHint = False
                ShowHint = True
                OnClick = actFiltroDestinosExecute
              end
              object ToolButton6: TToolButton
                Left = 92
                Top = 0
                Action = actImprimirExtrato
                ParentShowHint = False
                ShowHint = True
              end
            end
            object StrGridResumo: TStringGrid
              Left = 1
              Top = 30
              Width = 206
              Height = 131
              Align = alClient
              FixedCols = 0
              TabOrder = 1
              OnDblClick = StrGridResumoDblClick
              ExplicitHeight = 123
            end
          end
          object Panel5: TPanel
            Left = 212
            Top = 26
            Width = 245
            Height = 162
            Align = alClient
            Caption = 'Panel5'
            TabOrder = 2
            ExplicitHeight = 154
            object ToolBar1: TToolBar
              Left = 1
              Top = 1
              Width = 243
              Height = 29
              Caption = 'ToolBar1'
              DrawingStyle = dsGradient
              EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
              Images = FrmPrincipal.ImageList1
              TabOrder = 0
              object BitBtn9: TBitBtn
                Left = 0
                Top = 0
                Width = 77
                Height = 22
                Action = actCalcular
                Caption = 'Calcular'
                ParentShowHint = False
                ShowHint = True
                TabOrder = 0
              end
            end
            object StrGridRota: TStringGrid
              Left = 1
              Top = 30
              Width = 243
              Height = 131
              Align = alClient
              ColCount = 6
              DefaultColWidth = 20
              DefaultRowHeight = 21
              RowCount = 2
              Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goRowMoving]
              PopupMenu = PopupMenuRota
              TabOrder = 1
              OnRowMoved = StrGridRotaRowMoved
              OnSelectCell = StrGridRotaSelectCell
              ExplicitHeight = 123
              ColWidths = (
                20
                59
                42
                51
                54
                20)
            end
            object ComboBoxPlataforma: TComboBox
              Left = 88
              Top = 114
              Width = 100
              Height = 22
              AutoDropDown = True
              BevelInner = bvSpace
              BevelKind = bkFlat
              Style = csOwnerDrawFixed
              CharCase = ecUpperCase
              DragMode = dmAutomatic
              DropDownCount = 10
              ParentShowHint = False
              ShowHint = True
              TabOrder = 2
              Visible = False
              OnCloseUp = ComboBoxPlataformaCloseUp
              OnMouseLeave = ComboBoxPlataformaMouseLeave
            end
            object ComboBoxOrigemInicial: TComboBox
              Left = 88
              Top = 86
              Width = 100
              Height = 22
              AutoDropDown = True
              BevelInner = bvSpace
              BevelKind = bkFlat
              Style = csOwnerDrawFixed
              CharCase = ecUpperCase
              DragMode = dmAutomatic
              DropDownCount = 10
              ParentShowHint = False
              ShowHint = True
              TabOrder = 3
              Visible = False
              OnCloseUp = ComboBoxOrigemInicialCloseUp
              OnMouseLeave = ComboBoxOrigemInicialMouseLeave
              Items.Strings = (
                'TMIB'
                'PCM-09')
            end
          end
        end
      end
      object PanelVinculados: TPanel
        Left = 463
        Top = 0
        Width = 811
        Height = 479
        Align = alClient
        TabOrder = 1
        ExplicitWidth = 809
        ExplicitHeight = 471
        object SplitterAnalise: TSplitter
          Left = 1
          Top = 329
          Width = 809
          Height = 3
          Cursor = crVSplit
          Align = alTop
          ExplicitTop = 285
          ExplicitWidth = 988
        end
        object PanelExecutantesVinculados: TPanel
          Left = 1
          Top = 332
          Width = 809
          Height = 146
          Align = alClient
          TabOrder = 0
          ExplicitWidth = 807
          ExplicitHeight = 138
          object ToolBar10: TToolBar
            Left = 1
            Top = 26
            Width = 807
            Height = 29
            ButtonHeight = 23
            Caption = 'ToolBar10'
            DrawingStyle = dsGradient
            EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
            Images = FrmPrincipal.ImageList1
            TabOrder = 0
            ExplicitWidth = 805
            object ToolButton7: TToolButton
              Left = 0
              Top = 0
              Action = actSelRotaExecTODOS
              ParentShowHint = False
              ShowHint = True
            end
            object ToolButton13: TToolButton
              Left = 23
              Top = 0
              Action = actSelRotaExecLIMPAR
              ParentShowHint = False
              ShowHint = True
            end
            object btnExcluirVinculos: TToolButton
              Left = 46
              Top = 0
              Hint = 'Excluir executantes vinculados'
              Caption = 'btnExcluirVinculos'
              DropdownMenu = PopupMenuExcluirVinculados
              ImageIndex = 324
              Style = tbsDropDown
            end
            object btnFiltroClearRotaExecutante: TToolButton
              Left = 88
              Top = 0
              Hint = 'Limpar filtro'
              Caption = 'btnFiltroClearRotaExecutante'
              ImageIndex = 225
              ParentShowHint = False
              ShowHint = True
            end
            object BitBtn10: TBitBtn
              Left = 111
              Top = 0
              Width = 70
              Height = 23
              Action = actExcel
              Caption = 'Exportar'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 0
            end
          end
          object DBGridRotaExecutantes: TFilterDBGrid
            Left = 1
            Top = 55
            Width = 807
            Height = 71
            Align = alClient
            DataSource = FrmDataModule.DataSourceTM_RotaExecutantes
            ReadOnly = True
            TabOrder = 1
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
            OnCellClick = DBGridRotaExecutantesCellClick
            OnDrawColumnCell = DBGridRotaExecutantesDrawColumnCell
            OnKeyDown = DBGridRotaExecutantesKeyDown
            ClearFilterButton = btnFiltroClearRotaExecutante
            SearchAction = actProcuraRotaExecutantes
            LayoutGrid = ColunasLayoutRotaExecutantes
            EnableZebra = False
            Columns = <
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'booleanSelecao'
                Title.Alignment = taCenter
                Title.Caption = 'Sele'#231#227'o'
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'Origem'
                Title.Alignment = taCenter
                Width = 70
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'txtDestino'
                Title.Alignment = taCenter
                Title.Caption = 'Destino'
                Width = 70
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'NomeExecutante'
                Title.Alignment = taCenter
                Title.Caption = 'Nome do Executante'
                Width = 210
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'txtTipoEtapaServico'
                Title.Alignment = taCenter
                Title.Caption = 'Tipo de Etapa de Servico'
                Width = 137
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'Empresa'
                Title.Alignment = taCenter
                Width = 152
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'NomeEmbarcacao'
                Title.Alignment = taCenter
                Title.Caption = 'Embarca'#231#227'o'
                Width = 111
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'HoraRoteamento'
                Title.Alignment = taCenter
                Title.Caption = 'Hora 1'#176' Origem'
                Visible = True
              end>
          end
          object StatusBarRotaExecutantes: TStatusBar
            Left = 1
            Top = 126
            Width = 807
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
            ExplicitTop = 118
            ExplicitWidth = 805
          end
          object ColunasLayoutRotaExecutantes: TStringGrid
            Left = 72
            Top = 63
            Width = 113
            Height = 65
            ColCount = 7
            DefaultRowHeight = 21
            RowCount = 8
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
              21)
          end
          object PanelTituloRotaExecutantes: TPanel
            Left = 1
            Top = 1
            Width = 807
            Height = 25
            Align = alTop
            Color = clGradientActiveCaption
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentBackground = False
            ParentFont = False
            TabOrder = 4
            ExplicitWidth = 805
          end
        end
        object PanelAnalise: TPanel
          Left = 1
          Top = 26
          Width = 809
          Height = 303
          Align = alTop
          TabOrder = 1
          ExplicitWidth = 807
          object StatusBarExecutantes: TStatusBar
            Left = 1
            Top = 283
            Width = 807
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
            ExplicitWidth = 805
          end
          object ToolBar9: TToolBar
            Left = 1
            Top = 1
            Width = 807
            Height = 29
            ButtonHeight = 23
            Caption = 'ToolBar8'
            DrawingStyle = dsGradient
            EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
            HotTrackColor = clRed
            Images = FrmPrincipal.ImageList1
            TabOrder = 1
            ExplicitWidth = 805
            object ToolButton21: TToolButton
              Left = 0
              Top = 0
              Action = actSelProgExec1
              ParentShowHint = False
              ShowHint = True
            end
            object ToolButton24: TToolButton
              Left = 23
              Top = 0
              Action = actSelProgExec2
              ParentShowHint = False
              ShowHint = True
            end
            object btnVincluados: TToolButton
              Left = 46
              Top = 0
              Action = actMostrarVinculados
              Down = True
              ParentShowHint = False
              ShowHint = True
              Style = tbsCheck
            end
            object btnProgramados: TToolButton
              Left = 69
              Top = 0
              Hint = 'Mostrar lista completa de executantes programados'
              Caption = 'btnProgramados'
              ImageIndex = 8
              ParentShowHint = False
              ShowHint = True
              Style = tbsCheck
              OnClick = btnProgramadosClick
            end
            object btnDisponivel: TToolButton
              Left = 92
              Top = 0
              Action = actDisponivel
              Down = True
              ParentShowHint = False
              ShowHint = True
              Style = tbsCheck
            end
            object btnFiltroClearProcuraGeral: TToolButton
              Left = 115
              Top = 0
              Hint = 'Limpar filtro'
              Caption = 'btnFiltroClearProcuraGeral'
              ImageIndex = 225
              ParentShowHint = False
              ShowHint = True
            end
            object btnFiltroClearProgramados: TToolButton
              Left = 138
              Top = 0
              Hint = 'Limpar filtro'
              ImageIndex = 225
              ParentShowHint = False
              ShowHint = True
              Visible = False
            end
            object BitBtn13: TBitBtn
              Left = 161
              Top = 0
              Width = 70
              Height = 23
              Action = actExcelProcuraExecutantes
              Caption = 'Exportar'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 0
            end
            object BitBtn20: TBitBtn
              Left = 231
              Top = 0
              Width = 93
              Height = 23
              Action = actVincular
              Caption = 'Vincular Exec.'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 2
            end
            object BitBtn21: TBitBtn
              Left = 324
              Top = 0
              Width = 93
              Height = 23
              Action = actTransbordoInterno
              Caption = 'Transbordos'
              ParentShowHint = False
              ShowHint = True
              TabOrder = 3
            end
            object CheckBoxOrigemDestino: TCheckBox
              Left = 417
              Top = 0
              Width = 202
              Height = 23
              Action = actProcuraExecutantes
              Caption = 'DESTINO diferente da ORIGEM'
              Checked = True
              State = cbChecked
              TabOrder = 1
            end
          end
          object DBGridProcuraGeral: TFilterDBGrid
            Left = 1
            Top = 30
            Width = 807
            Height = 115
            Align = alTop
            DataSource = FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao
            Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
            ReadOnly = True
            TabOrder = 2
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
            OnCellClick = DBGridProcuraGeralCellClick
            OnDrawColumnCell = DBGridProcuraGeralDrawColumnCell
            ClearFilterButton = btnFiltroClearProcuraGeral
            SearchAction = actProcuraGeral
            LayoutGrid = ColunasLayoutProcuraGeral
            EnableZebra = False
            Columns = <
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'booleanSelecao'
                Title.Alignment = taCenter
                Title.Caption = 'Sele'#231#227'o'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'InseridoProgramacaoTransporte'
                ReadOnly = False
                Title.Alignment = taCenter
                Title.Caption = 'TM'
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'StatusProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Status [Programa'#231#227'o]'
                Width = 109
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
                Width = 177
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
                Alignment = taCenter
                Expanded = False
                FieldName = 'HorarioChegada'
                Title.Alignment = taCenter
                Title.Caption = 'Hora Chegada'
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
                FieldName = 'RT'
                Title.Alignment = taCenter
                Width = 105
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'TipoEmbarque'
                Title.Alignment = taCenter
                Title.Caption = 'Tipo de Embarque'
                Width = 108
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'RequisitantePT'
                Title.Alignment = taCenter
                Title.Caption = 'Req. PT'
                Width = 56
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'CriadoPor'
                Title.Alignment = taCenter
                Title.Caption = 'Criado Por'
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
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'AvaliadoPorProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Avalia'#231#227'o [Programa'#231#227'o]'
                Width = 81
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'DataAvaliacaoProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Data [Programa'#231#227'o]'
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'ComputadorProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Computador [Programa'#231#227'o]'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'MotivoProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Motivo [Programa'#231#227'o]'
                Width = 156
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'PalavraChaveProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Palavras Chave [Programa'#231#227'o]'
                Width = 300
                Visible = True
              end>
          end
          object DBGridProcuraProgramados: TFilterDBGrid
            Left = 1
            Top = 145
            Width = 807
            Height = 138
            Align = alClient
            DataSource = FrmDataModule.DataSourceProgramados
            ReadOnly = True
            TabOrder = 3
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
            OnDrawColumnCell = DBGridProcuraProgramadosDrawColumnCell
            ClearFilterButton = btnFiltroClearProgramados
            SearchAction = actProcuraProgramados
            LayoutGrid = ColunasLayoutProgramados
            EnableZebra = False
            Columns = <
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'StatusProgramacao'
                Title.Alignment = taCenter
                Title.Caption = 'Status [Programa'#231#227'o]'
                Width = 109
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
                Expanded = False
                FieldName = 'NomeRota'
                Title.Alignment = taCenter
                Title.Caption = 'Descri'#231#227'o da Rota'
                Width = 101
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'HoraRoteamento'
                Title.Alignment = taCenter
                Title.Caption = 'Hora 1'#176' Origem'
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'NomeEmbarcacao'
                Title.Alignment = taCenter
                Title.Caption = 'Embarca'#231#227'o'
                Width = 100
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
                Width = 177
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
                Alignment = taCenter
                Expanded = False
                FieldName = 'HorarioChegada'
                Title.Alignment = taCenter
                Title.Caption = 'Hora Chegada'
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
                FieldName = 'RT'
                Title.Alignment = taCenter
                Width = 105
                Visible = True
              end
              item
                Alignment = taCenter
                Expanded = False
                FieldName = 'TipoEmbarque'
                Title.Alignment = taCenter
                Title.Caption = 'Tipo de Embarque'
                Width = 108
                Visible = True
              end
              item
                Expanded = False
                FieldName = 'RequisitantePT'
                Title.Alignment = taCenter
                Title.Caption = 'Req. PT'
                Width = 56
                Visible = True
              end>
          end
          object ColunasLayoutProcuraGeral: TStringGrid
            Left = 288
            Top = 48
            Width = 113
            Height = 65
            ColCount = 7
            DefaultRowHeight = 21
            RowCount = 23
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
              21)
          end
          object ColunasLayoutProgramados: TStringGrid
            Left = 288
            Top = 198
            Width = 113
            Height = 65
            ColCount = 7
            DefaultRowHeight = 21
            RowCount = 16
            TabOrder = 5
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
              21)
          end
        end
        object PanelTituloExecutantes: TPanel
          Left = 1
          Top = 1
          Width = 809
          Height = 25
          Align = alTop
          Caption = 'Sele'#231#227'o de Executantes para vinculo ao roteamento'
          Color = clMenuBar
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 2
          ExplicitWidth = 807
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Relat'#243'rios'
      ImageIndex = 3
      object StatusBarImpressao: TStatusBar
        Left = 0
        Top = 460
        Width = 1274
        Height = 19
        Panels = <
          item
            Width = 50
          end>
      end
      object PanelAjuda: TPanel
        Left = 474
        Top = 125
        Width = 337
        Height = 220
        TabOrder = 1
        Visible = False
        object PanelTituloAjuda: TPanel
          Left = 1
          Top = 1
          Width = 335
          Height = 26
          Align = alTop
          Caption = 'Inserir Servi'#231'o'
          Color = clGradientActiveCaption
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -12
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 0
          OnMouseDown = PanelTituloAjudaMouseDown
          OnMouseMove = PanelTituloAjudaMouseMove
          OnMouseUp = PanelTituloAjudaMouseUp
          object SpeedButton4: TSpeedButton
            Left = 311
            Top = 1
            Width = 23
            Height = 24
            Hint = 'Cancelar opera'#231#227'o'
            Align = alRight
            Glyph.Data = {
              36040000424D3604000000000000360000002800000010000000100000000100
              2000000000000004000000000000000000000000000000000000FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
              FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00181A
              65003436CB003638DB00363ADF00363AE3003639E7003538EB003537EF00393D
              F0003C41F2004046F200454DF0004046CC000B0C3500FF00FF0003032F001717
              C2000000BD000000C4000000CB000000D2000202D9001B1BE3001718E7000E12
              E8001820EA00212CEC002B39ED003344EF003A45D1000000040003034F000000
              B5000000BC000000C3000E0ECD008D8DEB00EDEDFC00FFFFFF00FFFFFF00DEDF
              FC00797DF300232EEC002B39ED003444EF00394BEA000000150000004A000000
              B3000000BB000F0FC500D0D0F500FFFFFF00B4B4EE006D6DE6007878EB00D2D2
              F700FFFFFF00ACB0F8002A37ED003343EF003647E70000001400000045000A0A
              B2000E0EBC009999E500FFFFFF007272DE000303D4000000DC000000E2000B0D
              E400ABADF200FFFFFF006F78F300303FEE003242E30000001200000040001919
              B2002424BF00F2F2FB00D3D3F2002B2BD2002E2ED9002525DD001414E1000405
              E5001D21E600F9F9FE00BDC1F9002A38ED002D3BDF000000120000003C002626
              B2004141C400FFFFFF00B0B0EB003B3BD2003E3ED8004040DE004343E3003E3E
              E8001F20E800DBDBFB00D8D9FB00232DEC002531DC0000001200000037003232
              B1004444C100F8F8FC00D0D0F3004A4AD2004D4DD700C0C0F3009292ED005656
              E6005F5FEC00F8F8FE00BFC1F700181FEA001B24D90000001200000034003F3F
              B1005151C100C1C1E600FEFEFF008B8BE0005C5CD600FFFFFF00CACAF5006565
              E400BABAF500FFFFFF009797EF001215E8000E12D50000001200000033004B4B
              B2006060C2007373C900F0F0F700FCFCFE008282DC00FFFFFF00CFCFF400BEBE
              F100FFFFFF00DCDCF3007C7CE9006363EB000505D00000001200000032005B5B
              B5006F6FC3007171C8008484CC00B6B6DD007D7DD400FFFFFF00D3D3F300B2B2
              E400C6C6E7008A8AE1008989E7008C8CE9004747D00000001100000031007474
              BD007C7CC4007F7FC9008181CD008484D1008787D400FFFFFF00D8D8F3008F8F
              DE009292E0009595E2009898E3009A9AE4007171D00001011200101034008989
              C3008B8BC6008D8DCB009090CE009393D2009696D400C5C5E000B2B2DD009E9E
              DC00A0A0DE00A3A3E000A6A6E000A8A8E1009F9FD90000000C00000007007B7B
              AA00A4A4CF00A3A3D100A6A6D400A9A9D700ACACD900AFAFDC00B2B2DE00B5B5
              E000B7B7E200BABAE300BDBDE300C9C9E8006C6C8C00FF00FF00FF00FF000303
              0B004040590056566C0058586D005A5A6E005C5C6F005E5E70005F5F71006161
              720062627200636373006363730039394C0000000200FF00FF00}
            ParentShowHint = False
            ShowHint = True
            OnClick = SpeedButton4Click
            ExplicitLeft = 0
            ExplicitTop = -4
            ExplicitHeight = 39
          end
        end
        object PanelAjudaImpressao: TPanel
          Left = 1
          Top = 27
          Width = 335
          Height = 150
          Align = alTop
          TabOrder = 1
          object CheckListBoxImpressao: TCheckListBox
            Left = 1
            Top = 30
            Width = 333
            Height = 119
            Align = alClient
            ItemHeight = 17
            TabOrder = 0
          end
          object ToolBar6: TToolBar
            Left = 1
            Top = 1
            Width = 333
            Height = 29
            Caption = 'ToolBar6'
            DrawingStyle = dsGradient
            EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
            Images = FrmPrincipal.ImageList1
            TabOrder = 1
            object btnImprimirRelatorio: TToolButton
              Left = 0
              Top = 0
              Hint = 'Imprimir relat'#243'rio'
              Caption = 'actImprimirAjuda'
              ImageIndex = 21
            end
          end
        end
      end
      object Memo1: TMemo
        Left = 0
        Top = 371
        Width = 1274
        Height = 89
        Align = alBottom
        ScrollBars = ssBoth
        TabOrder = 2
      end
      object Panel11: TPanel
        Left = 0
        Top = 351
        Width = 1274
        Height = 20
        Align = alBottom
        Alignment = taLeftJustify
        Caption = '  Notas:'
        Color = clGradientActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 3
      end
      object ToolBar7: TToolBar
        Left = 0
        Top = 0
        Width = 1274
        Height = 29
        Caption = 'ToolBar7'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 4
        object BitBtn2: TBitBtn
          Left = 0
          Top = 0
          Width = 113
          Height = 22
          Action = actRLImprimir_Rota
          Align = alLeft
          Caption = 'Rota Selecionada'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
        end
        object BitBtn12: TBitBtn
          Left = 113
          Top = 0
          Width = 96
          Height = 22
          Action = actListaRTEmbarque
          Align = alLeft
          Caption = 'RT Embarque'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 8
        end
        object BitBtn7: TBitBtn
          Left = 209
          Top = 0
          Width = 71
          Height = 22
          Action = actLanche
          Align = alLeft
          Caption = 'Lanches'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
        end
        object BitBtn8: TBitBtn
          Left = 280
          Top = 0
          Width = 97
          Height = 22
          Action = actKIT_PS
          Align = alLeft
          Caption = 'KIT Embarque'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
        object BitBtn5: TBitBtn
          Left = 377
          Top = 0
          Width = 88
          Height = 22
          Action = actDuplicados
          Align = alLeft
          Caption = 'Transbordos'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
        object BitBtn3: TBitBtn
          Left = 465
          Top = 0
          Width = 67
          Height = 22
          Action = actImprimirPlanilha
          Align = alLeft
          Caption = 'Imprmir'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 5
        end
        object BitBtn1: TBitBtn
          Left = 532
          Top = 0
          Width = 62
          Height = 22
          Action = actImprimirRotas
          Align = alLeft
          Caption = 'Rotas'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 9
        end
        object BitBtn23: TBitBtn
          Left = 594
          Top = 0
          Width = 76
          Height = 22
          Action = actExcelImpressao
          Align = alLeft
          Caption = 'Exportar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 10
        end
        object BitBtn16: TBitBtn
          Left = 670
          Top = 0
          Width = 67
          Height = 22
          Action = actClassificarRegistros
          Caption = 'Numerar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 4
        end
        object BitBtn14: TBitBtn
          Left = 737
          Top = 0
          Width = 64
          Height = 22
          Action = actAjustar
          Caption = 'Ajustar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 7
        end
        object SpinEditFonte: TSpinEdit
          Left = 801
          Top = 0
          Width = 56
          Height = 22
          Hint = 'Alterar a fonte do relat'#243'rio'
          MaxValue = 14
          MinValue = 6
          ParentShowHint = False
          ShowHint = True
          TabOrder = 6
          Value = 9
          OnChange = SpinEditFonteChange
        end
      end
      object RLImpressao: TStringGrid
        Left = 0
        Top = 81
        Width = 1274
        Height = 270
        Align = alClient
        RowCount = 2
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goDrawFocusSelected, goColSizing, goRowMoving, goColMoving, goEditing, goFixedRowClick]
        PopupMenu = PopupMenuImpressao
        TabOrder = 5
        OnFixedCellClick = RLImpressaoFixedCellClick
        ColWidths = (
          64
          64
          64
          64
          64)
      end
      object RadioGroupImpressao: TRadioGroup
        Left = 168
        Top = 152
        Width = 185
        Height = 105
        Caption = 'RadioGroupImpressao'
        Items.Strings = (
          'Rota Selecionada'
          'RT - TIMB'
          'Lanches'
          'KIT (PS)')
        TabOrder = 6
        Visible = False
      end
      object PanelTitulo1: TPanel
        Left = 0
        Top = 55
        Width = 1274
        Height = 26
        Align = alTop
        TabOrder = 7
        object edtTitulo2: TEdit
          Left = 1
          Top = 1
          Width = 1272
          Height = 24
          Align = alClient
          Alignment = taCenter
          Color = clGradientActiveCaption
          TabOrder = 0
          ExplicitHeight = 21
        end
      end
      object PanelTitulo2: TPanel
        Left = 0
        Top = 29
        Width = 1274
        Height = 26
        Align = alTop
        TabOrder = 8
        object edtTitulo1: TEdit
          Left = 1
          Top = 1
          Width = 1272
          Height = 24
          Align = alClient
          Alignment = taCenter
          Color = clGradientActiveCaption
          TabOrder = 0
          ExplicitHeight = 21
        end
      end
    end
  end
  object Panel26: TPanel
    Left = 0
    Top = 25
    Width = 1282
    Height = 25
    Align = alTop
    TabOrder = 2
    ExplicitWidth = 1280
    object DBNavigator1: TDBNavigator
      Left = 228
      Top = 1
      Width = 96
      Height = 23
      DataSource = FrmDataModule.DataSourceRoteamento
      VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast]
      Align = alLeft
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
    object DateTimePickerProgramacao: TDateTimePicker
      Left = 123
      Top = 1
      Width = 105
      Height = 23
      Hint = 'Data da programa'#231#227'o da rota'
      Align = alLeft
      Date = 42658.000000000000000000
      Time = 0.639985844907641900
      Color = clWhite
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
      OnCloseUp = actLocalizarRoteamentosExecute
    end
    object Panel13: TPanel
      Left = 1
      Top = 1
      Width = 122
      Height = 23
      Align = alLeft
      Alignment = taLeftJustify
      Caption = '   Data da programa'#231#227'o:   '
      TabOrder = 2
    end
    object DBEdit1: TDBEdit
      Left = 324
      Top = 1
      Width = 309
      Height = 23
      Align = alLeft
      Color = clInfoBk
      DataField = 'NomeRota'
      DataSource = FrmDataModule.DataSourceRoteamento
      ReadOnly = True
      TabOrder = 3
      ExplicitHeight = 21
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 520
    Top = 162
    StyleName = 'Platform Default'
    object actPanTo: TAction
      Category = 'Mapa'
      Caption = 'actPanTo'
      Hint = 
        'Posicionar zoom para o N'#243' da Rota selecionada e calcular dist'#226'ni' +
        'cas'
      ImageIndex = 290
    end
    object actImprimirMapa: TAction
      Category = 'Mapa'
      Caption = 'actImprimir'
      Hint = 'Imprimir o mapa'
      ImageIndex = 21
    end
    object actSalvarImagem: TAction
      Category = 'Mapa'
      Caption = 'actSalvar'
      Hint = 'Salvar mapa como imagem'
      ImageIndex = 264
    end
    object actSalvarDesenho: TAction
      Category = 'Mapa'
      Caption = 'actSalvarDesenho'
      Hint = 'Salvar arquivo como desenho 2D (*.zuchi)...'
      ImageIndex = 462
    end
    object actAbrirDesenho: TAction
      Category = 'Mapa'
      Caption = 'actAbrirDesenho'
      Hint = 'Abrir arquivo de desenho 2D (*.zuchi)...'
      ImageIndex = 20
    end
    object actMarcar: TAction
      Category = 'Mapa'
      Caption = 'actMarcar'
      Hint = 'Inserir marcador para todas as plataformas'
      ImageIndex = 112
    end
    object actLocalizarRoteamentos: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actLocalizarRoteamentos'
      OnExecute = actLocalizarRoteamentosExecute
    end
    object actInserirRoteamento: TAction
      Category = 'Roteamento'
      Caption = 'Inserir'
      Hint = 'Inserir roteamento conforme os dados selecionados'
      ImageIndex = 2
      OnExecute = actInserirRoteamentoExecute
    end
    object actExcluirRoteamento: TAction
      Category = 'Roteamento'
      Caption = 'Excluir Rota'
      Hint = 'Excluir roteamento selecionado'
      ImageIndex = 324
      OnExecute = actExcluirRoteamentoExecute
    end
    object actEditarRoteamento: TAction
      Category = 'Roteamento'
      Caption = 'Salvar'
      Hint = 'Salvar altera'#231#245'es conforme os dados selecionados'
      ImageIndex = 18
      OnExecute = actEditarRoteamentoExecute
    end
    object actLimparGrafico: TAction
      Category = 'Mapa'
      Caption = 'Limpar o gr'#225'fico'
      Hint = 'Limpar o gr'#225'fico'
      ImageIndex = 324
    end
    object actGerarLinhas: TAction
      Category = 'Mapa'
      Caption = 'actGerarLinhas'
      Hint = 'Gerar linha da rota selecionada no ComboBox Embarca'#231#227'o'
      ImageIndex = 11
    end
    object actFiltroDestinos: TAction
      Category = 'Sele'#231#227'o destinos'
      Caption = 'Sele'#231#227'o destinos'
      OnExecute = actFiltroDestinosExecute
    end
    object actVincular: TAction
      Category = 'Analise'
      Caption = 'Vincular Exec.'
      Hint = 'Vincular executantes selecionados'
      ImageIndex = 2
      OnExecute = actVincularExecute
    end
    object actExcluirVinculoExecuntanteSelecionado: TAction
      Category = 'Rota Executantes'
      Caption = 'Excluir vinculo dos executantes selecionados'
      ImageIndex = 324
      OnExecute = actExcluirVinculoExecuntanteSelecionadoExecute
    end
    object actProcuraExecutantes: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Procurar executantes programados para a data de roteamento'
      ImageIndex = 28
      ShortCut = 116
      OnExecute = actProcuraExecutantesExecute
    end
    object actProcuraProgramados: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcuraProgramadosExecute
    end
    object actProcuraGeral: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Checked = True
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcuraGeralExecute
    end
    object actAprovarSelecionado: TAction
      Category = 'Executantes'
      Caption = 'Aprovar'
      Hint = 'Aprovar a programa'#231#227'o selecionada'
      ImageIndex = 194
    end
    object actCancelarSelecionado: TAction
      Category = 'Executantes'
      Caption = 'Cancelar'
      Hint = 'Cancelar a programa'#231#227'o selecionada'
      ImageIndex = 190
    end
    object actPAX: TAction
      Category = 'Analise'
      Caption = 'PAX'
      Hint = 'Mostrar executantes PAX da embarca'#231#227'o (Passageiros) '
      ImageIndex = 31
      ShortCut = 117
    end
    object actSelProgExec1: TAction
      Category = 'Executantes'
      Caption = 'actSelProgExec1'
      Hint = 'Selecionar todos os executantes'
      ImageIndex = 154
      OnExecute = actSelProgExec1Execute
    end
    object actSelProgExec2: TAction
      Category = 'Executantes'
      Caption = 'actSelProgExec2'
      Hint = 'Limpar sele'#231#227'o de todos os executantes'
      ImageIndex = 155
      OnExecute = actSelProgExec2Execute
    end
    object actGerarLinhasTodas: TAction
      Category = 'Mapa'
      Caption = 'actGerarLinhasTodas'
      Hint = 'Gerar linhas para todas as rotas'
      ImageIndex = 12
    end
    object actMapaZoomFit: TAction
      Category = 'Mapa'
      Caption = 'actZoomFit'
      Hint = 'Zoom fit'
      ImageIndex = 45
    end
    object actMapaZoomDinamico: TAction
      Category = 'Mapa'
      Caption = 'actZoomDinamico'
      Hint = 'Zoom din'#226'mico'
      ImageIndex = 48
    end
    object actMapaPan: TAction
      Category = 'Mapa'
      Caption = 'actMapaPan'
      Hint = 'Pan'
      ImageIndex = 41
    end
    object actMapaZoomWind: TAction
      Category = 'Mapa'
      Caption = 'actMapaZoomWind'
      Hint = 'Zoom janela'
      ImageIndex = 46
    end
    object actImprimirPlanilha: TAction
      Category = 'Imprimir'
      Caption = 'Imprmir'
      Hint = 'Imprimir planilha'
      ImageIndex = 21
      OnExecute = actImprimirPlanilhaExecute
    end
    object actImprimirRotas: TAction
      Category = 'Imprimir'
      Caption = 'Rotas'
      Hint = 'Imprimir todas as rotas'
      ImageIndex = 21
      OnExecute = actImprimirRotasExecute
    end
    object actRLImprimir_Rota: TAction
      Category = 'Imprimir'
      Caption = 'Rota Selecionada'
      Hint = 'Gerar lista de executantes para a Rota selecionada'
      ImageIndex = 11
      OnExecute = actRLImprimir_RotaExecute
    end
    object actListaRTEmbarque: TAction
      Category = 'Imprimir'
      Caption = 'RT Embarque'
      Hint = 'Lista de executantes com RT de Embarque'
      ImageIndex = 452
      OnExecute = actListaRTEmbarqueExecute
    end
    object actRT_TMIB: TAction
      Category = 'Imprimir'
      Caption = 'RT-TMIB'
      Hint = 'Analise de RT'#39's e manifesto de embarque'
      ImageIndex = 123
      OnExecute = actRT_TMIBExecute
    end
    object actExcelImpressao: TAction
      Category = 'Imprimir'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelImpressaoExecute
    end
    object actExcelProcuraExecutantes: TAction
      Category = 'Executantes'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelProcuraExecutantesExecute
    end
    object actClassificarRegistros: TAction
      Category = 'Imprimir'
      Caption = 'Numerar'
      Hint = 'Numerar registros'
      ImageIndex = 86
      OnExecute = actClassificarRegistrosExecute
    end
    object actKIT_PS: TAction
      Category = 'Imprimir'
      Caption = 'KIT Embarque'
      Hint = 'Lista distribui'#231#227'o KIT Embarque'
      ImageIndex = 458
      OnExecute = actKIT_PSExecute
    end
    object actLanche: TAction
      Category = 'Imprimir'
      Caption = 'Lanches'
      Hint = 'Lista Distribui'#231#227'o de Lanches'
      ImageIndex = 96
      OnExecute = actLancheExecute
    end
    object actExcluirVinculoExecuntanteTodos: TAction
      Category = 'Rota Executantes'
      Caption = 'Excluir vinculo de TODOS os executante filtrados'
      ImageIndex = 324
      OnExecute = actExcluirVinculoExecuntanteTodosExecute
    end
    object actExcel: TAction
      Category = 'Rota Executantes'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelExecute
    end
    object actRemoveRow: TAction
      Category = 'Imprimir'
      Caption = 'Excluir linha selecionada'
      OnExecute = actRemoveRowExecute
    end
    object actRemoveCol: TAction
      Category = 'Imprimir'
      Caption = 'Excluir coluna selecionada'
      OnExecute = actRemoveColExecute
    end
    object actFinalGridImpressao: TAction
      Category = 'Imprimir'
      Caption = 'actFinalGridImpressao'
      OnExecute = actFinalGridImpressaoExecute
    end
    object actFixarLinhaColuna: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actFixarLinhaColuna'
      OnExecute = actFixarLinhaColunaExecute
    end
    object actMapaMundi: TAction
      Category = 'Mapa'
      Caption = 'Mapa Mundi'
      Hint = 'Desenhar o Mapa Mundi'
      ImageIndex = 462
    end
    object actLatitudeLongitude: TAction
      Category = 'Mapa'
      Caption = 'Latitudes e Longitudes'
    end
    object actBrasil: TAction
      Category = 'Mapa'
      Caption = 'Brasil'
      Hint = '462'
      ImageIndex = 462
    end
    object actSalvarRota: TAction
      Category = 'Roteiro Plataformas'
      Caption = 'actSalvarRota'
      ImageIndex = 18
      OnExecute = actSalvarRotaExecute
    end
    object actNumerar: TAction
      Category = 'Rota Executantes'
      Caption = 'Numerar'
      Hint = 'Numerar registros'
      ImageIndex = 86
    end
    object actDisponivel: TAction
      Category = 'Executantes'
      Caption = 'Disponivel'
      Hint = 'Executantes disponiveis para a rota selecionada'
      ImageIndex = 239
      OnExecute = actDisponivelExecute
    end
    object actPreencherRota: TAction
      Category = 'Imprimir'
      Caption = 'actPreencherRota'
      OnExecute = actPreencherRotaExecute
    end
    object actSelecionados: TAction
      Category = 'Executantes'
      Caption = 'actSelecionados'
      OnExecute = actSelecionadosExecute
    end
    object actMostrarVinculados: TAction
      Category = 'Analise'
      Caption = 'actMostrarVinculados'
      Hint = 'Mostrar painel de executantes vinculados a rota selecionada'
      ImageIndex = 32
      OnExecute = actMostrarVinculadosExecute
    end
    object actExcluirLinha: TAction
      Category = 'Roteiro Plataformas'
      Caption = 'Excluir linha selecionada'
      Hint = 'Excluir linha selecionada'
      ImageIndex = 324
      OnExecute = actExcluirLinhaExecute
    end
    object actCalcular: TAction
      Category = 'Roteiro Plataformas'
      Caption = 'Calcular'
      Hint = 'Calcular N'#176' Exec. e Dist'#226'ncias entre Plataformas'
      ImageIndex = 35
      ShortCut = 113
      OnExecute = actCalcularExecute
    end
    object actAjustar: TAction
      Category = 'Imprimir'
      Caption = 'Ajustar'
      Hint = 'Ajustar Colunas'
      ImageIndex = 84
      OnExecute = actAjustarExecute
    end
    object actNumerarAutomatico: TAction
      Category = 'Imprimir'
      Caption = 'actNumerarAutomatico'
      OnExecute = actNumerarAutomaticoExecute
    end
    object actProcuraRotaExecutantes: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcuraRotaExecutantesExecute
    end
    object actTransbordoInterno: TAction
      Category = 'Analise'
      Caption = 'Transbordos'
      Hint = 'Transbordos Internos'
      ImageIndex = 456
      OnExecute = actTransbordoInternoExecute
    end
    object actExcelDestinos: TAction
      Category = 'Sele'#231#227'o destinos'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelDestinosExecute
    end
    object actDuplicados: TAction
      Category = 'Imprimir'
      Caption = 'Transbordos'
      Hint = 'Relat'#243'rio de transbordos internos (Duplicados)'
      ImageIndex = 456
      OnExecute = actDuplicadosExecute
    end
    object actCopiarCelula: TAction
      Category = 'Imprimir'
      Caption = 'Copiar'
      Hint = 'Copiar c'#233'lula selecionada'
      ImageIndex = 44
      ShortCut = 16451
      OnExecute = actCopiarCelulaExecute
    end
    object actColarCelula: TAction
      Category = 'Imprimir'
      Caption = 'Colar'
      Hint = 'Colar c'#233'lula copiada'
      ImageIndex = 24
      ShortCut = 16470
      OnExecute = actColarCelulaExecute
    end
    object actExcluirCelula: TAction
      Category = 'Imprimir'
      Caption = 'Limpar'
      Hint = 'Limpar c'#233'lula'
      ImageIndex = 325
      ShortCut = 46
      OnExecute = actExcluirCelulaExecute
    end
    object actSubstituirPor: TAction
      Category = 'Tabela Roteamento'
      Caption = 'actSubstituirPor'
      OnExecute = actSubstituirPorExecute
    end
    object actFiltroInserirRoteamento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'actFiltroInserirRoteamento'
    end
    object actGridASCRotemento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'actGridASCRotemento'
    end
    object actGridDESCRoteamento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'actGridDESCRoteamento'
    end
    object actLimparFiltrosRoteamento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'Limpar'
      Hint = 'Limpar Filtros'
      ImageIndex = 469
      OnExecute = actLimparFiltrosRoteamentoExecute
    end
    object actFiltrosTabelaRoteamento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'Filtros'
      Hint = 'Visualizar tabela de filtros'
      ImageIndex = 470
      OnExecute = actFiltrosTabelaRoteamentoExecute
    end
    object actProcuraFiltrosTabelaRoteamento: TAction
      Category = 'Tabela Roteamento'
      Caption = 'actProcuraFiltrosTabelaRoteamento'
      OnExecute = actProcuraFiltrosTabelaRoteamentoExecute
    end
    object actProcuraRoteamento: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcuraRoteamentoExecute
    end
    object actMapaRoteiro: TAction
      Category = 'Mapa'
      Caption = 'actMapaRoteiro'
    end
    object Action1: TAction
      Category = 'Rota Executantes'
      Caption = 'Action1'
    end
    object actSelRotaExecTODOS: TAction
      Category = 'Rota Executantes'
      Caption = 'actSelRotaExecTODOS'
      Hint = 'Selecionar todos os executantes'
      ImageIndex = 154
      OnExecute = actSelRotaExecTODOSExecute
    end
    object actSelRotaExecLIMPAR: TAction
      Category = 'Rota Executantes'
      Caption = 'actSelRotaExecLIMPAR'
      Hint = 'Limpar a sele'#231#227'o de todos os executantes'
      ImageIndex = 155
      OnExecute = actSelRotaExecLIMPARExecute
    end
    object actSelecionadosRotaExecutantes: TAction
      Category = 'Rota Executantes'
      Caption = 'actSelecionadosRotaExecutantes'
      OnExecute = actSelecionadosRotaExecutantesExecute
    end
    object actImprimirExtrato: TAction
      Category = 'Imprimir'
      Caption = 'actImprimirExtrato'
      Hint = 'Imprimir extrato de passageiros'
      ImageIndex = 21
      OnExecute = actImprimirExtratoExecute
    end
    object actExtratoOrigem: TAction
      Category = 'Imprimir'
      Caption = 'Extrato de executantes (Origem Cadastro)'
      Hint = 'Extrato de executantes (Origem Cadastro)'
      ImageIndex = 459
      OnExecute = actExtratoOrigemExecute
    end
  end
  object SavePictureDialog1: TSavePictureDialog
    Filter = 'Metafiles (*.wmf)|*.wmf'
    Left = 552
    Top = 160
  end
  object PopupMenuImpressao: TPopupMenu
    Images = FrmPrincipal.ImageList1
    Left = 564
    Top = 249
    object Excluirlinhaselecionada1: TMenuItem
      Action = actRemoveRow
    end
    object Excluircolunaselecionada1: TMenuItem
      Action = actRemoveCol
    end
    object ClassificarCrescentepelacolunaselecionada1: TMenuItem
      Caption = 'Classificar pela coluna selecionada'
      OnClick = ClassificarCrescentepelacolunaselecionada1Click
    end
    object Fixar1Linha1: TMenuItem
      AutoCheck = True
      Caption = 'Fixar 1'#176' Linha'
      Checked = True
      OnClick = actFixarLinhaColunaExecute
    end
    object Fixar1Coluna1: TMenuItem
      AutoCheck = True
      Caption = 'Fixar 1'#176' Coluna'
      Checked = True
      OnClick = actFixarLinhaColunaExecute
    end
    object Ajustartamanhodascolunas1: TMenuItem
      Caption = 'Ajustar tamanho das colunas'
      OnClick = Ajustartamanhodascolunas1Click
    end
    object Copiar1: TMenuItem
      Action = actCopiarCelula
    end
    object Colar1: TMenuItem
      Action = actColarCelula
    end
    object Limpar1: TMenuItem
      Action = actExcluirCelula
    end
  end
  object PopupMenuMapa: TPopupMenu
    Left = 605
    Top = 251
    object actMapaMundi1: TMenuItem
      Action = actMapaMundi
    end
    object LatitudeseLongitudes1: TMenuItem
      Action = actLatitudeLongitude
    end
    object Brasil1: TMenuItem
      Action = actBrasil
    end
  end
  object SaveDialog1: TSaveDialog
    FileName = 'Desenho.zuchi'
    Filter = 'Desenho (*.zuchi)|*.zuchi|Todos os Arquivos (*.*)|*.*'
    Left = 581
    Top = 163
  end
  object OpenDialog1: TOpenDialog
    Filter = 'Desenho (*.zuchi)|*.zuchi|Todos os Arquivos (*.*)|*.*'
    Left = 616
    Top = 160
  end
  object PopupMenuExcluirVinculados: TPopupMenu
    Left = 633
    Top = 252
    object Excluirvinculo1: TMenuItem
      Action = actExcluirVinculoExecuntanteSelecionado
      Caption = 'Excluir vinculo do executante selecionados'
    end
    object Excluirvinculodoexecutantetodos1: TMenuItem
      Action = actExcluirVinculoExecuntanteTodos
    end
  end
  object PopupMenuRota: TPopupMenu
    Left = 505
    Top = 252
    object Excluirlinhaselecionada2: TMenuItem
      Action = actExcluirLinha
    end
  end
end
