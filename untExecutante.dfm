object FrmExecutante: TFrmExecutante
  Left = 0
  Top = 0
  Caption = 'Executantes'
  ClientHeight = 488
  ClientWidth = 1159
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
  object Panel7: TPanel
    Left = 0
    Top = 0
    Width = 1159
    Height = 25
    Align = alTop
    Caption = 'Cadastro de Executantes'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 1157
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 25
    Width = 1159
    Height = 463
    ActivePage = TabSheet1
    Align = alClient
    TabOrder = 1
    ExplicitWidth = 1157
    ExplicitHeight = 455
    object TabSheet1: TTabSheet
      Caption = 'Executante'
      object DBGridExecutante: TFilterDBGrid
        Left = 0
        Top = 29
        Width = 1151
        Height = 387
        Align = alClient
        Color = clWhite
        DataSource = FrmDataModule.DataSourceExecutante
        Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGridExecutanteCellClick
        OnDrawColumnCell = DBGridExecutanteDrawColumnCell
        OnKeyPress = DBGridExecutanteKeyPress
        ClearFilterButton = btnFiltroClearExecutante
        SearchAction = actProcurarExecutante
        LayoutGrid = ColunasLayoutExecutante
        EnableZebra = False
        LayoutButton = btnLayoutExecutante
        ExcelButton = btnExcelExecutante
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Ativo'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Turma'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'NomeExecutante'
            Title.Alignment = taCenter
            Title.Caption = 'Nome do Executante'
            Width = 200
            Visible = True
          end
          item
            Alignment = taCenter
            DropDownRows = 30
            Expanded = False
            FieldName = 'OrigemCadastro'
            Title.Alignment = taCenter
            Title.Caption = 'Origem Cadastro'
            Width = 87
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Chave'
            Title.Alignment = taCenter
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'CodigoSAP'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo SAP'
            Width = 100
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'txtFuncao'
            Title.Alignment = taCenter
            Title.Caption = 'Fun'#231#227'o'
            Width = 150
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'txtEmpresa'
            Title.Alignment = taCenter
            Title.Caption = 'Empresa'
            Width = 150
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'CPF'
            Title.Alignment = taCenter
            Width = 100
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'OutroDocumento'
            Title.Alignment = taCenter
            Title.Caption = 'Passaporte'
            Width = 91
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'txtTipoEtapaServico'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo de Etapa de Servi'#231'o'
            Width = 150
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'codigoTipoEtapaServico'
            Title.Alignment = taCenter
            Title.Caption = 'C'#243'digo Tipo de Servico'
            Width = 120
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'RequisitantePT'
            Title.Alignment = taCenter
            Title.Caption = 'Requisitante PT'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CentroCusto'
            Title.Alignment = taCenter
            Title.Caption = 'Centro de Custo'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DiagramaRede'
            Title.Alignment = taCenter
            Title.Caption = 'Diagrama de Rede'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'OperRede'
            Title.Alignment = taCenter
            Title.Caption = 'Oper. Rede'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'ElementoPEP'
            Title.Alignment = taCenter
            Title.Caption = 'Elemento PEP'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataValidadePT'
            Title.Alignment = taCenter
            Title.Caption = 'Data Validade PT'
            Width = 92
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'booleanCESS'
            Title.Alignment = taCenter
            Title.Caption = 'Curso CESS'
            Visible = True
          end
          item
            Expanded = False
            FieldName = 'DataValidadeCESS'
            Title.Alignment = taCenter
            Title.Caption = 'Data Validade CESS'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataAtualizacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Atualiza'#231#227'o'
            Width = 90
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'AvaliadoPor'
            Title.Alignment = taCenter
            Title.Caption = 'Avaliado Por'
            Width = 66
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'Computador'
            Title.Alignment = taCenter
            Width = 69
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'idExecutante'
            ReadOnly = True
            Title.Caption = 'ID'
            Visible = True
          end>
      end
      object ToolBar1: TToolBar
        Left = 0
        Top = 0
        Width = 1151
        Height = 29
        Caption = 'ToolBar1'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 1
        ExplicitWidth = 1149
        object DBNavigatorExecutante: TDBNavigator
          Left = 0
          Top = 0
          Width = 156
          Height = 22
          DataSource = FrmDataModule.DataSourceExecutante
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
        object ToolButton1: TToolButton
          Left = 156
          Top = 0
          Width = 8
          Caption = 'ToolButton1'
          ImageIndex = 126
          Style = tbsSeparator
        end
        object ToolButton2: TToolButton
          Left = 164
          Top = 0
          Hint = 'Procura por filtro ou duplicados (F5)'
          Caption = 'ToolButton2'
          DropdownMenu = PopupMenuFiltros
          ImageIndex = 28
          ParentShowHint = False
          ShowHint = True
          Style = tbsDropDown
          OnClick = actProcurarExecutanteExecute
        end
        object btnAtualizar: TToolButton
          Left = 206
          Top = 0
          Hint = 
            'Atualizar "Tipo de Etapa de Servi'#231'o" e Corrigir "ma'#237'sculo e espa' +
            #231'os"'
          Caption = 'btnAtualizar'
          DropdownMenu = PopupMenuAtualizar
          ImageIndex = 238
          ParentShowHint = False
          ShowHint = True
          Style = tbsDropDown
        end
        object ToolButton3: TToolButton
          Left = 248
          Top = 0
          Hint = 'Excluir registros'
          Caption = 'ToolButton3'
          DropdownMenu = PopupMenuExcluir
          ImageIndex = 324
          ParentShowHint = False
          ShowHint = True
          Style = tbsDropDown
        end
        object btnFiltroClearExecutante: TToolButton
          Left = 290
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnFiltroClearExecutante'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object btnLayoutExecutante: TToolButton
          Left = 313
          Top = 0
          Hint = 'Editar layout de colunas'
          Caption = 'btnLayoutExecutante'
          ImageIndex = 196
          ParentShowHint = False
          ShowHint = True
        end
        object btnExcelExecutante: TToolButton
          Left = 336
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'btnExcelExecutante'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
      end
      object DBLookupComboBox_TipoServico: TDBLookupComboBox
        Left = 520
        Top = 125
        Width = 172
        Height = 21
        Color = clWhite
        DataField = 'codigoTipoEtapaServico'
        DataSource = FrmDataModule.DataSourceExecutante
        DropDownRows = 12
        KeyField = 'idTipoEtapaServico'
        ListField = 'TipoEtapaServico'
        ListSource = FrmDataModule.DataSourceTipoEtapaServico
        TabOrder = 2
        Visible = False
        OnCloseUp = DBLookupComboBox_TipoServicoCloseUp
      end
      object StatusBarExecutntes: TStatusBar
        Left = 0
        Top = 416
        Width = 1151
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
          end
          item
            Width = 50
          end>
        ExplicitTop = 408
        ExplicitWidth = 1149
      end
      object ColunasLayoutExecutante: TStringGrid
        Left = 21
        Top = 110
        Width = 113
        Height = 94
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
      object MaskEditCPF: TMaskEdit
        Left = 363
        Top = 183
        Width = 120
        Height = 21
        EditMask = '000\.000\.000\-00;;_'
        MaxLength = 14
        ParentShowHint = False
        ShowHint = True
        TabOrder = 5
        Text = '   .   .   -  '
        Visible = False
        OnChange = MaskEditCPFChange
      end
    end
    object TabSheet2: TTabSheet
      Caption = 'Tipo de Etapa de Servi'#231'o'
      ImageIndex = 2
      object PanelEspecialidade1: TPanel
        Left = 0
        Top = 0
        Width = 1151
        Height = 435
        Align = alClient
        TabOrder = 0
        object ToolBar3: TToolBar
          Left = 1
          Top = 1
          Width = 1149
          Height = 29
          Caption = 'ToolBar2'
          DrawingStyle = dsGradient
          EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
          Images = FrmPrincipal.ImageList1
          TabOrder = 0
          object DBNavigatorTipoEtapaServico: TDBNavigator
            Left = 0
            Top = 0
            Width = 138
            Height = 22
            DataSource = FrmDataModule.DataSourceTipoEtapaServico
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
          object btnFiltroClearTipoEtapaServico: TToolButton
            Left = 138
            Top = 0
            Hint = 'Limpar filtro'
            Caption = 'btnFiltroClearTipoEtapaServico'
            ImageIndex = 225
            ParentShowHint = False
            ShowHint = True
          end
          object btnLayoutTipoEtapaServico: TToolButton
            Left = 161
            Top = 0
            Hint = 'Editar layout de colunas'
            Caption = 'btnLayoutTipoEtapaServico'
            ImageIndex = 196
            ParentShowHint = False
            ShowHint = True
          end
          object btnExcelTIpoServico: TToolButton
            Left = 184
            Top = 0
            Hint = 'Exportar dados para o Excel'
            Caption = 'btnExcelTIpoServico'
            ImageIndex = 54
            ParentShowHint = False
            ShowHint = True
          end
        end
        object DBGridTipoServico: TFilterDBGrid
          Left = 1
          Top = 30
          Width = 1149
          Height = 385
          Align = alClient
          DataSource = FrmDataModule.DataSourceTipoEtapaServico
          Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
          TabOrder = 1
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          OnCellClick = DBGridTipoServicoCellClick
          OnDrawColumnCell = DBGridTipoServicoDrawColumnCell
          ClearFilterButton = btnFiltroClearTipoEtapaServico
          SearchAction = actProcurarTipoEtapaServico
          LayoutGrid = ColunasLayoutTipoEtapaServico
          EnableZebra = False
          LayoutButton = btnLayoutTipoEtapaServico
          ExcelButton = btnExcelTIpoServico
          Columns = <
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'idTipoEtapaServico'
              Title.Alignment = taCenter
              Title.Caption = 'C'#243'digo'
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'TipoEtapaServico'
              Title.Alignment = taCenter
              Title.Caption = 'Tipo de Etapa de Servi'#231'o'
              Width = 208
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'Sigla'
              Title.Alignment = taCenter
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'ServicoPadrao'
              Title.Alignment = taCenter
              Title.Caption = 'Servi'#231'o Padr'#227'o na Programa'#231#227'o Di'#225'ria'
              Width = 250
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'AvaliadoPor'
              Title.Alignment = taCenter
              Title.Caption = 'Avaliado Por'
              Width = 70
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
              Width = 68
              Visible = True
            end>
        end
        object StatusBarTipoEtapaServico: TStatusBar
          Left = 1
          Top = 415
          Width = 1149
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
        object ColunasLayoutTipoEtapaServico: TStringGrid
          Left = 158
          Top = 110
          Width = 113
          Height = 94
          ColCount = 7
          DefaultRowHeight = 21
          RowCount = 7
          TabOrder = 3
          Visible = False
          RowHeights = (
            21
            20
            21
            21
            21
            21
            21)
        end
      end
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 460
    Top = 153
    StyleName = 'Platform Default'
    object actExcelImportar: TAction
      Category = 'Excel'
      Caption = 'Importar do Excel'
      Hint = 'Importar tabela do Excel'
      ImageIndex = 251
      OnExecute = actExcelImportarExecute
    end
    object actProcurarExecutante: TAction
      Category = 'Localizar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecutanteExecute
    end
    object actAtualizarTodos: TAction
      Category = 'Atualizar'
      Caption = 'Atualizar e Corrigr (Todos)'
      ImageIndex = 238
      OnExecute = actAtualizarTodosExecute
    end
    object actAtualizarSelecionado: TAction
      Category = 'Atualizar'
      Caption = 'Atualizar e Corrigr (Selecionado)'
      ImageIndex = 238
      OnExecute = actAtualizarSelecionadoExecute
    end
    object actExcluirTodos: TAction
      Category = 'Excluir'
      Caption = 'Excluir todos os registros "FILTRADOS"'
      Hint = 'Excluir todos os registros "FILTRADOS"'
      ImageIndex = 324
      OnExecute = actExcluirTodosExecute
    end
    object actExcluirSelecionado: TAction
      Category = 'Excluir'
      Caption = 'Excluir Selecionado'
      Hint = 'Excluir registro selecionado sem perguntar'
      ImageIndex = 324
      OnExecute = actExcluirSelecionadoExecute
    end
    object actBuscaEmpregadoMatricula: TAction
      Category = 'Busca Empregado'
      Caption = 'Busca Empregado (Matricula)'
      Hint = 'Busca empregado por Matricula'
      ImageIndex = 27
      ShortCut = 117
    end
    object actBuscaEmpregadoChave: TAction
      Category = 'Busca Empregado'
      Caption = 'Busca Empregado (Chave)'
      Hint = 'Busca empregado pela Chave'
      ImageIndex = 27
      ShortCut = 118
    end
    object actBuscaEmpregadoNome: TAction
      Category = 'Busca Empregado'
      Caption = 'Busca Empregado (Nome)'
      Hint = 'Busca empregado pela Nome'
      ImageIndex = 27
      ShortCut = 119
    end
    object actFiltros: TAction
      Category = 'Localizar'
      Caption = 'Procurar'
      Hint = 
        'Procurar Executantes conforme filtro determinado nos campos abai' +
        'xo'
      ImageIndex = 28
      ShortCut = 116
      OnExecute = actFiltrosExecute
    end
    object actProcuraDuplicadoMatricula: TAction
      Category = 'Localizar'
      Caption = 'Procura Duplicados (Matricula)'
      Hint = 'Procura Duplicados (Matricula)'
      ImageIndex = 28
      OnExecute = actProcuraDuplicadoMatriculaExecute
    end
    object actDuplicadoNome: TAction
      Category = 'Localizar'
      Caption = 'Procura Duplicados (Nome)'
      Hint = 'Procura Duplicados (Nome)'
      ImageIndex = 28
      OnExecute = actDuplicadoNomeExecute
    end
    object actProcurarTipoEtapaServico: TAction
      Category = 'Localizar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarTipoEtapaServicoExecute
    end
    object actBuscaOrigemCadastro: TAction
      Category = 'Atualizar'
      Caption = 'Definir Origem'
      Hint = 
        'Busca de Origem na janela de consulta de executantes programados' +
        ' por periodo'
      ImageIndex = 123
      OnExecute = actBuscaOrigemCadastroExecute
    end
    object actTurma: TAction
      Category = 'Atualizar'
      Caption = 'Definir Turma'
      Hint = 'Defini'#231#227'o da turma'
      ImageIndex = 123
      OnExecute = actTurmaExecute
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 332
    Top = 153
  end
  object PopupMenuAtualizar: TPopupMenu
    Left = 188
    Top = 289
    object AtualizareCorrigrSelecionado1: TMenuItem
      Action = actAtualizarSelecionado
    end
    object AtualizareCorrigrTodos1: TMenuItem
      Action = actAtualizarTodos
    end
  end
  object PopupMenuExcluir: TPopupMenu
    Left = 548
    Top = 241
    object ExcluirSelecionado1: TMenuItem
      Action = actExcluirSelecionado
      Caption = 'Excluir registro selecionado sem perguntar'
      ShortCut = 123
    end
    object Excluirtodos1: TMenuItem
      Action = actExcluirTodos
    end
  end
  object PopupMenuFiltros: TPopupMenu
    Left = 308
    Top = 265
    object Procurar1: TMenuItem
      Action = actFiltros
      Checked = True
    end
    object ProcuraDuplicadosMatricula1: TMenuItem
      Action = actProcuraDuplicadoMatricula
    end
    object ProcuraDuplicadosNome1: TMenuItem
      Action = actDuplicadoNome
    end
  end
  object PopupMenuExcel: TPopupMenu
    Left = 420
    Top = 289
    object Exportar1: TMenuItem
      Caption = 'Exportar para o Excel'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
    end
    object Importar1: TMenuItem
      Action = actExcelImportar
    end
  end
end
