object FrmProgramacaoDiaria: TFrmProgramacaoDiaria
  Left = 0
  Top = 0
  Caption = 'Programa'#231#227'o Di'#225'ria de Servi'#231'os'
  ClientHeight = 634
  ClientWidth = 1374
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
    Width = 1374
    Height = 25
    Align = alTop
    Caption = 'Programa'#231#227'o Di'#225'ria de Servi'#231'os'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 1372
  end
  object PanelCriacao: TPanel
    Left = 0
    Top = 25
    Width = 900
    Height = 609
    Align = alLeft
    Caption = 'PanelCriacao'
    TabOrder = 1
    ExplicitHeight = 601
    object Splitter3: TSplitter
      Left = 577
      Top = 1
      Height = 607
      ExplicitLeft = 382
      ExplicitHeight = 462
    end
    object PanelResultados: TPanel
      Left = 580
      Top = 1
      Width = 319
      Height = 607
      Align = alClient
      TabOrder = 0
      ExplicitHeight = 599
      object Splitter2: TSplitter
        Left = 1
        Top = 281
        Width = 317
        Height = 3
        Cursor = crVSplit
        Align = alTop
        ExplicitTop = 209
        ExplicitWidth = 204
      end
      object PanelExecutante: TPanel
        Left = 1
        Top = 1
        Width = 317
        Height = 280
        Align = alTop
        TabOrder = 0
        object DBGridExecutante: TFilterDBGrid
          Left = 1
          Top = 55
          Width = 315
          Height = 205
          Align = alClient
          DataSource = FrmDataModule.DataSourceProgramacaoExecutante_Cadastro
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          OnDrawColumnCell = DBGridExecutanteDrawColumnCell
          OnKeyDown = DBGridExecutanteKeyDown
          OnKeyPress = DBGridExecutanteKeyPress
          ClearFilterButton = btnFiltroClearExecutante
          SearchAction = actProcurarExecutantes
          EnableZebra = False
          Columns = <
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'booleanRecolhimento'
              Title.Alignment = taCenter
              Title.Caption = 'Recolhimento'
              Width = 80
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'Origem'
              Title.Alignment = taCenter
              Width = 100
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
              Expanded = False
              FieldName = 'Funcao'
              Title.Alignment = taCenter
              Title.Caption = 'Fun'#231#227'o'
              Width = 170
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'Empresa'
              Title.Alignment = taCenter
              Width = 100
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'StatusProgramacao'
              Title.Alignment = taCenter
              Title.Caption = 'Status [Programa'#231#227'o]'
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
              FieldName = 'Documento'
              Title.Alignment = taCenter
              Width = 107
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
              Alignment = taCenter
              Expanded = False
              FieldName = 'RequisitantePT'
              Title.Caption = 'Req. PT'
              Visible = True
            end>
        end
        object ToolBar4: TToolBar
          Left = 1
          Top = 26
          Width = 315
          Height = 29
          Caption = 'ToolBar2'
          DrawingStyle = dsGradient
          EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
          Images = FrmPrincipal.ImageList1
          TabOrder = 1
          object btnFiltroClearExecutante: TToolButton
            Left = 0
            Top = 0
            Hint = 'Limpar filtro'
            Caption = 'btnFiltroClearExecutante'
            ImageIndex = 225
            ParentShowHint = False
            ShowHint = True
          end
        end
        object Panel19: TPanel
          Left = 1
          Top = 1
          Width = 315
          Height = 25
          Align = alTop
          Caption = 'Executantes Programados'
          Color = clMenuBar
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 2
        end
        object StatusBarExecutantes: TStatusBar
          Left = 1
          Top = 260
          Width = 315
          Height = 19
          Panels = <
            item
              Width = 50
            end
            item
              Width = 50
            end>
        end
      end
      object PanelServico: TPanel
        Left = 1
        Top = 284
        Width = 317
        Height = 322
        Align = alClient
        TabOrder = 1
        ExplicitHeight = 314
        object DBGridServico: TFilterDBGrid
          Left = 1
          Top = 55
          Width = 315
          Height = 247
          Align = alClient
          DataSource = FrmDataModule.DataSourceProgramacaoServico_Cadastro
          Options = [dgEditing, dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgAlwaysShowSelection, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
          TabOrder = 0
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          OnKeyDown = DBGridServicoKeyDown
          OnKeyPress = DBGridServicoKeyPress
          ClearFilterButton = btnFiltroClearServicos
          SearchAction = actProcurarServicos
          EnableZebra = False
          Columns = <
            item
              Expanded = False
              FieldName = 'CampoSelecao'
              Title.Alignment = taCenter
              Title.Caption = 'Servi'#231'o/Ano-Etapa'
              Width = 100
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'TextoBreveOP'
              Title.Alignment = taCenter
              Title.Caption = 'Texto Breve da Opera'#231#227'o'
              Width = 250
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'TextoBreveOM'
              Title.Alignment = taCenter
              Title.Caption = 'Texto Breve da Ordem'
              Visible = True
            end
            item
              Expanded = False
              FieldName = 'OrdemManutencao'
              Title.Alignment = taCenter
              Title.Caption = 'Ordem de manuten'#231#227'o'
              Width = 120
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'Operacao'
              Title.Alignment = taCenter
              Title.Caption = 'Opera'#231#227'o'
              Visible = True
            end
            item
              Alignment = taCenter
              Expanded = False
              FieldName = 'CentroTrabalhoOP'
              Title.Alignment = taCenter
              Title.Caption = 'Centro Trabalho [OP]'
              Visible = True
            end>
        end
        object Panel20: TPanel
          Left = 1
          Top = 1
          Width = 315
          Height = 25
          Align = alTop
          Caption = 'Servi'#231'os Programados'
          Color = clMenuBar
          Font.Charset = DEFAULT_CHARSET
          Font.Color = clWindowText
          Font.Height = -13
          Font.Name = 'Tahoma'
          Font.Style = [fsBold]
          ParentBackground = False
          ParentFont = False
          TabOrder = 1
        end
        object StatusBarServicos: TStatusBar
          Left = 1
          Top = 302
          Width = 315
          Height = 19
          Panels = <
            item
              Width = 50
            end
            item
              Width = 50
            end>
          ExplicitTop = 294
        end
        object ToolBar7: TToolBar
          Left = 1
          Top = 26
          Width = 315
          Height = 29
          Caption = 'ToolBar2'
          DrawingStyle = dsGradient
          EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
          Images = FrmPrincipal.ImageList1
          TabOrder = 3
          object btnFiltroClearServicos: TToolButton
            Left = 0
            Top = 0
            Hint = 'Limpar filtro'
            Caption = 'btnFiltroClearServicos'
            ImageIndex = 225
            ParentShowHint = False
            ShowHint = True
          end
          object BitBtn15: TBitBtn
            Left = 23
            Top = 0
            Width = 60
            Height = 22
            Action = actServicoInserir1
            Caption = 'Inserir'
            ParentShowHint = False
            ShowHint = True
            TabOrder = 0
          end
          object BitBtn14: TBitBtn
            Left = 83
            Top = 0
            Width = 60
            Height = 22
            Action = actServicoExcluir
            Caption = 'Excluir'
            ParentShowHint = False
            ShowHint = True
            TabOrder = 1
          end
        end
      end
    end
    object Panel6: TPanel
      Left = 1
      Top = 1
      Width = 576
      Height = 607
      Align = alLeft
      TabOrder = 1
      ExplicitHeight = 599
      object ToolBar3: TToolBar
        Left = 1
        Top = 26
        Width = 574
        Height = 29
        Caption = 'ToolBar1'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 0
        object DateProgramacao: TDateTimePicker
          Left = 0
          Top = 0
          Width = 94
          Height = 22
          Hint = 'Selecine a data da programa'#231#227'o para consulta'
          Align = alLeft
          Date = 42651.000000000000000000
          Time = 0.462823229157947900
          ParentShowHint = False
          ShowHint = True
          TabOrder = 3
          OnChange = DateProgramacaoChange
        end
        object btnFiltroClearProgramacao: TToolButton
          Left = 94
          Top = 0
          Hint = 'Limpar filtro'
          Caption = 'btnFiltroClearProgramacao'
          ImageIndex = 225
          ParentShowHint = False
          ShowHint = True
        end
        object BitBtn3: TBitBtn
          Left = 117
          Top = 0
          Width = 70
          Height = 22
          Action = actProcurarProgramacao
          Caption = 'Procurar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 5
        end
        object BitBtn11: TBitBtn
          Left = 187
          Top = 0
          Width = 70
          Height = 22
          Action = actInserirProgramacaoBranco
          Caption = 'Inserir'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
        object BitBtn5: TBitBtn
          Left = 257
          Top = 0
          Width = 70
          Height = 22
          Action = actEditarProgramacao
          Caption = 'Editar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 4
        end
        object BitBtn12: TBitBtn
          Left = 327
          Top = 0
          Width = 70
          Height = 22
          Action = actLogAcao
          Caption = 'Hist'#243'rico'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
        end
        object ToolButton7: TToolButton
          Left = 397
          Top = 0
          Hint = 'Excluir programa'#231#227'o'
          Caption = 'ToolButton7'
          DropdownMenu = PopupMenuExcluir
          ImageIndex = 324
          ParentShowHint = False
          ShowHint = True
          Style = tbsDropDown
        end
        object ToolButton8: TToolButton
          Left = 439
          Top = 0
          Hint = 'Copiar programa'#231#227'o para data no futuro'
          Caption = 'ToolButton8'
          DropdownMenu = PopupMenuCopiar
          ImageIndex = 138
          ParentShowHint = False
          ShowHint = True
          Style = tbsDropDown
        end
        object ToolButton9: TToolButton
          Left = 481
          Top = 0
          Action = actZoomMais
        end
        object ToolButton10: TToolButton
          Left = 504
          Top = 0
          Action = actZoomMenos
        end
        object DBEditProgramacao: TDBEdit
          Left = 527
          Top = 0
          Width = 31
          Height = 22
          DataField = 'idProgramacaoDiaria'
          DataSource = FrmDataModule.DataSourceProgramacaoDiaria_Cadastro
          TabOrder = 0
          Visible = False
          OnChange = DBEditProgramacaoChange
        end
      end
      object DBGridProgramacao: TFilterDBGrid
        Left = 1
        Top = 55
        Width = 574
        Height = 532
        Align = alClient
        DataSource = FrmDataModule.DataSourceProgramacaoDiaria_Cadastro
        ReadOnly = True
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGridProgramacaoDrawColumnCell
        ClearFilterButton = btnFiltroClearProgramacao
        SearchAction = actProcurarProgramacao
        EnableZebra = False
        Columns = <
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'booleanFavorito'
            Title.Alignment = taCenter
            Title.Caption = 'Fav.'
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
            FieldName = 'txtTipoEtapaServico'
            Title.Alignment = taCenter
            Title.Caption = 'Tipo de Etapa de Servi'#231'o'
            Width = 210
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumExecutantes'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Exe.'
            Width = 45
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumCancelados'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Can.'
            Width = 45
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'NumAprovados'
            Title.Alignment = taCenter
            Title.Caption = 'N'#176' Apr.'
            Width = 45
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'CriadoPor'
            Title.Alignment = taCenter
            Title.Caption = 'Criado Por'
            Width = 70
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataCriacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Cria'#231#227'o'
            Width = 150
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'ComputadorCriacao'
            Title.Alignment = taCenter
            Title.Caption = 'Computador'
            Width = 82
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'AvaliadoPor'
            Title.Alignment = taCenter
            Title.Caption = 'Atualizado Por'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'ComputadorAtualizacao'
            Title.Alignment = taCenter
            Title.Caption = 'Computador Atualiza'#231#227'o'
            Visible = True
          end
          item
            Alignment = taCenter
            Expanded = False
            FieldName = 'DataAtualizacao'
            Title.Alignment = taCenter
            Title.Caption = 'Data Atualiza'#231#227'o'
            Visible = True
          end>
      end
      object Panel7: TPanel
        Left = 1
        Top = 1
        Width = 574
        Height = 25
        Align = alTop
        Caption = 'Programa'#231#245'es Criadas'
        Color = clMenuBar
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 2
      end
      object StatusBarProgramacao: TStatusBar
        Left = 1
        Top = 587
        Width = 574
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
        ExplicitTop = 579
      end
    end
  end
  object PanelProgramacao: TPanel
    Left = 900
    Top = 25
    Width = 474
    Height = 609
    Align = alClient
    TabOrder = 2
    Visible = False
    ExplicitWidth = 472
    ExplicitHeight = 601
    object Splitter1: TSplitter
      Left = 1
      Top = 385
      Width = 472
      Height = 3
      Cursor = crVSplit
      Align = alBottom
      ExplicitTop = 264
      ExplicitWidth = 199
    end
    object ToolBar1: TToolBar
      Left = 1
      Top = 1
      Width = 472
      Height = 29
      ButtonWidth = 24
      Caption = 'ToolBar1'
      DrawingStyle = dsGradient
      EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
      HotTrackColor = clRed
      Images = FrmPrincipal.ImageList1
      TabOrder = 0
      ExplicitWidth = 470
      object BitBtn21: TBitBtn
        Left = 0
        Top = 0
        Width = 145
        Height = 22
        Action = actExecutantesServicos
        Caption = 'Executantes e Servi'#231'os'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
      end
      object BitBtn22: TBitBtn
        Left = 145
        Top = 0
        Width = 75
        Height = 22
        Action = actcalcelarProgramacao
        Caption = 'Cancelar'
        ImageIndex = 190
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000FF00FF00FF00
          FF00BFC6CA00B5BCC000BEC5C900FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00C0C8CC00B7BEC200FF00FF00FF00FF00BFC6
          CA008A909200565A5B007A7E8100BBC3C700FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00C7CFD300B3BABE0083888B0071757800B7BEC200FF00FF007665
          D7002900DF0036219F004F5254009EA4A700C6CED200FF00FF00FF00FF00FF00
          FF00C6CED200A7AEB1005141A5002C06D70083888B00C0C8CC00FF00FF00968F
          D7002900DF002E08DA0042405D0064686A00ADB4B800FF00FF00FF00FF00C6CE
          D2009DA1B2003D21BC002900DF006657BB00BCC4C800FF00FF00FF00FF00FF00
          FF006F5BDA002900DF003717CD004B4D4F0074797B00B0B8BB00C6CED2009DA1
          B2003414C9002C06D700827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00654EDB002900DF00462EB9004E525300737779009093A3003414
          C9002C06D700827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00654EDB002900DF00452DB90041415000300FC4002B05
          D600827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00644DDA002900DF002B05D6002A04D5005550
          7B00ABB2B500C7CFD300FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00C6CED2009DA1B2002D07D8002900DF003E23BD005B5F
          600072767900A7AEB100C3CBCF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00C6CED2009DA1B2003414C9002C06D7005D4BBC00320CDD004C32
          CC0062647400686C6F00999FA200BFC6CA00FF00FF00FF00FF00FF00FF00FF00
          FF00C3CBCF009DA1B2003414C9002900DF006F68A300BCC4C800AAA9D6003C19
          DD004121D700615D8900616567008C929500C0C8CC00FF00FF00FF00FF00C5CD
          D1009497A8003413C8002900DF005D51A600B8BFC300FF00FF00FF00FF00BEC3
          D5005134DC00320BDD00635AA1006B707200ACB3B600C6CED200FF00FF009D9B
          C8003413C8002900DF004B39AA00B2B9BD00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF007868DA002900DF00A4A5C400C6CED200FF00FF00FF00FF003B18
          DB002900DF003D21BC00A7AEB100C7CFD300FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00330D
          DE004424DA00B0B4C600C7CFD300FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
        ParentShowHint = False
        ShowHint = True
        TabOrder = 1
      end
      object btnInserirEditarProgramacao: TBitBtn
        Left = 220
        Top = 0
        Width = 60
        Height = 22
        Action = actInserirProgramacao
        Caption = 'Inserir'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 0
      end
      object BitBtn9: TBitBtn
        Left = 280
        Top = 0
        Width = 75
        Height = 22
        Action = actAtualizar
        Caption = 'Atualizar'
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
      end
    end
    object PanelSelecaoExecutante: TPanel
      Left = 1
      Top = 98
      Width = 472
      Height = 287
      Align = alClient
      TabOrder = 1
      ExplicitWidth = 470
      ExplicitHeight = 279
      object PanelTituloExecutantes: TPanel
        Left = 1
        Top = 1
        Width = 470
        Height = 25
        Align = alTop
        Caption = 'Executantes Dispon'#237'veis para Programa'#231#227'o'
        Color = clGradientActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 0
        ExplicitWidth = 468
      end
      object StatusBarDisponivelExecutantes: TStatusBar
        Left = 1
        Top = 267
        Width = 470
        Height = 19
        Panels = <
          item
            Width = 50
          end
          item
            Width = 50
          end>
        ExplicitTop = 259
        ExplicitWidth = 468
      end
      object RLExecutantes: TStringGrid
        Left = 1
        Top = 55
        Width = 470
        Height = 212
        Align = alClient
        ColCount = 11
        DefaultRowHeight = 22
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goFixedRowClick]
        TabOrder = 2
        OnDrawCell = RLExecutantesDrawCell
        OnFixedCellClick = RLExecutantesFixedCellClick
        OnKeyPress = RLExecutantesKeyPress
        OnMouseDown = RLExecutantesMouseDown
        OnSelectCell = RLExecutantesSelectCell
        ExplicitWidth = 468
        ExplicitHeight = 204
      end
      object ToolBar2: TToolBar
        Left = 1
        Top = 26
        Width = 470
        Height = 29
        ButtonWidth = 24
        Caption = 'ToolBar2'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 3
        ExplicitWidth = 468
        object ToolButton5: TToolButton
          Left = 0
          Top = 0
          Action = actSelecaoExecutantes
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton6: TToolButton
          Left = 24
          Top = 0
          Action = actSelecaoExecutantesLimpar
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton12: TToolButton
          Left = 48
          Top = 0
          Action = actExecutantesNaoEncontrados
          ParentShowHint = False
          ShowHint = True
        end
        object edtLocalizar: TEdit
          Left = 72
          Top = 0
          Width = 273
          Height = 22
          Hint = 'Buscar executante pelo Nome ou Fun'#231#227'o'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
          TextHint = 'Buscar executante pelo Nome ou Fun'#231#227'o'
          OnKeyPress = edtLocalizarKeyPress
        end
      end
      object ComboBoxOrigem: TComboBox
        Left = 256
        Top = 101
        Width = 89
        Height = 22
        Hint = 'Selecione a origem do executante'
        AutoDropDown = True
        BevelInner = bvSpace
        BevelKind = bkFlat
        Style = csOwnerDrawFixed
        DragMode = dmAutomatic
        DropDownCount = 30
        ParentShowHint = False
        ShowHint = True
        TabOrder = 4
        OnCloseUp = ComboBoxOrigemCloseUp
        OnMouseLeave = ComboBoxOrigemMouseLeave
      end
    end
    object TPanel
      Left = 1
      Top = 30
      Width = 472
      Height = 68
      Align = alTop
      AutoSize = True
      TabOrder = 2
      ExplicitWidth = 470
      object PanelTipoEtapaServico: TPanel
        Left = 1
        Top = 45
        Width = 470
        Height = 22
        Align = alTop
        TabOrder = 0
        ExplicitWidth = 468
        object Panel5: TPanel
          Left = 1
          Top = 1
          Width = 150
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Tipo de Etapa de Servico:'
          ParentBackground = False
          TabOrder = 0
        end
        object ComboBoxTipoEtapaServico: TComboBox
          Left = 151
          Top = 1
          Width = 318
          Height = 22
          Hint = 'Selecione o Tipo de Etapa de Servi'#231'o'
          Align = alClient
          AutoDropDown = True
          BevelInner = bvSpace
          BevelKind = bkFlat
          Style = csOwnerDrawFixed
          DragMode = dmAutomatic
          DropDownCount = 30
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnKeyPress = ComboBoxDestinoKeyPress
          ExplicitWidth = 316
        end
      end
      object PanelDestino: TPanel
        Left = 1
        Top = 23
        Width = 470
        Height = 22
        Align = alTop
        TabOrder = 1
        ExplicitWidth = 468
        object Panel2: TPanel
          Left = 1
          Top = 1
          Width = 150
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Destino:'
          ParentBackground = False
          TabOrder = 0
        end
        object ComboBoxDestino: TComboBox
          Left = 151
          Top = 1
          Width = 318
          Height = 22
          Hint = 'Selecione o Destino da Programa'#231#227'o'
          Align = alClient
          AutoDropDown = True
          BevelInner = bvSpace
          BevelKind = bkFlat
          Style = csOwnerDrawFixed
          CharCase = ecUpperCase
          DragMode = dmAutomatic
          DropDownCount = 30
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnKeyPress = ComboBoxDestinoKeyPress
          ExplicitWidth = 316
        end
      end
      object PanelDataProgramacao: TPanel
        Left = 1
        Top = 1
        Width = 470
        Height = 22
        Align = alTop
        TabOrder = 2
        ExplicitWidth = 468
        object Panel3: TPanel
          Left = 1
          Top = 1
          Width = 150
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '   Data programa'#231#227'o:'
          ParentBackground = False
          TabOrder = 0
        end
        object DateTimePickerProgramacao2: TDateTimePicker
          Left = 151
          Top = 1
          Width = 318
          Height = 20
          Hint = 'Selecione a data da programa'#231#227'o para gravar'
          Align = alClient
          Date = 42651.000000000000000000
          Time = 0.681124687500414400
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
          OnCloseUp = DateTimePickerProgramacao2CloseUp
          OnChange = DateTimePickerProgramacao2Change
          ExplicitWidth = 316
        end
      end
    end
    object PanelSelecaoServico: TPanel
      Left = 1
      Top = 388
      Width = 472
      Height = 220
      Align = alBottom
      TabOrder = 3
      ExplicitTop = 380
      ExplicitWidth = 470
      object ToolBar6: TToolBar
        Left = 1
        Top = 26
        Width = 470
        Height = 29
        ButtonWidth = 24
        Caption = 'ToolBar2'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 0
        ExplicitWidth = 468
        object ToolButton3: TToolButton
          Left = 0
          Top = 0
          Action = actSelecaoServicos
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton4: TToolButton
          Left = 24
          Top = 0
          Action = actSelecaoServicosLimpar
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton2: TToolButton
          Left = 48
          Top = 0
          Action = actInserirServico
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton1: TToolButton
          Left = 72
          Top = 0
          Action = actEditarServico
          ParentShowHint = False
          ShowHint = True
        end
        object ToolButton11: TToolButton
          Left = 96
          Top = 0
          Action = actServicosNaoEncontrados
          ParentShowHint = False
          ShowHint = True
        end
      end
      object Panel9: TPanel
        Left = 1
        Top = 1
        Width = 470
        Height = 25
        Align = alTop
        Caption = 'Servi'#231'os Dispon'#237'veis para Programa'#231#227'o'
        Color = clGradientActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 1
        ExplicitWidth = 468
      end
      object StatusBarDisponivelServicos: TStatusBar
        Left = 1
        Top = 200
        Width = 470
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
          end>
        ExplicitWidth = 468
      end
      object RLServicos: TStringGrid
        Left = 1
        Top = 55
        Width = 470
        Height = 145
        Align = alClient
        ColCount = 10
        DefaultRowHeight = 22
        FixedCols = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goDrawFocusSelected, goColSizing, goFixedRowClick]
        TabOrder = 3
        OnDrawCell = RLServicosDrawCell
        OnFixedCellClick = RLServicosFixedCellClick
        OnMouseDown = RLServicosMouseDown
        ExplicitWidth = 468
      end
    end
  end
  object PanelAjuda: TPanel
    Left = 219
    Top = 151
    Width = 337
    Height = 338
    TabOrder = 3
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
    object PanelAjudaInserirServico: TPanel
      Left = 1
      Top = 201
      Width = 335
      Height = 156
      Align = alTop
      TabOrder = 1
      object Panel23: TPanel
        Left = 1
        Top = 1
        Width = 333
        Height = 22
        Align = alTop
        TabOrder = 0
        object Panel24: TPanel
          Left = 1
          Top = 1
          Width = 140
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Servi'#231'o/Ano-Etapa:'
          ParentBackground = False
          TabOrder = 0
        end
        object edtServicoPT: TEdit
          Left = 141
          Top = 1
          Width = 191
          Height = 20
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 21
        end
      end
      object Panel25: TPanel
        Left = 1
        Top = 45
        Width = 333
        Height = 22
        Align = alTop
        TabOrder = 1
        object Panel26: TPanel
          Left = 1
          Top = 1
          Width = 140
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Texto Breve da Ordem:'
          ParentBackground = False
          TabOrder = 0
        end
        object edtTextoBreveOM: TEdit
          Left = 141
          Top = 1
          Width = 191
          Height = 20
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 21
        end
      end
      object BitBtn1: TBitBtn
        Left = 8
        Top = 126
        Width = 80
        Height = 24
        Action = actServicoInserir
        Caption = 'Inserir'
        ImageIndex = 194
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF000080000000800000FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF0000800000008000000080000000800000FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF0000800000008000000080000000800000FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF0000800000008000000080000000800000008000000080
          0000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF000080000000800000008000000080000000800000008000000080
          0000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF0000800000008000000080000000800000FF00FF0000800000008000000080
          000000800000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF000080
          00000080000000800000FF00FF00FF00FF00FF00FF00FF00FF00008000000080
          000000800000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00008000000080
          00000080000000800000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF000080
          00000080000000800000FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00008000000080000000800000FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00008000000080000000800000FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00008000000080000000800000FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
        ParentShowHint = False
        ShowHint = True
        TabOrder = 2
      end
      object BitBtn2: TBitBtn
        Left = 94
        Top = 126
        Width = 80
        Height = 24
        Action = actAjudaLimpar
        Caption = 'Cancelar'
        ImageIndex = 190
        Glyph.Data = {
          36040000424D3604000000000000360000002800000010000000100000000100
          2000000000000004000000000000000000000000000000000000FF00FF00FF00
          FF00BFC6CA00B5BCC000BEC5C900FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00C0C8CC00B7BEC200FF00FF00FF00FF00BFC6
          CA008A909200565A5B007A7E8100BBC3C700FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00C7CFD300B3BABE0083888B0071757800B7BEC200FF00FF007665
          D7002900DF0036219F004F5254009EA4A700C6CED200FF00FF00FF00FF00FF00
          FF00C6CED200A7AEB1005141A5002C06D70083888B00C0C8CC00FF00FF00968F
          D7002900DF002E08DA0042405D0064686A00ADB4B800FF00FF00FF00FF00C6CE
          D2009DA1B2003D21BC002900DF006657BB00BCC4C800FF00FF00FF00FF00FF00
          FF006F5BDA002900DF003717CD004B4D4F0074797B00B0B8BB00C6CED2009DA1
          B2003414C9002C06D700827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00654EDB002900DF00462EB9004E525300737779009093A3003414
          C9002C06D700827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00654EDB002900DF00452DB90041415000300FC4002B05
          D600827FAC00C1C9CD00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00644DDA002900DF002B05D6002A04D5005550
          7B00ABB2B500C7CFD300FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00C6CED2009DA1B2002D07D8002900DF003E23BD005B5F
          600072767900A7AEB100C3CBCF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00C6CED2009DA1B2003414C9002C06D7005D4BBC00320CDD004C32
          CC0062647400686C6F00999FA200BFC6CA00FF00FF00FF00FF00FF00FF00FF00
          FF00C3CBCF009DA1B2003414C9002900DF006F68A300BCC4C800AAA9D6003C19
          DD004121D700615D8900616567008C929500C0C8CC00FF00FF00FF00FF00C5CD
          D1009497A8003413C8002900DF005D51A600B8BFC300FF00FF00FF00FF00BEC3
          D5005134DC00320BDD00635AA1006B707200ACB3B600C6CED200FF00FF009D9B
          C8003413C8002900DF004B39AA00B2B9BD00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF007868DA002900DF00A4A5C400C6CED200FF00FF00FF00FF003B18
          DB002900DF003D21BC00A7AEB100C7CFD300FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00330D
          DE004424DA00B0B4C600C7CFD300FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
          FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
        ParentShowHint = False
        ShowHint = True
        TabOrder = 3
      end
      object Panel4: TPanel
        Left = 1
        Top = 89
        Width = 333
        Height = 22
        Align = alTop
        TabOrder = 4
        object Panel11: TPanel
          Left = 1
          Top = 1
          Width = 140
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Opera'#231#227'o:'
          ParentBackground = False
          TabOrder = 0
        end
        object edtOperacao: TEdit
          Left = 141
          Top = 1
          Width = 191
          Height = 20
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 21
        end
      end
      object Panel1: TPanel
        Left = 1
        Top = 67
        Width = 333
        Height = 22
        Align = alTop
        TabOrder = 5
        object Panel8: TPanel
          Left = 1
          Top = 1
          Width = 140
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Ordem de manuten'#231#227'o:'
          ParentBackground = False
          TabOrder = 0
        end
        object edtOrdemManutencao: TEdit
          Left = 141
          Top = 1
          Width = 191
          Height = 20
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 21
        end
      end
      object Panel10: TPanel
        Left = 1
        Top = 23
        Width = 333
        Height = 22
        Align = alTop
        TabOrder = 6
        object Panel12: TPanel
          Left = 1
          Top = 1
          Width = 140
          Height = 20
          Align = alLeft
          Alignment = taLeftJustify
          Caption = '  Texto Breve da Operacao:'
          ParentBackground = False
          TabOrder = 0
        end
        object edtTextoBreveOP: TEdit
          Left = 141
          Top = 1
          Width = 191
          Height = 20
          Align = alClient
          TabOrder = 1
          ExplicitHeight = 21
        end
      end
    end
    object PanelNaoEncontrados: TPanel
      Left = 1
      Top = 357
      Width = 335
      Height = 2
      Align = alClient
      TabOrder = 2
      object ToolBar5: TToolBar
        Left = 1
        Top = 1
        Width = 333
        Height = 29
        Caption = 'ToolBar5'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        TabOrder = 0
        object BitBtn4: TBitBtn
          Left = 0
          Top = 0
          Width = 75
          Height = 22
          Action = actExcelNaoEncontrado
          Caption = 'Exportar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 0
        end
      end
      object RLNaoEncontrados: TStringGrid
        Left = 1
        Top = 30
        Width = 333
        Height = 20
        Align = alClient
        ColCount = 11
        DefaultRowHeight = 22
        FixedCols = 0
        RowCount = 1
        FixedRows = 0
        Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goAlwaysShowEditor, goFixedRowClick]
        TabOrder = 1
        OnFixedCellClick = RLNaoEncontradosFixedCellClick
      end
    end
    object PanelLogAcao: TPanel
      Left = 1
      Top = 27
      Width = 335
      Height = 174
      Align = alTop
      TabOrder = 3
      object ToolBar8: TToolBar
        Left = 1
        Top = 1
        Width = 333
        Height = 29
        Caption = 'ToolBar8'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        TabOrder = 0
      end
      object DBMemoLogAcao: TDBMemo
        Left = 1
        Top = 30
        Width = 333
        Height = 143
        Align = alClient
        DataField = 'LogAcao'
        DataSource = FrmDataModule.DataSourceProgramacaoDiaria_Cadastro
        ReadOnly = True
        ScrollBars = ssBoth
        TabOrder = 1
        WordWrap = False
      end
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 56
    Top = 161
    StyleName = 'Platform Default'
    object actConfigGRID: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actConfigGRID'
      OnExecute = actConfigGRIDExecute
    end
    object actProcurarProgramacao: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      ShortCut = 116
      OnExecute = actProcurarProgramacaoExecute
    end
    object actCarregarExecutantes: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actCarregarExecutantes'
      OnExecute = actCarregarExecutantesExecute
    end
    object actAtualizarCombo: TAction
      Category = 'Configura'#231#245'es'
      Hint = 'Atualizar combobox de Destino e Tipo de Servi'#231'o'
      ImageIndex = 28
      OnExecute = actAtualizarComboExecute
    end
    object actInserirServico: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actInserirServico'
      Hint = 'Inserir um novo servi'#231'o'
      ImageIndex = 449
      OnExecute = actInserirServicoExecute
    end
    object actAjudaLimpar: TAction
      Category = 'Servi'#231'o'
      Caption = 'Cancelar'
      Hint = 'Cancelar opera'#231#227'o'
      ImageIndex = 188
      ShortCut = 8219
      OnExecute = actAjudaLimparExecute
    end
    object actInserirProgramacao: TAction
      Category = 'Inserir Programa'#231#227'o'
      Caption = 'Inserir'
      Hint = 'Inserir programa'#231#227'o'
      ImageIndex = 2
      OnExecute = actInserirProgramacaoExecute
    end
    object actGravarProgramacao: TAction
      Category = 'Inserir Programa'#231#227'o'
      Caption = 'Gravar'
      Hint = 'Gravar Programa'#231#227'o'
      ImageIndex = 18
      ShortCut = 119
      OnExecute = actGravarProgramacaoExecute
    end
    object actcalcelarProgramacao: TAction
      Category = 'Inserir Programa'#231#227'o'
      Caption = 'Cancelar'
      Hint = 'Cancelar altera'#231#245'es'
      ImageIndex = 188
      ShortCut = 8219
      OnExecute = actcalcelarProgramacaoExecute
    end
    object actInserirProgramacaoBranco: TAction
      Category = 'Programados'
      Caption = 'Inserir'
      Hint = 'Inserir programa'#231#227'o'
      ImageIndex = 2
      OnExecute = actInserirProgramacaoBrancoExecute
    end
    object actEditarProgramacao: TAction
      Category = 'Programados'
      Caption = 'Editar'
      Hint = 'Editar programa'#231#227'o para edi'#231#227'o'
      ImageIndex = 20
      ShortCut = 118
      OnExecute = actEditarProgramacaoExecute
    end
    object actExcluirProgramacao: TAction
      Category = 'Programados'
      Caption = 'Excluir selecionado'
      Hint = 'Excluir programa'#231#227'o selecionada'
      ImageIndex = 324
      ShortCut = 46
      OnExecute = actExcluirProgramacaoExecute
    end
    object actExcluirProgramacaoTodos: TAction
      Category = 'Programados'
      Caption = 'Excluir todos filtrados'
      ImageIndex = 324
      OnExecute = actExcluirProgramacaoTodosExecute
    end
    object actCopiarProgramacaoSELECAO: TAction
      Category = 'Programados'
      Caption = 'Copiar a programa'#231#227'o selecionada'
      Hint = 'Copiar a programa'#231#227'o selecionada'
      ImageIndex = 138
      OnExecute = actCopiarProgramacaoSELECAOExecute
    end
    object actExecutantesServicos: TAction
      Category = 'Inserir Programa'#231#227'o'
      Caption = 'Executantes e Servi'#231'os'
      Hint = 'Listar Executantes e Servi'#231'os dipon'#237'veis'
      ImageIndex = 0
      ShortCut = 116
      OnExecute = actExecutantesServicosExecute
    end
    object actConfigBotoes1: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actConfigBotoes1'
      OnExecute = actConfigBotoes1Execute
    end
    object actEditarServico: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actEditarServico'
      Hint = 'Editar Servi'#231'o selecionado'
      ImageIndex = 198
      OnExecute = actEditarServicoExecute
    end
    object actServicoInserir: TAction
      Category = 'Servi'#231'o'
      Caption = 'Inserir'
      Hint = 'Inserir'
      ImageIndex = 192
      OnExecute = actServicoInserirExecute
    end
    object actServicoEditar: TAction
      Category = 'Servi'#231'o'
      Caption = 'Gravar'
      Hint = 'Gravar Servi'#231'o selecionado'
      ImageIndex = 18
      OnExecute = actServicoEditarExecute
    end
    object actNumSelecao: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actNumSelecao'
      OnExecute = actNumSelecaoExecute
    end
    object actSelecaoServicos: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'actSelecaoServicos'
      Hint = 'Selecionar todos'
      ImageIndex = 232
      OnExecute = actSelecaoServicosExecute
    end
    object actSelecaoServicosLimpar: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'actSelecaoServicosLimpar'
      Hint = 'Limpar todas as sele'#231#245'es'
      ImageIndex = 231
      OnExecute = actSelecaoServicosLimparExecute
    end
    object actSelecaoExecutantes: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'actSelecaoExecutantes'
      Hint = 'Selecionar todos'
      ImageIndex = 232
      OnExecute = actSelecaoExecutantesExecute
    end
    object actSelecaoExecutantesLimpar: TAction
      Category = 'Sele'#231#227'o'
      Caption = 'actSelecaoExecutantesLimpar'
      Hint = 'Limpar todas as sele'#231#245'es'
      ImageIndex = 231
      OnExecute = actSelecaoExecutantesLimparExecute
    end
    object actExcelNaoEncontrado: TAction
      Category = 'Excel'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelNaoEncontradoExecute
    end
    object actNaoEncontrados: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Lista de execunates n'#227'o encontrados na importa'#231#227'o'
      Hint = 'Lista de execunates n'#227'o encontrados na importa'#231#227'o'
      ImageIndex = 119
      OnExecute = actNaoEncontradosExecute
    end
    object actZoomMais: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actZoomMais'
      Hint = 'Zoom Mais'
      ImageIndex = 46
      OnExecute = actZoomMaisExecute
    end
    object actZoomMenos: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actZoomMenos'
      Hint = 'Zoom Menos'
      ImageIndex = 47
      OnExecute = actZoomMenosExecute
    end
    object actAutoFit: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actAutoFit'
      OnExecute = actAutoFitExecute
    end
    object actCopiarProgramacaoTODAS: TAction
      Category = 'Programados'
      Caption = 'Copiar todas as programa'#231#245'es filtradas'
      Hint = 'Copiar todas as programa'#231#245'es filtradas'
      ImageIndex = 138
      OnExecute = actCopiarProgramacaoTODASExecute
    end
    object actServicoInserir1: TAction
      Category = 'Alterar Registros'
      Caption = 'Inserir'
      Hint = 'Inserir servi'#231'o'
      ImageIndex = 449
      OnExecute = actServicoInserir1Execute
    end
    object actServicoExcluir: TAction
      Category = 'Alterar Registros'
      Caption = 'Excluir'
      Hint = 'Excluir servi'#231'o selecionado'
      ImageIndex = 450
      OnExecute = actServicoExcluirExecute
    end
    object actExecutanteExcluir: TAction
      Category = 'Alterar Registros'
      Caption = 'Excluir'
      Hint = 'Excliur executante selecionado'
      ImageIndex = 450
      OnExecute = actExecutanteExcluirExecute
    end
    object actServicosNaoEncontrados: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actServicosNaoEncontrados'
      Hint = 'Lista de Servi'#231'os n'#227'o encontrados'
      ImageIndex = 22
      OnExecute = actServicosNaoEncontradosExecute
    end
    object actExecutantesNaoEncontrados: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actExecutantesNaoEncontrados'
      Hint = 'Lista de Executantes n'#227'o encontrados'
      ImageIndex = 22
      OnExecute = actExecutantesNaoEncontradosExecute
    end
    object actLogAcao: TAction
      Category = 'Programados'
      Caption = 'Hist'#243'rico'
      Hint = 'Log de a'#231#227'o (Hist'#243'rico)'
      ImageIndex = 124
      OnExecute = actLogAcaoExecute
    end
    object actCarregarServico: TAction
      Category = 'Executantes e Servi'#231'os'
      Caption = 'actCarregarServico'
      OnExecute = actCarregarServicoExecute
    end
    object actAtualizar: TAction
      Category = 'Inserir Programa'#231#227'o'
      Caption = 'Atualizar'
      Hint = 'Atualizar tabela de Executantes'
      ImageIndex = 288
      OnExecute = actAtualizarExecute
    end
    object actProcurarExecutantes: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecutantesExecute
    end
    object actProcurarServicos: TAction
      Category = 'Procurar'
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarServicosExecute
    end
  end
  object PopupMenuExcluir: TPopupMenu
    Left = 49
    Top = 338
    object Excluirselecionado1: TMenuItem
      Action = actExcluirProgramacao
    end
    object Excluirtodos1: TMenuItem
      Action = actExcluirProgramacaoTodos
    end
  end
  object PopupMenuCopiar: TPopupMenu
    Left = 81
    Top = 338
    object Importarselecionado1: TMenuItem
      Action = actCopiarProgramacaoSELECAO
    end
    object Importarsemverificao1: TMenuItem
      Action = actCopiarProgramacaoTODAS
    end
  end
end
