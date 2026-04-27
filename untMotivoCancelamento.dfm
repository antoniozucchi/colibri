object FrmMotivoCancelamento: TFrmMotivoCancelamento
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Motivos de cancelamento'
  ClientHeight = 441
  ClientWidth = 624
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object PanelCancelamentoMudanca: TPanel
    Left = 0
    Top = 0
    Width = 624
    Height = 441
    Align = alClient
    TabOrder = 0
    ExplicitWidth = 622
    ExplicitHeight = 433
    object PanelMotivoCancelamento: TPanel
      Left = 1
      Top = 1
      Width = 622
      Height = 439
      Align = alClient
      TabOrder = 0
      ExplicitWidth = 620
      ExplicitHeight = 431
      object DBGridPalavraChave: TFilterDBGrid
        Left = 1
        Top = 55
        Width = 620
        Height = 364
        Align = alClient
        DataSource = FrmDataModule.DataSourcePalavraChave
        TabOrder = 0
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -12
        TitleFont.Name = 'Segoe UI'
        TitleFont.Style = []
        ClearFilterButton = btnClearFiltro
        SearchAction = actProcurar
        EnableZebra = False
        ExcelButton = btnExcel
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
            FieldName = 'PalavraChave'
            Title.Alignment = taCenter
            Title.Caption = 'Ocorr'#234'ncia'
            Width = 300
            Visible = True
          end>
      end
      object ToolBar9: TToolBar
        Left = 1
        Top = 26
        Width = 620
        Height = 29
        Caption = 'ToolBar9'
        DrawingStyle = dsGradient
        EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
        Images = FrmPrincipal.ImageList1
        TabOrder = 2
        ExplicitTop = 1
        ExplicitWidth = 618
        object DBNavigatorPalavraChave: TDBNavigator
          Left = 0
          Top = 0
          Width = 138
          Height = 22
          DataSource = FrmDataModule.DataSourcePalavraChave
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
        object btnExcel: TToolButton
          Left = 161
          Top = 0
          Hint = 'Exportar dados para o Excel'
          Caption = 'btnExcel'
          ImageIndex = 54
          ParentShowHint = False
          ShowHint = True
        end
        object BitBtn15: TBitBtn
          Left = 184
          Top = 0
          Width = 64
          Height = 22
          Action = actInserir
          Align = alLeft
          Caption = 'Inserir'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 2
        end
        object BitBtn14: TBitBtn
          Left = 248
          Top = 0
          Width = 81
          Height = 22
          Action = actCancelar
          Align = alLeft
          Caption = 'Fechar'
          ParentShowHint = False
          ShowHint = True
          TabOrder = 1
        end
      end
      object StatusBarPalavraChave: TStatusBar
        Left = 1
        Top = 419
        Width = 620
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
          end>
        ExplicitTop = 411
        ExplicitWidth = 618
      end
      object RadioGroupFonteCancelamento: TRadioGroup
        Left = 28
        Top = 282
        Width = 274
        Height = 98
        Caption = 'RadioGroupFonteCancelamento'
        Items.Strings = (
          'Gerenciar Solicita'#231#245'es (Cancelamento)'
          'Consulta Executante (N'#227'o Executado)'
          'Gerenciar Solicita'#231#245'es (Mudan'#231'a)'
          'Programa'#231#227'o FIFI')
        TabOrder = 4
        Visible = False
      end
      object PanelTitulo: TPanel
        Left = 1
        Top = 1
        Width = 620
        Height = 25
        Align = alTop
        Caption = 'Executantes Programados por Per'#237'odo'
        Color = clGradientActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -13
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentBackground = False
        ParentFont = False
        TabOrder = 5
        ExplicitLeft = 2
        ExplicitTop = 9
      end
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 304
    Top = 224
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
    object actInserir: TAction
      Caption = 'Inserir'
      Hint = 'Inserir registro'
      ImageIndex = 2
      OnExecute = actInserirExecute
    end
    object actCancelar: TAction
      Caption = 'Fechar'
      Hint = 'Cancelar'
      ImageIndex = 188
      OnExecute = actCancelarExecute
    end
  end
end
