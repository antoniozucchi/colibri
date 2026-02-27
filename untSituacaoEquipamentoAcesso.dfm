object FrmSituacaoEquipamentoAcesso: TFrmSituacaoEquipamentoAcesso
  Left = 0
  Top = 0
  Caption = 'Situa'#231#227'o dos Equipamentos e Acessos das Plataformas'
  ClientHeight = 362
  ClientWidth = 823
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
    Top = 232
    Width = 823
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
    Width = 823
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    TabOrder = 0
    ExplicitWidth = 821
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
      TabOrder = 3
    end
    object BitBtn4: TBitBtn
      Left = 132
      Top = 0
      Width = 75
      Height = 22
      Action = actExcel
      Align = alLeft
      Caption = 'Exportar'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
    object BitBtn2: TBitBtn
      Left = 207
      Top = 0
      Width = 64
      Height = 22
      Action = actLimparFiltros
      Align = alLeft
      Caption = 'Limpar'
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FDFDFD029B9B9B643434
        34CB070707F80D0D0DF2474747B8BDBDBD42FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FBFBFB04515151AE000000FF0000
        00FF0C0C0CF3070707F8000000FF010101FE89898976FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF008181817E000000FF232323DCC1C1
        C13EFEFEFE01F9F9F906A1A1A15E0B0B0BF4010101FEBDBDBD42FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00F9F9F906101010EF050505FAD9D9D926FF00
        FF00FF00FF00FF00FF00FF00FF00A1A1A15E000000FF474747B8FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00CFCFCF30000000FF3F3F3FC0FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00F9F9F906070707F80D0D0DF2FF00FF00FF00
        FF00FF00FF00FF00FF008B8B8B74181818E7000000FF0A0A0AF5282828D7FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00151515EA000000FFFF00FF00FF00
        FF00FF00FF00FF00FF007B7B7B84000000FF000000FF000000FF0B0B0BF4FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00151515EA000000FFFF00FF00FDFD
        FD02C5C5C53AAEAEAE51DBDBDB24F5F5F50AF5F5F50AF5F5F50AF2F2F20DB8B8
        B847B2B2B24DE6E6E619FF00FF00FF00FF00151515EA000000FFE6E6E6194D4D
        4DB28F8F8F70A5A5A55A484848B7DCDCDC23FF00FF00FF00FF007474748B9696
        9669A1A1A15E575757A89191916EFF00FF00151515EA000000FF5F5F5FA0B0B0
        B04F1E1E1EE1030303FCA5A5A55A484848B7F9F9F906B1B1B14E9292926D3E3E
        3EC1050505FA787878875F5F5FA0CFCFCF30151515EA000000FF595959A65757
        57A8000000FF000000FF1E1E1EE1A6A6A659494949B67575758A8A8A8A750000
        00FF000000FF010101FEA6A6A65988888877151515EA000000FF5B5B5BA45252
        52AD000000FF000000FF000000FF292929D6878787786262629D020202FD0000
        00FF000000FF000000FFA6A6A65986868679151515EA000000FF5B5B5BA4ADAD
        AD52050505FA000000FF000000FF000000FF000000FF000000FF000000FF0000
        00FF000000FF4E4E4EB179797986C5C5C53A232323DC131313ECD5D5D52A5757
        57A8B3B3B34C9292926D8F8F8F708F8F8F708F8F8F708F8F8F708F8F8F708F8F
        8F70A3A3A35C9191916E6B6B6B94FF00FF005C5C5CA35C5C5CA3FF00FF00E4E4
        E41B777777885D5D5DA25D5D5DA25D5D5DA25D5D5DA25D5D5DA25D5D5DA25D5D
        5DA25D5D5DA2A6A6A659FEFEFE01FF00FF005C5C5CA35C5C5CA3FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF002E2E2ED1212121DE}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
    end
    object BitBtn1: TBitBtn
      Left = 271
      Top = 0
      Width = 64
      Height = 22
      Action = actFiltrosTabela
      Align = alLeft
      Caption = 'Filtros'
      Glyph.Data = {
        36040000424D3604000000000000360000002800000010000000100000000100
        2000000000000004000000000000000000000000000000000000FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FCFC
        FC03DDDDDD25C6C6C63F9F9F9F6A878686866F6E6EA15B5A5AB794949477E2E2
        E21FFDFDFD02C4C4C441DEDEDE24FF00FF00FF00FF00FF00FF00E2E2E21F6766
        66A99D9D9D6D9B9B9B6F87878785373636DF292828EF232222F51C1B1BFF1B1A
        1AFF383737DE252424F3403F3FD48281818BFF00FF00FF00FF00FEFEFE019594
        94768585858796959575515050C28181818CE7E7E71AC0C0C0451C1B1BFD1C1B
        1BFF1C1B1BFF1C1B1BFF1C1B1BFF3F3E3ED6DBDBDB27FF00FF00DDDDDD258D8C
        8C7FAAA9A95FA6A6A662D0D0D034FF00FF00FF00FF008080808D1C1B1BFF1C1B
        1BFF1C1B1BFF1C1B1BFF1C1B1BFF282727F05A5959B7E1E0E021FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00E9E9E918A6A6A6626261
        61AF232222F61C1B1BFF292828EF222121F71C1B1BFFC9C9C93BFF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00F0F0F010ADACAC5B89898982A8A8A860686868A8EEEEEE12FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00
        FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00}
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
    end
  end
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 823
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
    ExplicitWidth = 821
  end
  object DBGridSituacaoEquipamentoAcesso: TDBGrid
    Left = 0
    Top = 54
    Width = 823
    Height = 178
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
    OnTitleClick = DBGridSituacaoEquipamentoAcessoTitleClick
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
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoLinhaBCI'
        Title.Alignment = taCenter
        Title.Caption = 'Linha BCI'
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoAgua'
        Title.Alignment = taCenter
        Title.Caption = 'Ab.'#193'gua'
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoBalsa'
        Title.Alignment = taCenter
        Title.Caption = 'Balsa'
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoAcesso'
        Title.Alignment = taCenter
        Title.Caption = 'Int.Acesso'
        Width = 75
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'SituacaoDegraus'
        Title.Alignment = taCenter
        Title.Caption = 'Limp.Degraus'
        Width = 75
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
    Top = 343
    Width = 823
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
    ExplicitTop = 335
    ExplicitWidth = 821
  end
  object ColunasLayout: TStringGrid
    Left = 48
    Top = 102
    Width = 113
    Height = 94
    ColCount = 2
    DefaultRowHeight = 21
    TabOrder = 4
    Visible = False
    RowHeights = (
      21
      21
      21
      21
      21)
  end
  object Panel2: TPanel
    Left = 0
    Top = 235
    Width = 823
    Height = 108
    Align = alBottom
    Caption = 'Panel2'
    TabOrder = 5
    ExplicitTop = 227
    ExplicitWidth = 821
    object PanelTituloNotas: TPanel
      Left = 1
      Top = 1
      Width = 821
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
      ExplicitWidth = 819
    end
    object DBMemo1: TDBMemo
      Left = 1
      Top = 21
      Width = 821
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
      ExplicitWidth = 819
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 312
    Top = 150
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'actProcurar'
      OnExecute = actProcurarExecute
    end
    object actExcel: TAction
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelExecute
    end
    object actFiltroInserir: TAction
      Category = 'Tabela'
      Caption = 'actFiltroInserir'
      OnExecute = actFiltroInserirExecute
    end
    object actGridASC: TAction
      Category = 'Tabela'
      Caption = 'actGridASC'
      OnExecute = actGridASCExecute
    end
    object actGridDESC: TAction
      Category = 'Tabela'
      Caption = 'actGridDESC'
      OnExecute = actGridDESCExecute
    end
    object actSubstituirPor: TAction
      Category = 'Tabela'
      Caption = 'actSubstituirPor'
      OnExecute = actSubstituirPorExecute
    end
    object actLimparFiltros: TAction
      Category = 'Tabela'
      Caption = 'Limpar'
      Hint = 'Limpar Filtros'
      ImageIndex = 469
      OnExecute = actLimparFiltrosExecute
    end
    object actFiltrosTabela: TAction
      Category = 'Tabela'
      Caption = 'Filtros'
      Hint = 'Visualizar tabela de filtros'
      ImageIndex = 470
      OnExecute = actFiltrosTabelaExecute
    end
    object actProcuraFiltrosTabela: TAction
      Category = 'Tabela'
      Caption = 'actProcuraFiltrosTabela'
      OnExecute = actProcuraFiltrosTabelaExecute
    end
  end
end
