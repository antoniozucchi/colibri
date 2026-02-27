object FrmPlataforma: TFrmPlataforma
  Left = 0
  Top = 0
  Caption = 'Cadastro de Instala'#231#227'o'
  ClientHeight = 411
  ClientWidth = 1182
  Color = clMenu
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Arial'
  Font.Style = []
  FormStyle = fsMDIChild
  Position = poScreenCenter
  Visible = True
  WindowState = wsMaximized
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  TextHeight = 14
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 1182
    Height = 25
    Align = alTop
    Caption = 'Cadastro de Instala'#231#227'o'
    Color = clGradientInactiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
  end
  object ColunasLayout: TStringGrid
    Left = 41
    Top = 209
    Width = 113
    Height = 94
    ColCount = 2
    DefaultRowHeight = 21
    TabOrder = 1
    Visible = False
    RowHeights = (
      21
      21
      21
      21
      21)
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 392
    Width = 1182
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
  object DBGridPlataformas: TDBGrid
    Left = 0
    Top = 54
    Width = 1182
    Height = 338
    Align = alClient
    DataSource = FrmDataModule.DataSourcePlataforma
    TabOrder = 3
    TitleFont.Charset = ANSI_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Arial'
    TitleFont.Style = []
    OnCellClick = DBGridPlataformasCellClick
    OnDrawColumnCell = DBGridPlataformasDrawColumnCell
    OnTitleClick = DBGridPlataformasTitleClick
    Columns = <
      item
        Expanded = False
        FieldName = 'booleanPlataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Controle'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'booleanOrigem'
        Title.Alignment = taCenter
        Title.Caption = 'Origem'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Plataforma'
        Title.Alignment = taCenter
        Title.Caption = 'Instala'#231#227'o'
        Width = 162
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'Salvatagem'
        Title.Alignment = taCenter
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'ElevacaoTopDECK'
        Title.Alignment = taCenter
        Title.Caption = 'Eleva'#231#227'o conv'#233's superior (El) [cm]'
        Width = 200
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'InicioOperacao'
        Title.Alignment = taCenter
        Title.Caption = 'In'#237'cio da opera'#231#227'o'
        Width = 150
        Visible = True
      end
      item
        Alignment = taCenter
        Expanded = False
        FieldName = 'LaminaAgua'
        Title.Alignment = taCenter
        Title.Caption = 'L'#226'mina d'#39#225'gua [m]'
        Width = 150
        Visible = True
      end
      item
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
        ReadOnly = True
        Title.Alignment = taCenter
        Width = 70
        Visible = True
      end>
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 1182
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 4
    object DBNavigatorCadastro: TDBNavigator
      Left = 0
      Top = 0
      Width = 144
      Height = 22
      DataSource = FrmDataModule.DataSourcePlataforma
      VisibleButtons = [nbInsert, nbDelete, nbEdit, nbPost, nbCancel, nbRefresh]
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
    object ToolButton2: TToolButton
      Left = 144
      Top = 0
      Width = 8
      Caption = 'ToolButton2'
      ImageIndex = 58
      Style = tbsSeparator
    end
    object BitBtn3: TBitBtn
      Left = 152
      Top = 0
      Width = 70
      Height = 22
      Action = actExcel
      Align = alLeft
      Caption = 'Exportar'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 4
    end
    object BitBtn4: TBitBtn
      Left = 222
      Top = 0
      Width = 70
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
      TabOrder = 3
    end
    object BitBtn1: TBitBtn
      Left = 292
      Top = 0
      Width = 70
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
      TabOrder = 2
    end
    object BitBtn2: TBitBtn
      Left = 362
      Top = 0
      Width = 63
      Height = 22
      Action = actLayoutJanela
      Align = alLeft
      Caption = 'Layout'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 6
    end
    object BitBtn5: TBitBtn
      Left = 425
      Top = 0
      Width = 97
      Height = 22
      Action = actCalcularXY
      Align = alLeft
      Caption = 'Calcular (X,Y)'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 5
    end
    object DBEdit1: TDBEdit
      Left = 522
      Top = 0
      Width = 121
      Height = 22
      DataField = 'idPlataforma'
      DataSource = FrmDataModule.DataSourcePlataforma
      TabOrder = 1
      Visible = False
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 120
    Top = 122
    StyleName = 'Platform Default'
    object actPanTo: TAction
      Category = 'Ferramentas'
      Caption = 'actPanTo'
      Hint = 'Posicionar (Latitude, Longitude)'
      ImageIndex = 290
    end
    object actImprimir: TAction
      Category = 'Ferramentas'
      Caption = 'actImprimir'
      Hint = 'Imprimir o mapa'
      ImageIndex = 21
    end
    object actSalvar: TAction
      Category = 'Ferramentas'
      Caption = 'actSalvar'
      Hint = 'Salvar mapa como imagem'
      ImageIndex = 264
    end
    object actAtualizarTituo: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actAtualizarTituo'
      OnExecute = actAtualizarTituoExecute
    end
    object actMarcar: TAction
      Category = 'Ferramentas'
      Caption = 'actMarcar'
      Hint = 'Marcador'
      ImageIndex = 112
    end
    object actProcurar: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actProcurar'
      OnExecute = actProcurarExecute
    end
    object actProcurarTabuaMare: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'actProcurarTabuaMare'
    end
    object actCalcularXY: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Calcular (X,Y)'
      Hint = 'Calcular coordenadas X,Y com base nas Latitudes e Longitudes'
      ImageIndex = 51
      OnExecute = actCalcularXYExecute
    end
    object actZoomFit: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomFit'
      Hint = 'Zoom Fit'
      ImageIndex = 45
    end
    object actPan: TAction
      Category = 'Ferramentas'
      Caption = 'actPan'
      Hint = 'Pan'
      ImageIndex = 41
    end
    object actZoomDinamico: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomDinamico'
      Hint = 'Zoom din'#226'mico'
      ImageIndex = 48
    end
    object actZoomWindow: TAction
      Category = 'Ferramentas'
      Caption = 'actZoomWindow'
      Hint = 'Zoom janela'
      ImageIndex = 46
    end
    object actClearGrafico: TAction
      Category = 'Ferramentas'
      Caption = 'actClearGrafico'
      Hint = 'Limpar o Mapa'
      ImageIndex = 324
    end
    object actMapaMundi: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Mapa Mundi'
    end
    object actBrasil: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Brasil'
    end
    object actLatitudeLongitude: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Latitudes e Longitudes'
    end
    object actCarregarLinhas: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Carregar linhas 2D (x1;y1;x2;y2)...'
      Hint = 
        'Carregar linhas a partir de arquivo *.txt com as coordenadas das' +
        ' linhas no formato: x1;y1;x2;y2 '
    end
    object actAbrirDesenho: TAction
      Category = 'Ferramentas'
      Caption = 'actAbrirDesenho'
      Hint = 'Abrir arquivo de desenho 2D (*.zuchi)...'
      ImageIndex = 20
    end
    object actSalvarDesenhoData: TAction
      Category = 'Ferramentas'
      Caption = 'actSalvarDesenhoData'
      Hint = 'Salvar arquivo como desenho 2D (*.zuchi)...'
      ImageIndex = 462
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
    object actExcel: TAction
      Category = 'Excel'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
      OnExecute = actExcelExecute
    end
    object actTabuaMare: TAction
      Category = 'Configura'#231#245'es'
      Caption = 'Calcular'
      Hint = 'Consultar altura da mar'#233' agora'
      ImageIndex = 458
    end
    object actExcelTabuaMAre: TAction
      Category = 'Excel'
      Caption = 'Exportar'
      Hint = 'Exportar para o Excel'
      ImageIndex = 54
    end
    object actExcelImportarTabuaMare: TAction
      Category = 'Excel'
      Caption = 'Importar'
      Hint = 'Importar planilha do Excel'
      ImageIndex = 251
    end
    object actCimaTudo: TAction
      Category = 'Layout'
      Caption = 'actCimaTudo'
      OnExecute = actCimaTudoExecute
    end
    object actBaixoTudo: TAction
      Category = 'Layout'
      Caption = 'actBaixoTudo'
      OnExecute = actBaixoTudoExecute
    end
    object actBaixo: TAction
      Category = 'Layout'
      Caption = 'actBaixo'
      OnExecute = actBaixoExecute
    end
    object actCima: TAction
      Category = 'Layout'
      Caption = 'actCima'
      OnExecute = actCimaExecute
    end
    object actSalvarLayout: TAction
      Category = 'Layout'
      Caption = 'actSalvarLayout'
      OnExecute = actSalvarLayoutExecute
    end
    object actLayoutJanela: TAction
      Category = 'Layout'
      Caption = 'Layout'
      Hint = 'Layout da tabela'
      ImageIndex = 466
      OnExecute = actLayoutJanelaExecute
    end
  end
  object SavePictureDialog1: TSavePictureDialog
    Filter = 'Metafiles (*.wmf)|*.wmf'
    Left = 152
    Top = 120
  end
  object PopupMenuMapa: TPopupMenu
    Left = 181
    Top = 123
    object MapaMundi1: TMenuItem
      Action = actMapaMundi
    end
    object LatitudeseLongitudes1: TMenuItem
      Action = actLatitudeLongitude
    end
    object Brasil1: TMenuItem
      Action = actBrasil
    end
    object Carregarlinhas2Dx1y1x2y21: TMenuItem
      Action = actCarregarLinhas
    end
  end
  object OpenDialog1: TOpenDialog
    Left = 208
    Top = 120
  end
  object SaveDialog1: TSaveDialog
    Left = 237
    Top = 115
  end
end
