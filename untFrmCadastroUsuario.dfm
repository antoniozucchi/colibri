object FrmCadastroUsuario: TFrmCadastroUsuario
  Left = 0
  Top = 0
  Caption = 'Cadastro de usu'#225'rios'
  ClientHeight = 416
  ClientWidth = 832
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
  object Panel2: TPanel
    Left = 0
    Top = 0
    Width = 832
    Height = 25
    Align = alTop
    Caption = 'Cadastro de usu'#225'rios'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 830
  end
  object ToolBar1: TToolBar
    Left = 0
    Top = 25
    Width = 832
    Height = 29
    Caption = 'ToolBar1'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 1
    ExplicitWidth = 830
    object DBNavigatorComissao: TDBNavigator
      Left = 0
      Top = 0
      Width = 138
      Height = 22
      DataSource = FrmDataModule.DataSourceCadastroUsuario
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
    object btnLayout: TToolButton
      Left = 161
      Top = 0
      Hint = 'Editar layout de colunas'
      Caption = 'btnLayout'
      ImageIndex = 196
      ParentShowHint = False
      ShowHint = True
    end
    object btnExcel: TToolButton
      Left = 184
      Top = 0
      Hint = 'Exportar dados para o Excel'
      Caption = 'Exportar'
      ImageIndex = 54
      ParentShowHint = False
      ShowHint = True
    end
  end
  object StatusBarRegrasRecolhimento: TStatusBar
    Left = 0
    Top = 397
    Width = 832
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
    ExplicitTop = 389
    ExplicitWidth = 830
  end
  object DBGridUsuarios: TFilterDBGrid
    Left = 0
    Top = 54
    Width = 832
    Height = 343
    Align = alClient
    DataSource = FrmDataModule.DataSourceCadastroUsuario
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnDrawColumnCell = DBGridUsuariosDrawColumnCell
    OnKeyPress = DBGridUsuariosKeyPress
    ClearFilterButton = btnClearFiltro
    SearchAction = actProcurar
    LayoutGrid = ColunasLayout
    EnableZebra = False
    LayoutButton = btnLayout
    ExcelButton = btnExcel
    Columns = <
      item
        Expanded = False
        FieldName = 'Chave'
        Title.Alignment = taCenter
        Width = 75
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Nome'
        Title.Alignment = taCenter
        Width = 254
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Perfil'
        PickList.Strings = (
          'Administrador'
          'Consulta'
          'Hotelaria'
          'Programa'#231#227'o'
          'Supervisor'
          'Requisi'#231#227'o de Transporte')
        Title.Alignment = taCenter
        Width = 156
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'AvaliadoPor'
        Title.Alignment = taCenter
        Title.Caption = 'Avaliado Por'
        Width = 85
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DataAtualizacao'
        Title.Alignment = taCenter
        Title.Caption = 'Data Atualiza'#231#227'o'
        Width = 104
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'Computador'
        Title.Alignment = taCenter
        Width = 81
        Visible = True
      end>
  end
  object ColunasLayout: TStringGrid
    Left = 49
    Top = 209
    Width = 113
    Height = 94
    ColCount = 7
    DefaultRowHeight = 21
    RowCount = 6
    TabOrder = 4
    Visible = False
    RowHeights = (
      21
      21
      21
      21
      21
      21)
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 360
    Top = 192
    StyleName = 'Platform Default'
    object actProcurar: TAction
      Caption = 'Procurar'
      Hint = 'Buscar registros no banco de dados'
      ImageIndex = 27
      OnExecute = actProcurarExecute
    end
  end
end
