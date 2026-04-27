object FrmConexaoLOCAL: TFrmConexaoLOCAL
  Left = 0
  Top = 0
  BorderStyle = bsSizeToolWin
  Caption = 'Conex'#227'o LOCAL'
  ClientHeight = 281
  ClientWidth = 697
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  OnCreate = FormCreate
  TextHeight = 15
  object PanelTitulo: TPanel
    Left = 0
    Top = 0
    Width = 697
    Height = 26
    Align = alTop
    Color = clGreen
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
    ExplicitWidth = 695
  end
  object GridPanel1: TGridPanel
    Left = 0
    Top = 61
    Width = 697
    Height = 142
    Align = alTop
    ColumnCollection = <
      item
        Value = 54.935370152761460000
      end
      item
        Value = 21.562867215041130000
      end
      item
        Value = 23.501762632197420000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = Label1
        Row = 0
      end
      item
        Column = 0
        Control = Label4
        Row = 1
      end
      item
        Column = 1
        Control = chkUPColibri
        Row = 1
      end
      item
        Column = 0
        Control = Label5
        Row = 2
      end
      item
        Column = 0
        Control = Label7
        Row = 3
      end
      item
        Column = 2
        Control = chkDownColibri
        Row = 1
      end
      item
        Column = 1
        Control = chkUPConsulta
        Row = 2
      end
      item
        Column = 2
        Control = chkDownConsulta
        Row = 2
      end
      item
        Column = 1
        Control = chkUPRT
        Row = 3
      end
      item
        Column = 2
        Control = chkDownRT
        Row = 3
      end
      item
        Column = 1
        Control = BitBtn1
        Row = 0
      end
      item
        Column = 2
        Control = BitBtn2
        Row = 0
      end>
    RowCollection = <
      item
        Value = 20.163983955494950000
      end
      item
        Value = 20.299685510111630000
      end
      item
        Value = 19.840463261971320000
      end
      item
        Value = 20.337495706050150000
      end
      item
        SizeStyle = ssAuto
      end>
    TabOrder = 1
    ExplicitWidth = 695
    DesignSize = (
      697
      142)
    object Label1: TLabel
      Left = 150
      Top = 7
      Width = 84
      Height = 15
      Anchors = []
      Caption = 'Banco de dados'
      ExplicitLeft = 152
    end
    object Label4: TLabel
      Left = 61
      Top = 36
      Width = 262
      Height = 15
      Anchors = []
      Caption = 'Programa'#231#227'o Di'#225'ria de Servi'#231'os (bdColibri.colibri)'
      ExplicitLeft = 165
    end
    object chkUPColibri: TCheckBox
      Left = 447
      Top = 32
      Width = 21
      Height = 22
      Anchors = []
      TabOrder = 0
    end
    object Label5: TLabel
      Left = 60
      Top = 64
      Width = 264
      Height = 15
      Anchors = []
      Caption = 'Tabelas de Cadastro e Consulta (dbConsulta.mdb)'
      ExplicitLeft = 165
    end
    object Label7: TLabel
      Left = 87
      Top = 91
      Width = 209
      Height = 15
      Anchors = []
      Caption = 'Tabelas de Cria'#231#227'o das RT'#39's (dbRT.mdb)'
      ExplicitLeft = 165
    end
    object chkDownColibri: TCheckBox
      Left = 604
      Top = 35
      Width = 20
      Height = 17
      Anchors = []
      TabOrder = 1
    end
    object chkUPConsulta: TCheckBox
      Left = 448
      Top = 63
      Width = 20
      Height = 17
      Anchors = []
      TabOrder = 2
    end
    object chkDownConsulta: TCheckBox
      Left = 604
      Top = 63
      Width = 20
      Height = 17
      Anchors = []
      TabOrder = 3
    end
    object chkUPRT: TCheckBox
      Left = 448
      Top = 90
      Width = 20
      Height = 17
      Anchors = []
      TabOrder = 4
    end
    object chkDownRT: TCheckBox
      Left = 604
      Top = 90
      Width = 20
      Height = 17
      Anchors = []
      TabOrder = 5
    end
    object BitBtn1: TBitBtn
      Left = 419
      Top = 2
      Width = 78
      Height = 25
      Action = actSelUpload
      Anchors = []
      Caption = 'Upload'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 8
    end
    object BitBtn2: TBitBtn
      Left = 570
      Top = 2
      Width = 88
      Height = 25
      Action = actSelDownload
      Anchors = []
      Caption = 'Download'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 9
    end
  end
  object ToolBar7: TToolBar
    Left = 0
    Top = 26
    Width = 697
    Height = 35
    ButtonHeight = 30
    Caption = 'ToolBar7'
    DrawingStyle = dsGradient
    EdgeBorders = [ebLeft, ebTop, ebRight, ebBottom]
    Images = FrmPrincipal.ImageList1
    TabOrder = 2
    ExplicitWidth = 695
    object BitBtn8: TBitBtn
      Left = 0
      Top = 0
      Width = 80
      Height = 30
      Action = actLOCAL
      Caption = 'LOCAL'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 1
    end
    object BitBtn10: TBitBtn
      Left = 80
      Top = 0
      Width = 73
      Height = 30
      Action = actREDE
      Caption = 'REDE'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 3
    end
    object BitBtn9: TBitBtn
      Left = 153
      Top = 0
      Width = 80
      Height = 30
      Action = actUpload
      Caption = 'Upload'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 2
    end
    object BitBtn7: TBitBtn
      Left = 233
      Top = 0
      Width = 90
      Height = 30
      Action = actDownload
      Caption = 'Download'
      ParentShowHint = False
      ShowHint = True
      TabOrder = 0
    end
  end
  object GridPanel2: TGridPanel
    Left = 0
    Top = 203
    Width = 697
    Height = 78
    Align = alTop
    ColumnCollection = <
      item
        Value = 22.178988326848250000
      end
      item
        Value = 77.821011673151760000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = Label8
        Row = 0
      end
      item
        Column = 1
        Control = edtRede
        Row = 0
      end
      item
        Column = 0
        Control = Label9
        Row = 1
      end
      item
        Column = 1
        Control = edtLocal
        Row = 1
      end>
    RowCollection = <
      item
        Value = 50.204081632653060000
      end
      item
        Value = 49.795918367346940000
      end>
    TabOrder = 3
    ExplicitWidth = 695
    DesignSize = (
      697
      78)
    object Label8: TLabel
      Left = 37
      Top = 12
      Width = 82
      Height = 15
      Anchors = []
      Caption = 'Endere'#231'o REDE:'
      ExplicitLeft = 40
    end
    object edtRede: TEdit
      Left = 155
      Top = 1
      Width = 541
      Height = 38
      Align = alClient
      Color = clInfoBk
      ReadOnly = True
      TabOrder = 0
      ExplicitHeight = 23
    end
    object Label9: TLabel
      Left = 32
      Top = 50
      Width = 92
      Height = 15
      Anchors = []
      Caption = 'Endere'#231'o LOCAL:'
      ExplicitLeft = 35
    end
    object edtLocal: TEdit
      Left = 155
      Top = 39
      Width = 541
      Height = 38
      Align = alClient
      Color = clInfoBk
      ReadOnly = True
      TabOrder = 1
      ExplicitWidth = 539
      ExplicitHeight = 23
    end
  end
  object ActionManager1: TActionManager
    Images = FrmPrincipal.ImageList1
    Left = 16
    Top = 120
    StyleName = 'Platform Default'
    object actLOCAL: TAction
      Category = 'Rede Local'
      Caption = 'LOCAL'
      Hint = 'Trabahar com banco de dados LOCAL'
      ImageIndex = 199
      OnExecute = actLOCALExecute
    end
    object actREDE: TAction
      Category = 'Rede Local'
      Caption = 'REDE'
      Enabled = False
      ImageIndex = 248
      OnExecute = actREDEExecute
    end
    object actUpload: TAction
      Category = 'Rede Local'
      Caption = 'Upload'
      Enabled = False
      Hint = 'Gravar dados LOCAL na REDE'
      ImageIndex = 482
      OnExecute = actUploadExecute
    end
    object actDownload: TAction
      Category = 'Rede Local'
      Caption = 'Download'
      Enabled = False
      Hint = 'Baixar dados da REDE para banco LOCAL'
      ImageIndex = 483
      OnExecute = actDownloadExecute
    end
    object actSelUpload: TAction
      Category = 'Rede Local'
      Caption = 'Upload'
      Hint = 'Limpar sele'#231#227'o'
      ImageIndex = 231
      OnExecute = actSelUploadExecute
    end
    object actSelDownload: TAction
      Category = 'Rede Local'
      Caption = 'Download'
      Hint = 'Selecionar todos'
      ImageIndex = 231
      OnExecute = actSelDownloadExecute
    end
  end
end
