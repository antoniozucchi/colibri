object FrmLogin: TFrmLogin
  Left = 0
  Top = 0
  BorderIcons = []
  BorderStyle = bsSizeToolWin
  Caption = 'Login'
  ClientHeight = 118
  ClientWidth = 437
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsStayOnTop
  KeyPreview = True
  Position = poScreenCenter
  Visible = True
  OnCreate = FormCreate
  OnShow = FormShow
  TextHeight = 13
  object GridPanel1: TGridPanel
    Left = 0
    Top = 0
    Width = 437
    Height = 65
    Align = alTop
    ColumnCollection = <
      item
        Value = 24.242424242424240000
      end
      item
        Value = 75.757575757575760000
      end>
    ControlCollection = <
      item
        Column = 0
        Control = Label1
        Row = 0
      end
      item
        Column = 1
        Control = edtUsuario
        Row = 0
      end
      item
        Column = 0
        Control = Label2
        Row = 1
      end
      item
        Column = 1
        Control = edtSenha
        Row = 1
      end>
    RowCollection = <
      item
        Value = 50.000000000000000000
      end
      item
        Value = 50.000000000000000000
      end>
    TabOrder = 0
    ExplicitWidth = 435
    DesignSize = (
      437
      65)
    object Label1: TLabel
      Left = 11
      Top = 10
      Width = 84
      Height = 13
      Anchors = []
      Caption = 'Nome de usu'#225'rio:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ExplicitLeft = 132
    end
    object edtUsuario: TEdit
      Left = 106
      Top = 1
      Width = 330
      Height = 32
      Align = alClient
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      ExplicitHeight = 21
    end
    object Label2: TLabel
      Left = 36
      Top = 42
      Width = 34
      Height = 13
      Anchors = []
      Caption = 'Senha:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      ExplicitLeft = 132
    end
    object edtSenha: TEdit
      Left = 106
      Top = 33
      Width = 330
      Height = 31
      Align = alClient
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      PasswordChar = '*'
      TabOrder = 1
      ExplicitWidth = 328
      ExplicitHeight = 21
    end
  end
  object GridPanel2: TGridPanel
    Left = 0
    Top = 65
    Width = 437
    Height = 53
    Align = alClient
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
        Control = BitBtnLogin
        Row = 0
      end
      item
        Column = 1
        Control = BitBtnCancel
        Row = 0
      end>
    RowCollection = <
      item
        Value = 100.000000000000000000
      end>
    TabOrder = 1
    ExplicitWidth = 435
    ExplicitHeight = 45
    DesignSize = (
      437
      53)
    object BitBtnLogin: TBitBtn
      Left = 60
      Top = 14
      Width = 100
      Height = 25
      Anchors = []
      Caption = 'Efetuar login'
      Kind = bkOK
      NumGlyphs = 2
      TabOrder = 0
      OnClick = actLoginExecute
    end
    object BitBtnCancel: TBitBtn
      Left = 277
      Top = 14
      Width = 100
      Height = 25
      Anchors = []
      Caption = 'Cancelar'
      Kind = bkCancel
      NumGlyphs = 2
      TabOrder = 1
      OnClick = BitBtnCancelClick
    end
  end
  object ActionManager1: TActionManager
    Top = 56
    StyleName = 'Platform Default'
    object actLogin: TAction
      Category = 'Login'
      Caption = 'actLogin'
      OnExecute = actLoginExecute
    end
  end
end
