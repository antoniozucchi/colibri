object FrmLogin: TFrmLogin
  Left = 0
  Top = 0
  AutoSize = True
  BorderIcons = []
  Caption = 'Login'
  ClientHeight = 105
  ClientWidth = 353
  Color = clBtnFace
  Font.Charset = ANSI_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  FormStyle = fsMDIChild
  KeyPreview = True
  OldCreateOrder = False
  Position = poScreenCenter
  Visible = True
  OnClose = FormClose
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 353
    Height = 105
    Color = clWhite
    ParentBackground = False
    TabOrder = 0
    object Label1: TLabel
      Left = 54
      Top = 11
      Width = 84
      Height = 13
      Caption = 'Nome de usu'#225'rio:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object Label2: TLabel
      Left = 54
      Top = 41
      Width = 34
      Height = 13
      Caption = 'Senha:'
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
    end
    object BitBtnLogin: TBitBtn
      Left = 146
      Top = 73
      Width = 95
      Height = 25
      Caption = 'Efetuar login'
      Kind = bkOK
      NumGlyphs = 2
      TabOrder = 2
      OnClick = actLoginExecute
    end
    object BitBtnCancel: TBitBtn
      Left = 248
      Top = 73
      Width = 95
      Height = 25
      Caption = 'Cancelar'
      Kind = bkCancel
      NumGlyphs = 2
      TabOrder = 3
      OnClick = BitBtnCancelClick
    end
    object edtUsuario: TEdit
      Left = 146
      Top = 8
      Width = 197
      Height = 21
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
    end
    object edtSenha: TEdit
      Left = 146
      Top = 38
      Width = 197
      Height = 21
      Font.Charset = ANSI_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      PasswordChar = '*'
      TabOrder = 1
    end
  end
  object ActionManager1: TActionManager
    Left = 112
    Top = 8
    StyleName = 'Platform Default'
    object actLogin: TAction
      Category = 'Login'
      Caption = 'actLogin'
      OnExecute = actLoginExecute
    end
  end
end
