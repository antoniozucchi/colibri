object FrmAplatAnexos: TFrmAplatAnexos
  Left = 0
  Top = 0
  Caption = 'APLAT - Anexar PT'
  ClientHeight = 586
  ClientWidth = 1125
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
    Width = 1125
    Height = 25
    Align = alTop
    Caption = 'Fluxo Operacional do APLAT'
    Color = clGradientActiveCaption
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -13
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentBackground = False
    ParentFont = False
    TabOrder = 0
  end
  object PanelConteudo: TPanel
    Left = 0
    Top = 25
    Width = 1125
    Height = 561
    Align = alClient
    TabOrder = 1
    object MemoOrientacoes: TMemo
      Left = 1
      Top = 377
      Width = 1123
      Height = 113
      Align = alTop
      Lines.Strings = (
        'Fluxo sugerido para o lote de anexos no APLAT:'
        ''
        '1. Clique em "Abrir APLAT" e fa'#231'a o login manual.'
        '2. Informe a pasta que cont'#233'm os PDFs no formato PT.ANO.pdf.'
        '3. Clique em "Carregar PDFs" para montar a fila no ListBox.'
        '4. Revise os arquivos encontrados.'
        '5. Clique em "Anexar no APLAT" para executar o lote.'
        ''
        'O log operacional desta execu'#231#227'o tamb'#233'm aparece neste quadro.')
      ReadOnly = True
      TabOrder = 0
    end
    object GridPanel1: TGridPanel
      Left = 1
      Top = 1
      Width = 1123
      Height = 192
      Align = alTop
      Alignment = taLeftJustify
      ColumnCollection = <
        item
          Value = 100.000000000000000000
        end>
      ControlCollection = <
        item
          Column = 0
          Control = LabelUrl
          Row = 0
        end
        item
          Column = 0
          Control = edtUrlAplat
          Row = 1
        end
        item
          Column = 0
          Control = LabelEdge
          Row = 2
        end
        item
          Column = 0
          Control = edtEdgePath
          Row = 3
        end
        item
          Column = 0
          Control = LabelNumeroPT
          Row = 4
        end
        item
          Column = 0
          Control = edtNumeroPT
          Row = 5
        end
        item
          Column = 0
          Control = LabelPdf
          Row = 6
        end
        item
          Column = 0
          Control = edtCaminhoPdf
          Row = 7
        end>
      RowCollection = <
        item
          SizeStyle = ssAbsolute
          Value = 19.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 25.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 19.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 25.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 19.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 25.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 19.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 25.000000000000000000
        end>
      TabOrder = 1
      object LabelUrl: TLabel
        Left = 1
        Top = 1
        Width = 68
        Height = 19
        Align = alLeft
        Caption = 'URL do APLAT'
        ExplicitHeight = 13
      end
      object edtUrlAplat: TEdit
        Left = 1
        Top = 20
        Width = 800
        Height = 25
        Align = alLeft
        TabOrder = 0
        ExplicitHeight = 21
      end
      object LabelEdge: TLabel
        Left = 1
        Top = 45
        Width = 137
        Height = 19
        Align = alLeft
        Caption = 'Execut'#225'vel do Edge opcional'
        ExplicitHeight = 13
      end
      object edtEdgePath: TEdit
        Left = 1
        Top = 64
        Width = 800
        Height = 25
        Align = alLeft
        TabOrder = 1
        ExplicitHeight = 21
      end
      object LabelNumeroPT: TLabel
        Left = 1
        Top = 89
        Width = 67
        Height = 19
        Align = alLeft
        Caption = 'N'#250'mero da PT'
        ExplicitHeight = 13
      end
      object edtNumeroPT: TEdit
        Left = 1
        Top = 108
        Width = 800
        Height = 25
        Align = alLeft
        TabOrder = 2
        TextHint = 'Ex.: 002221/2024'
        ExplicitHeight = 21
      end
      object LabelPdf: TLabel
        Left = 1
        Top = 133
        Width = 74
        Height = 19
        Align = alLeft
        Caption = 'Pasta dos PDFs'
        ExplicitHeight = 13
      end
      object edtCaminhoPdf: TEdit
        Left = 1
        Top = 152
        Width = 800
        Height = 25
        Align = alLeft
        TabOrder = 3
        ExplicitHeight = 21
      end
    end
    object GridPanel2: TGridPanel
      Left = 1
      Top = 490
      Width = 1123
      Height = 41
      Align = alTop
      ColumnCollection = <
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end
        item
          SizeStyle = ssAbsolute
          Value = 180.000000000000000000
        end>
      ControlCollection = <
        item
          Column = 0
          Control = btnAbrirAplat
          Row = 0
        end
        item
          Column = 1
          Control = btnCarregarPT_PDF
          Row = 0
        end
        item
          Column = 2
          Control = btnAbrirPastaPdf
          Row = 0
        end
        item
          Column = 3
          Control = btnSalvarConfig
          Row = 0
        end
        item
          Column = 4
          Control = btnPrepararFluxo
          Row = 0
        end
        item
          Column = 5
          Control = btnAnexarPT_APLAT
          Row = 0
        end>
      RowCollection = <
        item
          Value = 100.000000000000000000
        end>
      TabOrder = 2
      ExplicitLeft = 0
      ExplicitTop = 506
      DesignSize = (
        1123
        41)
      object btnAbrirAplat: TButton
        Left = 16
        Top = 8
        Width = 150
        Height = 25
        Anchors = []
        Caption = 'Abrir APLAT'
        TabOrder = 0
        OnClick = btnAbrirAplatClick
      end
      object btnCarregarPT_PDF: TButton
        Left = 199
        Top = 8
        Width = 144
        Height = 25
        Anchors = []
        Caption = 'Carregar PDFs'
        TabOrder = 1
        OnClick = btnCarregarPT_PDFClick
      end
      object btnAbrirPastaPdf: TButton
        Left = 376
        Top = 8
        Width = 150
        Height = 25
        Anchors = []
        Caption = 'Abrir Pasta dos PDFs'
        TabOrder = 2
        OnClick = btnAbrirPastaPdfClick
        ExplicitLeft = 349
        ExplicitTop = 6
      end
      object btnSalvarConfig: TButton
        Left = 556
        Top = 8
        Width = 150
        Height = 25
        Anchors = []
        Caption = 'Salvar Configura'#231#245'es'
        TabOrder = 3
        OnClick = btnSalvarConfigClick
      end
      object btnPrepararFluxo: TButton
        Left = 739
        Top = 8
        Width = 144
        Height = 25
        Anchors = []
        Caption = 'Mostrar Passo a Passo'
        TabOrder = 4
        OnClick = btnPrepararFluxoClick
      end
      object btnAnexarPT_APLAT: TButton
        Left = 935
        Top = 8
        Width = 112
        Height = 25
        Anchors = []
        Caption = 'Anexar no APLAT'
        TabOrder = 5
        OnClick = btnAnexarPT_APLATClick
      end
    end
    object ListBoxPT_PDF: TListBox
      Left = 1
      Top = 193
      Width = 1123
      Height = 184
      Align = alTop
      ItemHeight = 13
      TabOrder = 3
    end
  end
  object OpenDialogPdf: TOpenDialog
    DefaultExt = 'pdf'
    Filter = 'Arquivos PDF|*.pdf'
    Left = 680
    Top = 128
  end
  object OpenDialogEdge: TOpenDialog
    DefaultExt = 'exe'
    Filter = 'Execut'#225'veis|*.exe'
    Left = 728
    Top = 128
  end
end
