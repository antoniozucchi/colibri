unit untConsultaProgramacao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ToolWin, Vcl.ExtCtrls, Vcl.StdCtrls, System.ImageList,
  Vcl.ImgList, System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls,
  Vcl.ActnMan, Data.Win.ADODB, Registry, ComOBJ, Vcl.Mask, Vcl.DBCtrls,DateUtils,
  Vcl.Menus, Vcl.Buttons,UITYPES,Math, Vcl.CheckLst,
  Vcl.FileCtrl, ComCtrls, Vcl.ExtDlgs, SDL_NumLab, untDBGridFilter, uZucchi;

type
  TFrmConsultaProgramacao = class(TForm)
    ActionManager1: TActionManager;
    actProcurarProgramacao: TAction;
    ProgressBar1: TProgressBar;
    PanelTitulo: TPanel;
    actExcluirTudo: TAction;
    actExcluirSelecionado: TAction;
    PopupMenuExcluir: TPopupMenu;
    ExcluirSelecionado1: TMenuItem;
    ExcluirTodos1: TMenuItem;
    actProcurarImportar: TAction;
    actAbrirImportar: TAction;
    actImportar: TAction;
    OpenDialog1: TOpenDialog;
    actImportarFiltrado: TAction;
    actExcluirFiltrados: TAction;
    ExcluirtodososregistrosdeProgramaoFILTRADOS1: TMenuItem;
    actGanttImprimir: TAction;
    actGanttAbrir: TAction;
    actGanttSalvar: TAction;
    SaveDialog1: TSaveDialog;
    actGanttDesfazer: TAction;
    actGanttRefazer: TAction;
    actGantt: TAction;
    actServicoInserir: TAction;
    actServicoExcluir: TAction;
    actExecutanteInserir: TAction;
    actExecutanteExcluir: TAction;
    actCopiarProgramacaoSELECAO: TAction;
    actCopiarProgramacaoTODAS: TAction;
    PopupMenuCopiar: TPopupMenu;
    Importarselecionado1: TMenuItem;
    Importarsemverificao1: TMenuItem;
    actInserirProgramacao: TAction;
    actExcluirRegistroProgramacao: TAction;
    actLimparExecutantes: TAction;
    actSelecionarServicos: TAction;
    actAjudaLimpar: TAction;
    actSelTodos: TAction;
    actLimparTodos: TAction;
    actInserirServicoSelecao: TAction;
    actGravarPIOP: TAction;
    actCopiar: TAction;
    PopupMenuListCopiar: TPopupMenu;
    Limpartudo1: TMenuItem;
    Excluirselecionado2: TMenuItem;
    actCopiarIncluirDate: TAction;
    actCopiarLimparDatas: TAction;
    actGRFSalvarImagem: TAction;
    actZoomMais: TAction;
    actZoomMenos: TAction;
    actGRFTitulo: TAction;
    actGRFImprimir: TAction;
    actGRFSalvarTXT: TAction;
    actEtiqueta: TAction;
    actGrafico: TAction;
    SavePictureDialog1: TSavePictureDialog;
    actGraficoTempo: TAction;
    actGraficoPlataforma: TAction;
    actGraficoTipoEtapaServico: TAction;
    actProgramacaoCancelar: TAction;
    actProgramacaoGravar: TAction;
    actExecutanteJanela: TAction;
    actVerificarAcessoExecutante: TAction;
    actExecutanteGravar: TAction;
    actCarregarDatas: TAction;
    PopupMenuGantt: TPopupMenu;
    Dias1: TMenuItem;
    Mes1: TMenuItem;
    Ano1: TMenuItem;
    Alterarlarguradacoluna1: TMenuItem;
    actProcurarServico: TAction;
    actLogAcao: TAction;
    PopupMenuTipoEtapaServico: TPopupMenu;
    IncluirTipoEtapadeServico1: TMenuItem;
    RemoverTipoEtapadeServico1: TMenuItem;
    actGravarProgramacao: TAction;
    actCancelarProgramacao: TAction;
    Splitter1: TSplitter;
    PanelProgramacao: TPanel;
    ToolBar1: TToolBar;
    BitBtn24: TBitBtn;
    BitBtn23: TBitBtn;
    BitBtn29: TBitBtn;
    BitBtn26: TBitBtn;
    BitBtn1: TBitBtn;
    btnExcluir: TToolButton;
    ToolButton1: TToolButton;
    DBEditProgramacao: TDBEdit;
    DBGridProgramacao: TFilterDBGrid;
    Panel13: TPanel;
    StatusBarProgramacao: TStatusBar;
    PanelFiltrosTodos: TPanel;
    PanelFiltros: TPanel;
    PanelDataFim: TPanel;
    Panel8: TPanel;
    Panel9: TPanel;
    dataFim: TDateTimePicker;
    PanelDataInicio: TPanel;
    Panel11: TPanel;
    Panel12: TPanel;
    dataInicio: TDateTimePicker;
    CheckListBoxStatus: TCheckListBox;
    PanelResultados: TPanel;
    Splitter2: TSplitter;
    PanelExecutante: TPanel;
    DBGridExecutantes: TFilterDBGrid;
    ToolBar2: TToolBar;
    BitBtn20: TBitBtn;
    BitBtn27: TBitBtn;
    DBEdit2: TDBEdit;
    Panel7: TPanel;
    StatusBarExecutantes: TStatusBar;
    PanelServico: TPanel;
    DBGridServicos: TFilterDBGrid;
    ToolBar3: TToolBar;
    DBNavigator1: TDBNavigator;
    BitBtn30: TBitBtn;
    BitBtn15: TBitBtn;
    BitBtn14: TBitBtn;
    Panel10: TPanel;
    StatusBarServicos: TStatusBar;
    PanelAjuda: TPanel;
    PanelTituloAjuda1: TPanel;
    SpeedButton4: TSpeedButton;
    PanelSelecaoServico: TPanel;
    ToolBar5: TToolBar;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    BitBtn35: TBitBtn;
    ComboBoxSelecaoServicoTipoEtapaServico: TComboBox;
    BitBtn28: TBitBtn;
    RLSelecionarServicos: TStringGrid;
    StatusBarSelecaoServico: TStatusBar;
    Panel18: TPanel;
    RadioGroup1: TRadioGroup;
    PanelImportar: TPanel;
    Panel1: TPanel;
    DBGridImportar: TFilterDBGrid;
    StatusBarImportar: TStatusBar;
    Panel3: TPanel;
    Panel4: TPanel;
    Panel5: TPanel;
    Panel6: TPanel;
    Panel14: TPanel;
    DateTimeImportarFIM: TDateTimePicker;
    Panel15: TPanel;
    Panel16: TPanel;
    Panel17: TPanel;
    DateTimeImportarINICIO: TDateTimePicker;
    CheckListBoxImportar: TCheckListBox;
    Panel2: TPanel;
    BitBtn2: TBitBtn;
    MemoImportar: TMemo;
    ToolBar4: TToolBar;
    BitBtn11: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn8: TBitBtn;
    PanelCopiar: TPanel;
    MonthCalendar1: TMonthCalendar;
    ToolBar7: TToolBar;
    BitBtn32: TBitBtn;
    BitBtn33: TBitBtn;
    BitBtn31: TBitBtn;
    ListBoxCopiar: TListBox;
    RadioGroupCopiar: TRadioGroup;
    PanelEditarExecutante: TPanel;
    ToolBar10: TToolBar;
    BitBtn38: TBitBtn;
    BitBtn39: TBitBtn;
    Panel40: TPanel;
    Panel41: TPanel;
    Panel42: TPanel;
    ComboBoxFuncao1: TComboBox;
    Panel46: TPanel;
    Panel47: TPanel;
    Panel48: TPanel;
    ComboBoxOrigem: TComboBox;
    PanelLogAcao: TPanel;
    DBMemoLogAcao: TDBMemo;
    ToolBar11: TToolBar;
    PanelInserirProgramacao: TPanel;
    ToolBar12: TToolBar;
    BitBtn40: TBitBtn;
    Panel71: TPanel;
    Panel72: TPanel;
    Panel73: TPanel;
    DateTimePickerProgramacao2: TDateTimePicker;
    Panel74: TPanel;
    Panel75: TPanel;
    Panel76: TPanel;
    ComboBoxJanelaDestino: TComboBox;
    Panel77: TPanel;
    Panel78: TPanel;
    Panel79: TPanel;
    ComboBoxJanelaTipoEtapaServico: TComboBox;
    btnClearFiltroProgramacao: TToolButton;
    btnExcelProgramacao: TToolButton;
    btnClearFiltroExecutante: TToolButton;
    btnExcelExecutante: TToolButton;
    btnClearFiltroServico: TToolButton;
    btnExcelServico: TToolButton;
    actProcurarExecutante: TAction;
    actProcurarServicos: TAction;
    procedure actProcurarProgramacaoExecute(Sender: TObject);
    procedure DBGridProgramacaoDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBEditProgramacaoChange(Sender: TObject);
    procedure DBGridExecutantesDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure actExcluirTudoExecute(Sender: TObject);
    procedure actExcluirSelecionadoExecute(Sender: TObject);
    procedure DBGridImportarDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actProcurarImportarExecute(Sender: TObject);
    procedure actAbrirImportarExecute(Sender: TObject);
    procedure actImportarExecute(Sender: TObject);
    procedure actImportarFiltradoExecute(Sender: TObject);
    procedure actGridASCImportarExecute(Sender: TObject);
    procedure actExcluirFiltradosExecute(Sender: TObject);
    procedure actServicoExcluirExecute(Sender: TObject);
    procedure actServicoInserirExecute(Sender: TObject);
    procedure actExecutanteExcluirExecute(Sender: TObject);
    procedure actExecutanteInserirExecute(Sender: TObject);
    procedure DBGridExecutantesKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridExecutantesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGridServicosKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure actCopiarProgramacaoSELECAOExecute(Sender: TObject);
    procedure actCopiarProgramacaoTODASExecute(Sender: TObject);
    procedure actExcluirRegistroProgramacaoExecute(Sender: TObject);
    procedure actInserirProgramacaoExecute(Sender: TObject);
    procedure DBGridServicosKeyPress(Sender: TObject; var Key: Char);
    procedure actLimparExecutantesExecute(Sender: TObject);
    procedure actSelecionarServicosExecute(Sender: TObject);
    procedure PanelTituloAjuda1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PanelTituloAjuda1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PanelTituloAjuda1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure actAjudaLimparExecute(Sender: TObject);
    procedure RLSelecionarServicosDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure RLSelecionarServicosFixedCellClick(Sender: TObject; ACol,
      ARow: Integer);
    procedure RLSelecionarServicosKeyPress(Sender: TObject; var Key: Char);
    procedure RLSelecionarServicosMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure actSelTodosExecute(Sender: TObject);
    procedure actLimparTodosExecute(Sender: TObject);
    procedure actInserirServicoSelecaoExecute(Sender: TObject);
    procedure MonthCalendar1DblClick(Sender: TObject);
    procedure ListBoxCopiarDblClick(Sender: TObject);
    procedure actCopiarExecute(Sender: TObject);
    procedure Excluirselecionado2Click(Sender: TObject);
    procedure actCopiarIncluirDateExecute(Sender: TObject);
    procedure actCopiarLimparDatasExecute(Sender: TObject);
    procedure actProgramacaoCancelarExecute(Sender: TObject);
    procedure actExecutanteJanelaExecute(Sender: TObject);
    procedure actVerificarAcessoExecutanteExecute(Sender: TObject);
    procedure actExecutanteGravarExecute(Sender: TObject);
    procedure ComboBoxTipoEtapaServicoKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBoxDestinoKeyPress(Sender: TObject; var Key: Char);
    procedure DBEdit2Change(Sender: TObject);
    procedure DBGridProgramacaoCellClick(Column: TColumn);
    procedure RadioGroup1Click(Sender: TObject);
    procedure actProcurarServicoExecute(Sender: TObject);
    procedure actLogAcaoExecute(Sender: TObject);

    procedure actGravarProgramacaoExecute(Sender: TObject);
    procedure actCancelarProgramacaoExecute(Sender: TObject);
    procedure actProcurarExecutanteExecute(Sender: TObject);
    procedure actProcurarServicosExecute(Sender: TObject);
    type
      TMemoriaExecutante = record
        idProgramacaoExecutante: Integer;
        Origem, NomeExecutante, Funcao, Empresa,
        CodigoSAP, Documento, RequisitantePT: String
    end;
    type
      THoras = record
        NumRegistros: Integer;
        SomaMinutos: Double;
    end;

  private
    { Private declarations }
    listProgramacao: TStringList;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    function carregaStatus(Grid: TFilterDBGrid): String;
    procedure selecionarServicos(TipoEtapaServico: String);
    procedure ExcluirFiltradosPorData(const DtIni, DtFim: TDateTime);
  public
    { Public declarations }
  end;

var
  FrmConsultaProgramacao: TFrmConsultaProgramacao;

implementation
  uses untPrincipal,untDataModule,untFrmPreview;

{$R *.dfm}

procedure TFrmConsultaProgramacao.actSelecionarServicosExecute(Sender: TObject);
begin
  selecionarServicos(FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('txtTipoEtapaServico').AsString);
end;

procedure TFrmConsultaProgramacao.actSelTodosExecute(Sender: TObject);
begin
  FrmPrincipal.selecaoGrid(RLSelecionarServicos,StatusBarSelecaoServico,' ');
end;

procedure TFrmConsultaProgramacao.actServicoExcluirExecute(Sender: TObject);
  var
    DataProgramacao: TDateTime;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;

    if DataProgramacao >= FrmPrincipal.carregaDataMinima(true) then
      FrmDataModule.ADOQueryProgramacaoServico_Consulta.Delete
    else
      MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
      'Colibri', MB_ICONERROR);
  except
    try
      FrmDataModule.ADOQueryProgramacaoServico_Consulta.Refresh;
    except
    end;
  end;
end;

procedure TFrmConsultaProgramacao.actServicoInserirExecute(Sender: TObject);
  var
    DataProgramacao: TDateTime;
    idProgramacaoDiaria: Integer;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;
    idProgramacaoDiaria:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.
    DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
    if idProgramacaoDiaria <> 0 then
    begin
      if DataProgramacao >= FrmPrincipal.carregaDataMinima(true) then
      begin
        FrmDataModule.ADOQueryProgramacaoServico_Consulta.Insert;
        FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
        FieldByName('CodigoProgramacaoDiaria').AsInteger:= idProgramacaoDiaria;
        FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('TextoBreveOP').AsString:=
        FrmPrincipal.incluirServicoPadrao(FrmDataModule.DataSourceProgramacaoDiaria_Consulta.
        DataSet.FieldByName('txtTipoEtapaServico').AsString);
        FrmDataModule.ADOQueryProgramacaoServico_Consulta.Post;
      end
      else
        MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
        'Colibri', MB_ICONERROR);
    end;
  except
  end;
end;

procedure TFrmConsultaProgramacao.actVerificarAcessoExecutanteExecute(
  Sender: TObject);
  var
    txtTipoEtapaServico,SQLBase,Origem,Funcao: String;
begin
  Origem:= FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.
  FieldByName('Origem').AsString;
  Funcao:= FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.
  FieldByName('Funcao').AsString;
  txtTipoEtapaServico:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.
  DataSet.FieldByName('txtTipoEtapaServico').AsString;
  SQLBase:= 'SELECT tblExecutante.* FROM tblExecutante '+
  'WHERE ((txtTipoEtapaServico LIKE '+QuotedStr(txtTipoEtapaServico)+')'+
  'AND(Ativo = True)) ORDER BY txtFuncao;';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBConsulta1.First;
  while not FrmDataModule.ADOQueryTemporarioDBConsulta1.Eof do
  begin
    ComboBoxFuncao1.Items.Add(FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtFuncao').AsString);
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
  FrmPrincipal.deleteRepetidosCombo(ComboBoxFuncao1);
  FrmPrincipal.selComboBox(ComboBoxOrigem,Origem);
  FrmPrincipal.selComboBox(ComboBoxFuncao1,Funcao);
end;

function TFrmConsultaProgramacao.carregaStatus(Grid: TFilterDBGrid): String;
  var
    SQLString: String;
begin
  SQLString:= BuildFilterSQL(Grid,false);
  //NumAprovados
  if (CheckListBoxStatus.Checked[0])and(SQLString <> '') then
    SQLString:=SQLString+'AND(NumAprovados > 0)'
  else if (CheckListBoxStatus.Checked[0])and(SQLString = '') then
    SQLString:=SQLString+'(NumAprovados > 0)';
  //NumCancelados
  if (CheckListBoxStatus.Checked[1])and(SQLString <> '') then
    SQLString:=SQLString+'AND(NumCancelados > 0)'
  else if (CheckListBoxStatus.Checked[1])and(SQLString = '') then
    SQLString:=SQLString+'(NumCancelados > 0)';
  //=========================================
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;

  Result:= SQLString;
end;

procedure TFrmConsultaProgramacao.ComboBoxDestinoKeyPress(Sender: TObject;
  var Key: Char);
begin
  key:= #0;
end;

procedure TFrmConsultaProgramacao.ComboBoxTipoEtapaServicoKeyPress(
  Sender: TObject; var Key: Char);
begin
  key:= #0;
end;

procedure TFrmConsultaProgramacao.actAbrirImportarExecute(Sender: TObject);
begin
  OpenDialog1.Filter:= 'Arquivo Colibri|*.colibri';
  OpenDialog1.DefaultExt:= '*.colibri';
  if OpenDialog1.Execute then
  begin
    MemoImportar.Text:= OpenDialog1.FileName;
    FrmPrincipal.registroEscrever('IMPORTAR',OpenDialog1.FileName);
    FrmPrincipal.conectarBDDireto(OpenDialog1.FileName,FrmDataModule.ADOConnectionImportar);
    actProcurarImportar.Execute;
  end;
end;

procedure TFrmConsultaProgramacao.actExcluirSelecionadoExecute(Sender: TObject);
  var
    buttonSelected: Integer;
begin
  buttonSelected :=
  MessageDlg('Deseja realmente excluir o registro selecionado?',
  mtConfirmation, [mbYes, mbNo], 0);
  if buttonSelected = mrYes then
    FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Delete;
end;

procedure TFrmConsultaProgramacao.actExcluirTudoExecute(Sender: TObject);
begin
  if Application.MessageBox(PChar(
  'Deseja realmente excluir todos os registros de Programa��o do banco de dados?'),
  '.::ATEN��O::.',36) = 6 then
  begin
    FrmPrincipal.deleteQueryRapido(
    FrmDataModule.ADOQueryProgramacaoDiaria_Consulta,'tblProgramacaoDiaria');
  end;
end;

procedure TFrmConsultaProgramacao.actExecutanteExcluirExecute(Sender: TObject);
  var
    DataProgramacao: TDateTime;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;

    if DataProgramacao >= FrmPrincipal.carregaDataMinima(false) then
    begin
      FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Delete;
      FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Edit;
      FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
      FieldByName('NumExecutantes').AsInteger:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
      FieldByName('NumExecutantes').AsInteger-1;
      FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
      FieldByName('NumAprovados').AsInteger:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
      FieldByName('NumAprovados').AsInteger-1;
      FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Post;
    end
    else
      MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
      'Colibri', MB_ICONERROR);
  except
    try
      FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Refresh;
    except
    end;
  end;
end;

procedure TFrmConsultaProgramacao.actExecutanteGravarExecute(Sender: TObject);
begin
  if (FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('DataProgramacao').AsDateTime >= FrmPrincipal.carregaDataMinima(false)) then
  begin
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Edit;
    FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Origem').AsString:=
    ComboBoxOrigem.Text;
    FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Funcao').AsString:=
    ComboBoxFuncao1.Text;
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Post;
  end
  else
    MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
    'Colibri', MB_ICONERROR);
end;

procedure TFrmConsultaProgramacao.actExecutanteInserirExecute(
  Sender: TObject);
  var
    DataProgramacao: TDateTime;
    idProgramacaoDiaria: Integer;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;
    idProgramacaoDiaria:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.
    DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
    if idProgramacaoDiaria <> 0 then
    begin
      if DataProgramacao >= FrmPrincipal.carregaDataMinima(false) then
      begin
        FrmPrincipal.ProgramarExecutantes('',idProgramacaoDiaria,
        FrmDataModule.ADOQueryProgramacaoExecutante_Consulta,
        FrmDataModule.ADOQueryProgramacaoDiaria_Consulta,
        FrmDataModule.DataSourceProgramacaoExecutante_Consulta,
        FrmDataModule.DataSourceProgramacaoDiaria_Consulta);
      end
      else
        MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
        'Colibri', MB_ICONERROR);
    end;
  except
  end;
end;

procedure TFrmConsultaProgramacao.actExecutanteJanelaExecute(Sender: TObject);
begin
  actAjudaLimpar.Execute;
  PanelTituloAjuda1.Caption:= 'Editar Executante';
  PanelEditarExecutante.Visible:= true;
  PanelEditarExecutante.Align:= alClient;
  PanelEditarExecutante.Visible:= true;
  PanelAjuda.Width:= 350;
  PanelAjuda.Height:= 110;
  PanelAjuda.Visible:= true;
  actVerificarAcessoExecutante.Execute;
end;

procedure TFrmConsultaProgramacao.actProcurarProgramacaoExecute(Sender: TObject);
  var
    DataProcuraIncio,DataProcuraFim,SQLString,SQLBase: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  SQLString:= carregaStatus(DBGridProgramacao);
  SQLBase:= 'SELECT tblProgramacaoDiaria.* FROM tblProgramacaoDiaria '+
  'WHERE (DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'
  +SQLString+' ORDER BY DataProgramacao;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryProgramacaoDiaria_Consulta,
  StatusBarProgramacao);
end;

procedure TFrmConsultaProgramacao.actProcurarExecutanteExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
    CodigoProgramacaoDiaria: Integer;
begin
  if FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.IsEmpty then
  begin
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Active := false;
    Exit;
  end;

  CodigoProgramacaoDiaria:= FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.
  FieldByName('idProgramacaoDiaria').AsInteger;
  SQLString:= BuildFilterSQL(DBGridExecutantes,false);
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:=
  'SELECT tblProgramacaoExecutante.* FROM tblProgramacaoExecutante '+
  'WHERE (CodigoProgramacaoDiaria = '+IntToStr(CodigoProgramacaoDiaria)+') '+
  SQLString+' ORDER BY NomeExecutante;';
  FrmPrincipal.ProcuraQuery(SQLBase,
  FrmDataModule.ADOQueryProgramacaoExecutante_Consulta,StatusBarExecutantes);
end;

procedure TFrmConsultaProgramacao.actProcurarImportarExecute(Sender: TObject);
  var
    DataProcuraIncio,DataProcuraFim,SQLString,SQLBase: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',DateTimeImportarINICIO.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',DateTimeImportarFIM.Date));
  SQLString:= carregaStatus(DBGridImportar);
  SQLBase:= 'SELECT tblProgramacaoDiaria.* FROM tblProgramacaoDiaria '+
  'WHERE (DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'
  +SQLString+' ORDER BY DataProgramacao;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryProgramacaoDiaria_Importar,
  StatusBarImportar);
end;

procedure TFrmConsultaProgramacao.actProcurarServicoExecute(Sender: TObject);
begin
  selecionarServicos(ComboBoxSelecaoServicoTipoEtapaServico.Text);
end;

procedure TFrmConsultaProgramacao.actProcurarServicosExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
    CodigoProgramacaoDiaria: Integer;
begin
  if FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.IsEmpty then
  begin
    FrmDataModule.ADOQueryProgramacaoServico_Consulta.Active := false;
    Exit;
  end;

  CodigoProgramacaoDiaria:= FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.
  FieldByName('idProgramacaoDiaria').AsInteger;
  SQLString:= BuildFilterSQL(DBGridServicos,false);
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:=
  'SELECT tblProgramacaoServico.* FROM tblProgramacaoServico '+
  'WHERE (CodigoProgramacaoDiaria = '+IntToStr(CodigoProgramacaoDiaria)+') '+
  SQLString+';';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryProgramacaoServico_Consulta,
  StatusBarServicos);
end;

procedure TFrmConsultaProgramacao.actGridASCImportarExecute(Sender: TObject);
begin
  FrmPrincipal.ClassificaDBGrid(DBGridImportar,FrmDataModule.ADOQueryProgramacaoDiaria_Importar,0);
end;

procedure TFrmConsultaProgramacao.actProgramacaoCancelarExecute(Sender: TObject);
begin
  actAjudaLimpar.Execute;
end;

procedure TFrmConsultaProgramacao.actImportarExecute(Sender: TObject);
  var
    Endereco: String;
begin
  Endereco:= FrmPrincipal.registroEndereco('Banco de dados');
  MemoImportar.Text:= Endereco;
  try
    FrmPrincipal.conectarBDDireto(Endereco,FrmDataModule.ADOConnectionImportar);
    actProcurarImportar.Execute;
  except
    MessageBox(0, 'O endere�o do banco de dados n�o pode ser aberto!',
    'Colibri', MB_ICONERROR);
  end;
  DateTimeImportarINICIO.Date:= IncDay(now,1);
  DateTimeImportarFIM.Date:= IncDay(now,1);
  actAjudaLimpar.Execute;
  PanelImportar.Visible:= true;
  PanelImportar.Align:= alClient;
  PanelAjuda.Visible:= true;
  PanelAjuda.Width:= 650;
  PanelAjuda.Height:= 450;
end;

procedure TFrmConsultaProgramacao.actImportarFiltradoExecute(Sender: TObject);
  var
    idProgramacaoDiariaImportar,idProgramacaoDiariaInserir: Integer;
    LogAcao,DataCriacao: String;
begin
  FrmDataModule.naoGravar:= true;
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.
  ADOQueryProgramacaoDiaria_Importar.RecordCount,'Importando prorama��o filtrada...');
  FrmDataModule.ADOQueryProgramacaoDiaria_Importar.First;
  FrmDataModule.ADOQueryInserirProgramacao.Active:= false;
  FrmDataModule.ADOQueryInserirProgramacao.Active:= true;
  FrmDataModule.ADOQueryInserirExecutante1.Active:= false;
  FrmDataModule.ADOQueryInserirExecutante1.Active:= true;
  FrmDataModule.ADOQueryInserirServico.Active:= false;
  FrmDataModule.ADOQueryInserirServico.Active:= true;
  FrmDataModule.DataSourceProgramacaoDiaria_Importar.Enabled:= false;
  DataCriacao:= DateToStr(now);
  while not FrmDataModule.ADOQueryProgramacaoDiaria_Importar.Eof do
  begin
    idProgramacaoDiariaImportar:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
    //===================================================
    FrmDataModule.ADOQueryInserirProgramacao.Insert;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('DataProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('DataProgramacao').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('txtTipoEtapaServico').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('txtTipoEtapaServico').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('txtDestino').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('txtDestino').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('NumExecutantes').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('NumExecutantes').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('NumAprovados').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('NumAprovados').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('NumCancelados').AsString:= FrmDataModule.DataSourceProgramacaoDiaria_Importar.
    DataSet.FieldByName('NumCancelados').AsString;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('CriadoPor').AsString:= FrmPrincipal.logChave;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('DataCriacao').AsString:= DataCriacao;
    FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('ComputadorCriacao').AsString:= FrmPrincipal.logMaquina;
    LogAcao:= FrmPrincipal.logChave+#9+DataCriacao+#9+FrmPrincipal.logMaquina+#9+'Cria��o';
    FrmPrincipal.MemoPrincipal.Text:= FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('LogAcao').AsString;
    FrmPrincipal.MemoPrincipal.Lines.Add(LogAcao);
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('LogAcao').AsString:= FrmPrincipal.MemoPrincipal.Text;
    FrmDataModule.ADOQueryInserirProgramacao.Post;
    //===================================================
    idProgramacaoDiariaInserir:= FrmDataModule.DataSourceInserirProgramacao.
    DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
    //===================================================
    FrmDataModule.ADOQueryProgramacaoExecutante_Importar.Active := false;
    FrmDataModule.ADOQueryProgramacaoExecutante_Importar.Parameters.Items[0].Value:=
    idProgramacaoDiariaImportar;
    FrmDataModule.ADOQueryProgramacaoExecutante_Importar.Active := true;
    FrmDataModule.ADOQueryProgramacaoExecutante_Importar.First;
    while not FrmDataModule.ADOQueryProgramacaoExecutante_Importar.Eof do
    begin
      FrmDataModule.ADOQueryInserirExecutante1.Insert;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger:= idProgramacaoDiariaInserir;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('RequisitantePT').AsBoolean:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('RequisitantePT').AsBoolean;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('Kit_PS').AsBoolean:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('Kit_PS').AsBoolean;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('CodigoSAP').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('CodigoSAP').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('Documento').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('Documento').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('NomeExecutante').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('NomeExecutante').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('Funcao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('Funcao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('Empresa').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('Empresa').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('Origem').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('Origem').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('TipoEmbarque').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('TipoEmbarque').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('RT').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('RT').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('MotivoProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('MotivoProgramacao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('PalavraChaveProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('PalavraChaveProgramacao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('AvaliadoPorProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('AvaliadoPorProgramacao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('ComputadorProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('ComputadorProgramacao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('StatusProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('StatusProgramacao').AsString;
      FrmDataModule.DataSourceInserirExecutante.
      DataSet.FieldByName('DataAvaliacaoProgramacao').AsString:= FrmDataModule.DataSourceProgramacaoExecutante_Importar.
      DataSet.FieldByName('DataAvaliacaoProgramacao').AsString;
      FrmDataModule.ADOQueryInserirExecutante1.Post;
      FrmDataModule.ADOQueryProgramacaoExecutante_Importar.Next;
    end;
    //===================================================
    FrmDataModule.ADOQueryProgramacaoServico_Importar.Active := false;
    FrmDataModule.ADOQueryProgramacaoServico_Importar.Parameters.Items[0].Value:=
    idProgramacaoDiariaImportar;
    FrmDataModule.ADOQueryProgramacaoServico_Importar.Active := true;
    FrmDataModule.ADOQueryProgramacaoServico_Importar.First;
    while not FrmDataModule.ADOQueryProgramacaoServico_Importar.Eof do
    begin
      FrmDataModule.ADOQueryInserirServico.Insert;
      FrmDataModule.DataSourceInserirServico.
      DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger:= idProgramacaoDiariaInserir;
      FrmDataModule.DataSourceInserirServico.
      DataSet.FieldByName('CampoSelecao').AsString:= FrmDataModule.DataSourceProgramacaoServico_Importar.
      DataSet.FieldByName('CampoSelecao').AsString;
      FrmDataModule.DataSourceInserirServico.
      DataSet.FieldByName('OrdemManutencao').AsString:= FrmDataModule.DataSourceProgramacaoServico_Importar.
      DataSet.FieldByName('OrdemManutencao').AsString;
      FrmDataModule.DataSourceInserirServico.
      DataSet.FieldByName('TextoBreveOP').AsString:= FrmDataModule.DataSourceProgramacaoServico_Importar.
      DataSet.FieldByName('TextoBreveOP').AsString;
      FrmDataModule.ADOQueryInserirServico.Post;
      FrmDataModule.ADOQueryProgramacaoServico_Importar.Next;
    end;
    //===================================================
    FrmPrincipal.ProgressBarIncremento(1);
    FrmDataModule.ADOQueryProgramacaoDiaria_Importar.Next;
  end;
  FrmDataModule.ADOQueryProgramacaoDiaria_Importar.First;
  FrmDataModule.DataSourceProgramacaoDiaria_Importar.Enabled:= true;
  FrmPrincipal.ProgressBarAtualizar;
  actProcurarProgramacao.Execute;
  FrmDataModule.naoGravar:= false;
end;

procedure TFrmConsultaProgramacao.actInserirProgramacaoExecute(Sender: TObject);
begin
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'Plataforma',
  'SELECT Plataforma FROM tblPlataforma ORDER BY Plataforma;',ComboBoxJanelaDestino);
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'TipoEtapaServico',
  'SELECT TipoEtapaServico FROM tblTipoEtapaServico ORDER BY TipoEtapaServico;',ComboBoxJanelaTipoEtapaServico);
  DateTimePickerProgramacao2.Date:= dataInicio.Date;
  DateTimePickerProgramacao2.Time:= 0;
  actAjudaLimpar.Execute;
  PanelInserirProgramacao.Visible:= true;
  PanelInserirProgramacao.Align:= alClient;
  PanelTituloAjuda1.Caption:= 'Inserir registro de programa��o di�ria de servi�o';
  PanelAjuda.Width:= 350;
  PanelAjuda.Height:= 140;
  PanelAjuda.Top:= 250;
  PanelAjuda.Left:= 200;
  PanelAjuda.Visible:= true;
end;

procedure TFrmConsultaProgramacao.actInserirServicoSelecaoExecute(
  Sender: TObject);
  var
    DataProgramacao: TDateTime;
    i,CodigoProgramacaoDiaria: Integer;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;
    CodigoProgramacaoDiaria:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('idProgramacaoDiaria').AsInteger;
    if CodigoProgramacaoDiaria <> 0 then
    begin
      if DataProgramacao >= FrmPrincipal.carregaDataMinima(false) then
      begin
        //Excluir todos os Servi�os existentes anteriormente
        FrmDataModule.ADOQueryProgramacaoServico_Cadastro.Active := false;
        FrmDataModule.ADOQueryProgramacaoServico_Cadastro.Parameters.Items[0].Value:= CodigoProgramacaoDiaria;
        FrmDataModule.ADOQueryProgramacaoServico_Cadastro.Active := true;
        FrmPrincipal.deleteQuery(FrmDataModule.ADOQueryProgramacaoServico_Cadastro,'Servi�os');
        for I := 1 to RLSelecionarServicos.RowCount-1 do
        begin
          if RLSelecionarServicos.Cells[0,i] = ' ' then
          begin
            FrmDataModule.ADOQueryProgramacaoServico_Consulta.Insert;
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger:=
            CodigoProgramacaoDiaria;
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('CampoSelecao').AsString:=
            RLSelecionarServicos.Cells[1,i];
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('TextoBreveOP').AsString:=
            RLSelecionarServicos.Cells[2,i];
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('TextoBreveOM').AsString:=
            RLSelecionarServicos.Cells[3,i];
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('OrdemManutencao').AsString:=
            RLSelecionarServicos.Cells[4,i];
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('Operacao').AsString:=
            RLSelecionarServicos.Cells[5,i];
            FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.FieldByName('CentroTrabalhoOP').AsString:=
            RLSelecionarServicos.Cells[6,i];
            FrmDataModule.ADOQueryProgramacaoServico_Consulta.Post;
          end;
        end;
      end
      else
        MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
        'Colibri', MB_ICONERROR);
    end;
  except
  end;
  actAjudaLimpar.Execute;
  FrmDataModule.ADOQueryProgramacaoServico_Consulta.Active:= false;
  FrmDataModule.ADOQueryProgramacaoServico_Consulta.Active:= true;
end;

procedure TFrmConsultaProgramacao.actGravarProgramacaoExecute(Sender: TObject);
  var
    Destino,TipoEtapaServico,LogAcao: String;
begin
  try
    Destino:= ComboBoxJanelaDestino.Text;
    TipoEtapaServico:= ComboBoxJanelaTipoEtapaServico.Text;
    if ((Destino <> '')AND(TipoEtapaServico<>'')) then
    begin
      FrmDataModule.ADOQueryInserirProgramacao.Active:= true;
      FrmDataModule.ADOQueryInserirProgramacao.Insert;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('DataProgramacao').AsDateTime:= DateTimePickerProgramacao2.DateTime;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('txtDestino').AsString:= Destino;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('txtTipoEtapaServico').AsString:= TipoEtapaServico;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('NumExecutantes').AsInteger:= 0;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('NumAprovados').AsInteger:= 0;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('NumCancelados').AsInteger:= 0;
      //=========================================
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('CriadoPor').AsString:= FrmPrincipal.logChave;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('DataCriacao').AsDateTime:= now;
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('ComputadorCriacao').AsString:= FrmPrincipal.logMaquina;
      //=========================================
      FrmPrincipal.MemoPrincipal.Text:= FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('LogAcao').AsString;
      LogAcao:= 'Programa��o: '+DateToStr(DateTimePickerProgramacao2.DateTime)+#9+Destino+#9+TipoEtapaServico+#9+
      '['+FrmPrincipal.logChave+'; '+DateToStr(now)+'; '+FrmPrincipal.logMaquina+'; Cria��o]';
      FrmPrincipal.MemoPrincipal.Lines.Add(LogAcao);
      FrmDataModule.DataSourceInserirProgramacao.DataSet.
      FieldByName('LogAcao').AsString:= FrmPrincipal.MemoPrincipal.Text;
      //==========================================
      FrmDataModule.ADOQueryInserirProgramacao.Post;
      FrmConsultaProgramacao.actProcurarProgramacao.Execute;
      PanelAjuda.Visible:= false;
    end
    else
    begin
      MessageBox(0, 'Necess�rio preencher antes os campos: "Destino" e "Tipo de Etapa de Servi�o"',
      'Colibri', MB_ICONERROR);
    end;
  except
  end;
end;

procedure TFrmConsultaProgramacao.actLimparTodosExecute(Sender: TObject);
begin
  FrmPrincipal.selecaoGrid(RLSelecionarServicos,StatusBarSelecaoServico,'  ');
end;

procedure TFrmConsultaProgramacao.actLogAcaoExecute(Sender: TObject);
begin
  actAjudaLimpar.Execute;
  PanelTituloAjuda1.Caption:= 'Log de A��o - Hist�rico';
  PanelLogAcao.Visible:= true;
  PanelLogAcao.Align:= alClient;
  PanelLogAcao.Visible:= true;
  PanelAjuda.Width:= 650;
  PanelAjuda.Height:= 300;
  PanelAjuda.Visible:= true;
end;

procedure TFrmConsultaProgramacao.actCopiarProgramacaoTODASExecute(Sender: TObject);
begin
  PanelTituloAjuda1.Caption:= 'Copiar todoas as programa��es "FILTRADAS"';
  actAjudaLimpar.Execute;
  PanelCopiar.Visible:= true;
  PanelCopiar.Align:= alClient;
  PanelAjuda.Visible:= true;
  PanelAjuda.Width:= 350;
  PanelAjuda.Height:= 200;
  RadioGroupCopiar.ItemIndex:= 1;
  MonthCalendar1.MultiSelect:= false;
  MonthCalendar1.Date:= today;
  MonthCalendar1.MultiSelect:= true;
end;

procedure TFrmConsultaProgramacao.actAjudaLimparExecute(Sender: TObject);
begin
  PanelSelecaoServico.Visible:= false;
  PanelImportar.Visible:= false;
  PanelCopiar.Visible:= false;
  PanelEditarExecutante.Visible:= false;
  PanelLogAcao.Visible:= false;
  PanelInserirProgramacao.Visible:= false;
  PanelAjuda.Visible:= false;
end;

procedure TFrmConsultaProgramacao.actCancelarProgramacaoExecute(
  Sender: TObject);
begin
  actAjudaLimpar.Execute;
end;

procedure TFrmConsultaProgramacao.actCopiarExecute(Sender: TObject);
  var
    DataProgramacao: String;
    i: Integer;
begin
  case RadioGroupCopiar.ItemIndex of
    0:
    begin
      if ListBoxCopiar.Count <> 0  then
      begin
        for I := 0 to ListBoxCopiar.Count-1 do
        begin
          DataProgramacao:= ListBoxCopiar.Items[i];
          if FrmPrincipal.isData(DataProgramacao) then
            FrmPrincipal.CopiarProgramacao(StrToDate(DataProgramacao),
            FrmDataModule.DataSourceProgramacaoDiaria_Consulta);
        end;
        dataInicio.DateTime:= StrToDate(ListBoxCopiar.Items[0]);
        dataFim.DateTime:= StrToDate(ListBoxCopiar.Items[ListBoxCopiar.Count-1]);
        actProcurarProgramacao.Execute;
      end
      else
        MessageBox(0,'Nenhuma data selecionda.','Colibri',MB_ICONEXCLAMATION)
    end;
    1:
    begin
      if ListBoxCopiar.Count <> 0  then
      begin
        for I := 0 to ListBoxCopiar.Count-1 do
        begin
          DataProgramacao:= ListBoxCopiar.Items[i];
          if FrmPrincipal.isData(DataProgramacao) then
          begin
            FrmDataModule.ADOQueryInserirProgramacao.Active:= false;
            FrmDataModule.ADOQueryInserirProgramacao.Active:= true;
            FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.First;
            FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.RecordCount,
            'Copiando...');
            while not FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Eof do
            begin
              FrmPrincipal.CopiarProgramacao(StrToDate(DataProgramacao),
              FrmDataModule.DataSourceProgramacaoDiaria_Consulta);
              FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Next;
              FrmPrincipal.ProgressBarIncremento(1);
            end;
          end;
        end;
        dataInicio.DateTime:= StrToDate(ListBoxCopiar.Items[0]);
        dataFim.DateTime:= StrToDate(ListBoxCopiar.Items[ListBoxCopiar.Count-1]);
        FrmPrincipal.ProgressBarAtualizar;
        actProcurarProgramacao.Execute;
      end
      else
        MessageBox(0,'Nenhuma data selecionda.','Colibri',MB_ICONEXCLAMATION)
    end;
  end;
end;

procedure TFrmConsultaProgramacao.actCopiarIncluirDateExecute(Sender: TObject);
  var
    i,TotalDias: Integer;
    data1,data2: TDateTime;
begin
  data1:= MonthCalendar1.Date;
  data2:= MonthCalendar1.EndDate;
  TotalDias:=  daysbetween(data2,data1);
  for I := 0 to TotalDias do
  begin
    ListBoxCopiar.Items.Add(FormatDateTime('dd/mm/yyyy',IncDay(data1,i)));
  end;
end;

procedure TFrmConsultaProgramacao.actCopiarProgramacaoSELECAOExecute(Sender: TObject);
begin
  PanelTituloAjuda1.Caption:= 'Copiar a programa��o "SELECIONADA"';
  actAjudaLimpar.Execute;
  PanelCopiar.Visible:= true;
  PanelCopiar.Align:= alClient;
  PanelAjuda.Visible:= true;
  PanelAjuda.Width:= 350;
  PanelAjuda.Height:= 200;
  RadioGroupCopiar.ItemIndex:= 0;
  MonthCalendar1.MultiSelect:= false;
  MonthCalendar1.Date:= today;
  MonthCalendar1.MultiSelect:= true;
end;

procedure TFrmConsultaProgramacao.ExcluirFiltradosPorData(const DtIni, DtFim: TDateTime);
var
  QDel: TADOQuery;
begin
  QDel := TADOQuery.Create(nil);
  try
    QDel.Connection := FrmDataModule.ADOConnectionColibri; // ajuste

    // transa��o ajuda no Access
    QDel.Connection.BeginTrans;
    try
      QDel.SQL.Text :=
        'DELETE FROM tblProgramacaoDiaria '+
        'WHERE DataProgramacao >= ? AND DataProgramacao < ?';

      QDel.Parameters[0].DataType := ftDateTime;
      QDel.Parameters[0].Value := Trunc(DtIni);

      QDel.Parameters[1].DataType := ftDateTime;
      QDel.Parameters[1].Value := Trunc(DtFim) + 1;

      QDel.ExecSQL;

      QDel.Connection.CommitTrans;
      Application.MessageBox(PChar('Exclus�o conclu�da com sucesso'),
      'Conclu�do', MB_OK + MB_ICONINFORMATION);
    except
      QDel.Connection.RollbackTrans;
      Application.MessageBox(PChar('Erro durante processo de exclus�o'),
      'Conclu�do', MB_OK + MB_ICONINFORMATION);
      raise;
    end;
  finally
    QDel.Free;
  end;
end;

procedure TFrmConsultaProgramacao.actExcluirFiltradosExecute(Sender: TObject);
{var
  DtIni, DtFimExclusivo: TDateTime;
  SQLString, SQLWhere, SQLDel: string;
  QDel: TADOQuery;
  Rows: Integer;}
begin
  if Application.MessageBox(PChar(
    'Deseja realmente excluir todos os registros de programa��o "DATAS FILTRADAS"?'),
    '.::ATEN��O::.', MB_YESNO + MB_ICONWARNING) <> IDYES then
    Exit;

  ExcluirFiltradosPorData(dataInicio.Date,dataFim.Date);


  {if Application.MessageBox(PChar(
    'Deseja realmente excluir todos os registros de programa��o "FILTRADOS"?'),
    '.::ATEN��O::.', MB_YESNO + MB_ICONWARNING) <> IDYES then
    Exit;

  DtIni := Trunc(dataInicio.Date);
  DtFimExclusivo := Trunc(dataFim.Date) + 1;

  SQLString := carregaStatus(DBGridProgramacao); // " AND ..."
  SQLWhere  := ' WHERE DataProgramacao >= ? AND DataProgramacao < ?' + SQLString;

  SQLDel := 'DELETE FROM tblProgramacaoDiaria' + SQLWhere;

  FrmDataModule.DataSourceProgramacaoDiaria_Consulta.Enabled := False;
  try
    QDel := TADOQuery.Create(nil);
    try
      QDel.Connection := FrmDataModule.ADOConnectionColibri; // ajuste se for outra conex�o
      QDel.SQL.Text := SQLDel;

      QDel.Parameters[0].DataType := ftDateTime;
      QDel.Parameters[0].Value := DtIni;

      QDel.Parameters[1].DataType := ftDateTime;
      QDel.Parameters[1].Value := DtFimExclusivo;

      QDel.Connection.BeginTrans;
      try
        QDel.ExecSQL;
        Rows := QDel.RowsAffected; // �s vezes pode vir -1 no Access, mas geralmente vem ok
        QDel.Connection.CommitTrans;
      except
        QDel.Connection.RollbackTrans;
        raise;
      end;

    finally
      QDel.Free;
    end;

    // Recarrega a grade
    FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Close;
    FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Open;

    Application.MessageBox(PChar(Format('Exclus�o conclu�da. Registros afetados: %d', [Rows])),
      'Conclu�do', MB_OK + MB_ICONINFORMATION);

  finally
    FrmDataModule.DataSourceProgramacaoDiaria_Consulta.Enabled := True;
  end; }
end;

procedure TFrmConsultaProgramacao.actExcluirRegistroProgramacaoExecute(
  Sender: TObject);
  var
    DataProgramacao: TDateTime;
begin
  try
    DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
    FieldByName('DataProgramacao').AsDateTime;

    if DataProgramacao >= FrmPrincipal.carregaDataMinima(false) then
    begin
      if Application.MessageBox(PChar('Deseja realmente excluir a programa��o selecionada?'),
      '.::ATEN��O::.',36) = 6 then
        FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Delete;
    end
    else
      MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
      'Colibri', MB_ICONERROR);
  except
    FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Refresh;
  end;
end;

procedure TFrmConsultaProgramacao.actCopiarLimparDatasExecute(Sender: TObject);
begin
  ListBoxCopiar.Items.Clear;
end;

procedure TFrmConsultaProgramacao.actLimparExecutantesExecute(Sender: TObject);
begin
  if Application.MessageBox(PChar('Deseja realmente limpar todas as programa��es?'),
  '.::ATEN��O::.',36) = 6 then
  begin
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.First;
    while not FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Eof do
    begin
      FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Edit;
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('NomeExecutante').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Empresa').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('CodigoSAP').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Documento').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('StatusProgramacao').AsString:= 'Aprovado';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('MotivoProgramacao').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean:= false;
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('PalavraChaveProgramacao').AsString:= '';
      FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('RT').AsString:= '';
      FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Post;
      FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Next;
    end;
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.First;
  end;
end;

procedure TFrmConsultaProgramacao.DBEdit2Change(Sender: TObject);
begin
  if (PanelEditarExecutante.Visible)AND(PanelAjuda.Visible) then
  begin
    actVerificarAcessoExecutante.Execute;
  end;
end;

procedure TFrmConsultaProgramacao.DBEditProgramacaoChange(Sender: TObject);
  var
    NumExecutantes: Integer;
begin
  if DBEditProgramacao.Text <> '' then
  begin
    try
      actProcurarExecutante.Execute;
      actProcurarServicos.Execute;

      NumExecutantes:= FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.RecordCount;
      StatusBarExecutantes.Panels[0].Text:= 'Nº Registros: '+IntToStr(NumExecutantes);
      AutoFitStatusBar(StatusBarExecutantes);
      StatusBarServicos.Panels[0].Text:= 'Nº Registros: '+
      IntToStr(FrmDataModule.ADOQueryProgramacaoServico_Consulta.RecordCount);
      AutoFitStatusBar(StatusBarServicos);
    except
      FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Cancel;
      FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Refresh;
    end;
  end
  else
  begin
    FrmDataModule.ADOQueryProgramacaoServico_Consulta.Active := false;
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Active := false;
  end;
end;

procedure TFrmConsultaProgramacao.DBGridExecutantesDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

  if (Column.Field.FieldName = 'StatusProgramacao')OR
  (Column.Field.FieldName = 'NomeExecutante') then
  begin
    if FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Aprovado' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clLime;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Cancelado' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clRed;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceProgramacaoExecutante_Consulta.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Mudan�a' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clYellow;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmConsultaProgramacao.DBGridExecutantesKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridExecutantes.DataSource.DataSet.RecNo=
  DBGridExecutantes.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmConsultaProgramacao.DBGridExecutantesKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key:= #0;
end;

procedure TFrmConsultaProgramacao.DBGridImportarDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin

  if (Column.Field.FieldName = 'NumCancelados') then
  begin
    if FrmDataModule.DataSourceProgramacaoDiaria_Importar.DataSet.
    FieldByName('NumCancelados').AsInteger > 0 then
    begin
      DBGridImportar.Canvas.Brush.Color:= clRed;
      DBGridImportar.Font.Color:= clBlack;
      DBGridImportar.Canvas.FillRect(Rect);
      DBGridImportar.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumAprovados') then
  begin
    if FrmDataModule.DataSourceProgramacaoDiaria_Importar.DataSet.
    FieldByName('NumAprovados').AsInteger > 0 then
    begin
      DBGridImportar.Canvas.Brush.Color:= clLime;
      DBGridImportar.Font.Color:= clBlack;
      DBGridImportar.Canvas.FillRect(Rect);
      DBGridImportar.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmConsultaProgramacao.DBGridProgramacaoCellClick(Column: TColumn);
begin
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)) then
  begin
    try
      if (Self.DBGridProgramacao.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'booleanFavorito') then
      begin
        DBGridProgramacao.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Edit;
        FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
        FieldByName('booleanFavorito').AsBoolean:= not Self.DBGridProgramacao.SelectedField.AsBoolean;
        FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Post;
      end
      else
      DBGridProgramacao.Options:=
      [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
    except
      FrmDataModule.ADOQueryProgramacaoDiaria_Consulta.Cancel;
    end;
  end;
end;

procedure TFrmConsultaProgramacao.DBGridProgramacaoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
    NumCancelados,NumAprovados,NumExecutantes: Integer;
begin
  NumCancelados:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('NumCancelados').AsInteger;
  NumAprovados:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('NumAprovados').AsInteger;
  NumExecutantes:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('NumExecutantes').AsInteger;

  if (Column.Field.FieldName = 'booleanFavorito') then
  begin
    Self.DBGridProgramacao.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridProgramacao.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end
  else if (Column.Field.FieldName = 'NumCancelados') then
  begin
    if  NumCancelados > 0 then
    begin
      DBGridProgramacao.Canvas.Brush.Color:= clRed;
      DBGridProgramacao.Font.Color:= clBlack;
      DBGridProgramacao.Canvas.FillRect(Rect);
      DBGridProgramacao.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumAprovados') then
  begin
    if NumAprovados > 0 then
    begin
      DBGridProgramacao.Canvas.Brush.Color:= clLime;
      DBGridProgramacao.Font.Color:= clBlack;
      DBGridProgramacao.Canvas.FillRect(Rect);
      DBGridProgramacao.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumExecutantes') then
  begin
    if ((NumExecutantes-(NumAprovados+NumCancelados))>0) then
    begin
      DBGridProgramacao.Canvas.Brush.Color:= clYellow;
      DBGridProgramacao.Font.Color:= clBlack;
      DBGridProgramacao.Canvas.FillRect(Rect);
      DBGridProgramacao.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmConsultaProgramacao.DBGridServicosKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridServicos.DataSource.DataSet.RecNo=
  DBGridServicos.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmConsultaProgramacao.DBGridServicosKeyPress(Sender: TObject;
  var Key: Char);
  var
    DataProgramacao: TDateTime;
begin
  DataProgramacao:= FrmDataModule.DataSourceProgramacaoDiaria_Consulta.DataSet.
  FieldByName('DataProgramacao').AsDateTime;
  if DataProgramacao < FrmPrincipal.carregaDataMinima(true) then
    key:= #0;
end;

procedure TFrmConsultaProgramacao.Excluirselecionado2Click(Sender: TObject);
begin
  ListBoxCopiar.Items.Delete(ListBoxCopiar.ItemIndex);
end;

procedure TFrmConsultaProgramacao.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmConsultaProgramacao:=nil;
end;

procedure TFrmConsultaProgramacao.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  listProgramacao:= TStringList.Create;
  FrmPrincipal.MDIChildCreated(self.Handle);
  FrmPrincipal.ProgressBarIncializa(5,'Inicializando....');
  if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) then
  begin
    actExcluirTudo.Enabled:= true;
    actExcluirSelecionado.Enabled:= true;
    actExcluirFiltrados.Enabled:= true;
    actCopiarProgramacaoSELECAO.Enabled:= true;
    actCopiarProgramacaoTODAS.Enabled:= true;
    actImportar.Enabled:= true;
    actInserirProgramacao.Enabled:= true;
    actExcluirRegistroProgramacao.Enabled:= true;
    actExecutanteInserir.Enabled:= true;
    actExecutanteExcluir.Enabled:= true;
    actServicoInserir.Enabled:= true;
    actServicoExcluir.Enabled:= true;
    actGravarPIOP.Enabled:= true;
    actLimparExecutantes.Enabled:= true;
    actCarregarDatas.Enabled:= true;
  end
  else if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)) then
  begin
    actExcluirTudo.Enabled:= false;
    actExcluirSelecionado.Enabled:= true;
    actExcluirFiltrados.Enabled:= true;
    actCopiarProgramacaoSELECAO.Enabled:= true;
    actCopiarProgramacaoTODAS.Enabled:= true;
    actImportar.Enabled:= true;
    actInserirProgramacao.Enabled:= true;
    actExcluirRegistroProgramacao.Enabled:= true;
    actExecutanteInserir.Enabled:= true;
    actExecutanteExcluir.Enabled:= true;
    actServicoInserir.Enabled:= true;
    actServicoExcluir.Enabled:= true;
    actGravarPIOP.Enabled:= false;
    actLimparExecutantes.Enabled:= true;
    actCarregarDatas.Enabled:= true;
  end
  else
  begin
    actExcluirTudo.Enabled:= false;
    actExcluirSelecionado.Enabled:= false;
    actExcluirFiltrados.Enabled:= false;
    actCopiarProgramacaoSELECAO.Enabled:= false;
    actCopiarProgramacaoTODAS.Enabled:= false;
    actImportar.Enabled:= false;
    actInserirProgramacao.Enabled:= false;
    actExcluirRegistroProgramacao.Enabled:= false;
    actExecutanteInserir.Enabled:= false;
    actExecutanteExcluir.Enabled:= false;
    actServicoInserir.Enabled:= false;
    actServicoExcluir.Enabled:= false;
    actGravarPIOP.Enabled:= false;
    actLimparExecutantes.Enabled:= false;
    actCarregarDatas.Enabled:= false;
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  //=========================================
  dataInicio.Date:= IncDay(now,1);
  dataFim.Date:= IncDay(now,1);
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'Plataforma',
  'SELECT Plataforma FROM tblPlataforma ORDER BY Plataforma;',ComboBoxOrigem);
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'TipoEtapaServico',
  'SELECT TipoEtapaServico FROM tblTipoEtapaServico ORDER BY TipoEtapaServico;',ComboBoxSelecaoServicoTipoEtapaServico);
  //Incicializa��o
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.ProgressBarIncremento(1);
  //=========================================
  FrmPrincipal.SetupGridFilterPickListSQL(FrmDataModule.ADOConnectionConsulta,'Origem',
  'SELECT Plataforma FROM tblPlataforma WHERE (BooleanOrigem = True);',
  DBGridExecutantes,false);
  FrmPrincipal.ProgressBarIncremento(1);
  actProcurarProgramacao.Execute;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmConsultaProgramacao.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmConsultaProgramacao.ListBoxCopiarDblClick(Sender: TObject);
begin
  ListBoxCopiar.Items.Delete(ListBoxCopiar.ItemIndex);
end;

procedure TFrmConsultaProgramacao.MonthCalendar1DblClick(Sender: TObject);
begin
  if MonthCalendar1.Date >= FrmPrincipal.carregaDataMinima(false) then
    ListBoxCopiar.Items.Add(FormatDateTime('dd/mm/yyyy',MonthCalendar1.Date))
  else
    MessageBox(0, 'N�o � possivel alterar uma programa��o do passado!',
    'Colibri', MB_ICONERROR);
end;

procedure TFrmConsultaProgramacao.PanelTituloAjuda1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmConsultaProgramacao.PanelTituloAjuda1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmConsultaProgramacao.PanelTituloAjuda1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    FrmPrincipal.Capturing := false;
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmConsultaProgramacao.RadioGroup1Click(Sender: TObject);
begin
  selecionarServicos(ComboBoxSelecaoServicoTipoEtapaServico.Text);
end;


procedure TFrmConsultaProgramacao.RLSelecionarServicosDrawCell(Sender: TObject;
  ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  if (ACol = 0)AND(ARow > 0) then
  begin
    //Selecionado
    if RLSelecionarServicos.Cells[ACol,ARow] = ' ' then
      FrmPrincipal.ImageList1.Draw(RLSelecionarServicos.Canvas,
      Rect.Left+4,Rect.Top+4,194)
    //N�o selecionado
    else if RLSelecionarServicos.Cells[ACol,ARow] = '  ' then
      FrmPrincipal.ImageList1.Draw(RLSelecionarServicos.Canvas,
      Rect.Left+4,Rect.Top+4,190);
  end;
end;

procedure TFrmConsultaProgramacao.RLSelecionarServicosFixedCellClick(
  Sender: TObject; ACol, ARow: Integer);
begin
  FrmPrincipal.clasifica(RLSelecionarServicos,ACol,true);
  AutoFitGrade(RLSelecionarServicos);
end;

procedure TFrmConsultaProgramacao.RLSelecionarServicosKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #32 then
  begin
    if RLSelecionarServicos.Cells[0,RLSelecionarServicos.Row] = '  ' then
      RLSelecionarServicos.Cells[0,RLSelecionarServicos.Row]:= ' '
    else if RLSelecionarServicos.Cells[0,RLSelecionarServicos.Row] = ' ' then
      RLSelecionarServicos.Cells[0,RLSelecionarServicos.Row]:= '  ';

    StatusBarSelecaoServico.Panels[0].Text:=
    'N� Selecionados: '+(IntToStr(FrmPrincipal.CountChecked(RLSelecionarServicos)));
    StatusBarSelecaoServico.Panels[1].Text:=
    'N� Registros: '+(IntToStr(RLSelecionarServicos.RowCount-1));
    AutoFitStatusBar(StatusBarSelecaoServico);
  end
end;

procedure TFrmConsultaProgramacao.RLSelecionarServicosMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
  var
    ACol,ARow: Integer;
    pt: TPoint;
begin
  GetCursorPos(pt);
  pt:= RLSelecionarServicos.ScreenToClient(pt);
  RLSelecionarServicos.MouseToCell(pt.X,pt.Y,ACol, ARow);
  if (ACol = 0)AND(ARow > 0) then
  begin
    if RLSelecionarServicos.Cells[ACol,ARow] = '  ' then
      RLSelecionarServicos.Cells[ACol,ARow]:= ' '
    else if RLSelecionarServicos.Cells[ACol,ARow] = ' ' then
      RLSelecionarServicos.Cells[ACol,ARow]:= '  ';
  end;
  StatusBarSelecaoServico.Panels[0].Text:=
  'N� Selecionados: '+(IntToStr(FrmPrincipal.CountChecked(RLSelecionarServicos)));
  StatusBarSelecaoServico.Panels[1].Text:=
  'N� Registros: '+(IntToStr(RLSelecionarServicos.RowCount-1));
  AutoFitStatusBar(StatusBarSelecaoServico);
end;

procedure TFrmConsultaProgramacao.selecionarServicos(TipoEtapaServico: String);
  var
    CampoSelecao,TextoBreveOP,TextoBreveOM,OrdemManutencao,
    Operacao,CentroTrabalhoOP: String;
    numLinha: Integer;
begin
  actAjudaLimpar.Execute;
  RLSelecionarServicos.ColCount:= 9;
  RLSelecionarServicos.Cells[0,0]:= 'Sele��o';
  RLSelecionarServicos.Cells[1,0]:= 'Servi�o/Ano-Etapa';
  RLSelecionarServicos.Cells[2,0]:= 'Texto Breve da Opera��o';
  RLSelecionarServicos.Cells[3,0]:= 'Texto Breve da Ordem';
  RLSelecionarServicos.Cells[4,0]:= 'Ordem Manuten��o';
  RLSelecionarServicos.Cells[5,0]:= 'Opera��o';
  RLSelecionarServicos.Cells[6,0]:= 'Centro Trabalho';
  RLSelecionarServicos.Cells[7,0]:= 'Data Base Inicio';
  RLSelecionarServicos.Cells[8,0]:= 'Data Base Fim';
  FrmPrincipal.selComboBox(ComboBoxSelecaoServicoTipoEtapaServico,TipoEtapaServico);
  RLSelecionarServicos.RowCount:= 1;
  //Selecionar servi�os j� existentes
  FrmDataModule.ADOQueryProgramacaoServico_Consulta.First;
  while not FrmDataModule.ADOQueryProgramacaoServico_Consulta.Eof do
  begin
    CampoSelecao:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('CampoSelecao').AsString;
    TextoBreveOP:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('TextoBreveOP').AsString;
    TextoBreveOM:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('TextoBreveOM').AsString;
    OrdemManutencao:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('OrdemManutencao').AsString;
    Operacao:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('Operacao').AsString;
    CentroTrabalhoOP:= FrmDataModule.DataSourceProgramacaoServico_Consulta.DataSet.
    FieldByName('CentroTrabalhoOP').AsString;
    if FrmPrincipal.selecionarServico(RLSelecionarServicos,CampoSelecao,TextoBreveOP,OrdemManutencao) = false then
    begin
      numLinha:= FrmPrincipal.InsertRow1(RLSelecionarServicos,1);
      //Preencher valores
      RLSelecionarServicos.Cells[0,numLinha]:= ' ';
      RLSelecionarServicos.Cells[1,numLinha]:= CampoSelecao;
      RLSelecionarServicos.Cells[2,numLinha]:= TextoBreveOP;
      RLSelecionarServicos.Cells[3,numLinha]:= TextoBreveOM;
      RLSelecionarServicos.Cells[4,numLinha]:= OrdemManutencao;
      RLSelecionarServicos.Cells[5,numLinha]:= Operacao;
      RLSelecionarServicos.Cells[6,numLinha]:= CentroTrabalhoOP;
      RLSelecionarServicos.Cells[7,numLinha]:= '';
      RLSelecionarServicos.Cells[8,numLinha]:= '';
    end;
    FrmDataModule.ADOQueryProgramacaoServico_Consulta.Next;
  end;
  FrmDataModule.ADOQueryProgramacaoServico_Consulta.First;
  //=================================================================
  AutoFitGrade(RLSelecionarServicos);
  try
    RLSelecionarServicos.FixedRows:= 1;
  except
    RLSelecionarServicos.FixedRows:= 0;
  end;
  //================================================
  StatusBarSelecaoServico.Panels[0].Text:=
  'N� Selecionados: '+(IntToStr(FrmPrincipal.CountChecked(RLSelecionarServicos)));
  StatusBarSelecaoServico.Panels[1].Text:=
  'N� Registros: '+(IntToStr(RLSelecionarServicos.RowCount-1));
  AutoFitStatusBar(StatusBarSelecaoServico);
  //================================================
  PanelTituloAjuda1.Caption:= 'Sele��o de servi�os disponiv�is';
  PanelSelecaoServico.Visible:= true;
  PanelSelecaoServico.Align:= alClient;
  PanelAjuda.Width:= 1000;
  PanelAjuda.Height:= 350;
  PanelAjuda.Visible:= true;

end;

procedure TFrmConsultaProgramacao.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
var
  active: TWinControl;
  idx: Integer;
begin
  active := FindControl(msg.ActiveWnd) ;

  if Assigned(active) then
  begin
    idx := FrmPrincipal.mdiChildrenTabs.Tabs.IndexOfObject(TObject(msg.ActiveWnd));
    FrmPrincipal.mdiChildrenTabs.Tag := -1;
    FrmPrincipal.mdiChildrenTabs.TabIndex := idx;
    FrmPrincipal.mdiChildrenTabs.Tag := 0;
  end;
end;

end.



