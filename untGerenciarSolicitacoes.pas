unit untGerenciarSolicitacoes;

interface

uses
  Winapi.Windows, Winapi.Messages, SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.ToolWin, Vcl.ComCtrls,
  Data.DB, Vcl.Grids, Vcl.DBGrids, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.Mask, Vcl.DBCtrls,DateUtils, Vcl.Menus, Vcl.CheckLst, ADODB, SDL_NumIO,
  Clipbrd,
  untDBGridFilter, uZucchi, uProgramacaoRTUtils;

type
  TFrmGerenciarSolicitacoes = class(TForm)
    PanelTitulo: TPanel;
    ActionManager1: TActionManager;
    actAprovarSelecionado: TAction;
    actCancelarSelecionado: TAction;
    actAjudaLimpar: TAction;
    PanelResultados: TPanel;
    Splitter1: TSplitter;
    PanelServicos: TPanel;
    PanelExecutantes: TPanel;
    Panel6: TPanel;
    ToolBar2: TToolBar;
    btnAprovar: TBitBtn;
    btnCancelar: TBitBtn;
    ToolButton1: TToolButton;
    Panel4: TPanel;
    PanelDestinosOrigens: TPanel;
    PanelTituloDestinoOrigem: TPanel;
    ToolBar5: TToolBar;
    btnDestino: TToolButton;
    btnOrigem: TToolButton;
    actDestino: TAction;
    actOrigem: TAction;
    actProcurar: TAction;
    DBEditidProgramacaoExecutante: TDBEdit;
    actSelecaoTodos: TAction;
    actSelecaoLimpar: TAction;
    ToolButton4: TToolButton;
    ToolButton5: TToolButton;
    actDestinoOrigem: TAction;
    actOcultar: TAction;
    actMostrar: TAction;
    btnMostrarOcultar: TBitBtn;
    StatusBarExecutantes: TStatusBar;
    actAbrirServico: TAction;
    actZoomMais: TAction;
    actZoomMenos: TAction;
    ToolButton2: TToolButton;
    ToolButton3: TToolButton;
    actAutoFit: TAction;
    actStatusTodos: TAction;
    RLDestinoOrigem: TStringGrid;
    PopupMenuDestinoOrigem: TPopupMenu;
    StatusdaprogramaoTODOS1: TMenuItem;
    StatusdaprogramaoTODOS2: TMenuItem;
    actCarregarDestino: TAction;
    actInfo: TAction;
    ToolButton7: TToolButton;
    ColunasLayoutExecutantes: TStringGrid;
    DateProgramacao: TDateTimePicker;
    actTransbordos: TAction;
    actMatrizOrigemDestino: TAction;
    BitBtn6: TBitBtn;
    btnCopiarImagemDestinoOrigem: TBitBtn;
    BitBtn7: TBitBtn;
    actContadorSolicitacao: TAction;
    actCopiarKITPS: TAction;
    btnTransbordo: TBitBtn;
    PanelAjuda: TPanel;
    PanelTituloAjuda1: TPanel;
    SpeedButton4: TSpeedButton;
    PanelAlterarDestinos: TPanel;
    ToolBar1: TToolBar;
    BitBtn13: TBitBtn;
    RLAlterarDestinos: TStringGrid;
    StatusBarAlterarDestinos: TStatusBar;
    actCarregarDestinoJanela: TAction;
    ComboBoxDestino: TComboBox;
    actCarregarDestinoProgramacao: TAction;
    BitBtn2: TBitBtn;
    actExcluirProgramacao: TAction;
    btnKITPS: TBitBtn;
    actexcluirExecutante: TAction;
    PopupMenuExcluir: TPopupMenu;
    Excluirprogramaoexecutanteselecionado1: TMenuItem;
    Excluirprogramaocompletaselecionada1: TMenuItem;
    PopupMenuForaOperacao: TPopupMenu;
    actMudanca: TAction;
    btnMudanca: TBitBtn;
    actDuplicados: TAction;
    PanelNotas: TPanel;
    Panel1: TPanel;
    ToolBar3: TToolBar;
    DBMemoNotas: TDBMemo;
    DBNavigator1: TDBNavigator;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    Splitter3: TSplitter;
    btnExcelExecutante: TToolButton;
    Panel18: TPanel;
    DBGridExecutantes: TFilterDBGrid;
    PanelContadorSolicitacao: TPanel;
    RLContadorSolicitacao: TStringGrid;
    PanelConfigurarContador: TPanel;
    ToolBar6: TToolBar;
    DBNavigatorComissao: TDBNavigator;
    Panel2: TPanel;
    Panel48: TPanel;
    Panel49: TPanel;
    Panel50: TPanel;
    DBEdit1: TDBEdit;
    Panel3: TPanel;
    Panel5: TPanel;
    Panel7: TPanel;
    DBEdit4: TDBEdit;
    Panel8: TPanel;
    Panel9: TPanel;
    Panel10: TPanel;
    DBEdit5: TDBEdit;
    Panel11: TPanel;
    Panel12: TPanel;
    Panel13: TPanel;
    DBEdit7: TDBEdit;
    Panel14: TPanel;
    Panel15: TPanel;
    Panel16: TPanel;
    DBEdit6: TDBEdit;
    Panel17: TPanel;
    DBCheckBox1: TDBCheckBox;
    DBGridContador: TFilterDBGrid;
    ColunasLayoutContador: TStringGrid;
    actConfigurarContador: TAction;
    ToolButton11: TToolButton;
    Panel19: TPanel;
    DBGridServicos: TFilterDBGrid;
    ColunasLayoutServico: TStringGrid;
    btnFiltroClearExecutantes: TToolButton;
    btnClearFiltroContador: TToolButton;
    btnExcelContador: TToolButton;
    ToolButton6: TToolButton;
    actExcelProgramacao: TAction;
    actCopyImage: TAction;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridServicosDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actCancelarSelecionadoExecute(Sender: TObject);
    procedure PanelTituloAjudaMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure btnDestinoClick(Sender: TObject);
    procedure btnOrigemClick(Sender: TObject);
    procedure actDestinoExecute(Sender: TObject);
    procedure actOrigemExecute(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
    procedure DBGridExecutantesDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGridExecutantesCellClick(Column: TColumn);
    procedure actSelecaoTodosExecute(Sender: TObject);
    procedure actSelecaoLimparExecute(Sender: TObject);
    procedure actDestinoOrigemExecute(Sender: TObject);
    procedure edtTipoEtapaServicoKeyPress(Sender: TObject; var Key: Char);
    procedure edtNomeExecutanteKeyPress(Sender: TObject; var Key: Char);
    procedure edtDestinoKeyPress(Sender: TObject; var Key: Char);
    procedure actOcultarExecute(Sender: TObject);
    procedure actMostrarExecute(Sender: TObject);
    procedure actAbrirServicoExecute(Sender: TObject);
    procedure actZoomMaisExecute(Sender: TObject);
    procedure actZoomMenosExecute(Sender: TObject);
    procedure actAutoFitExecute(Sender: TObject);
    procedure RLDestinoOrigemSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure actStatusTodosExecute(Sender: TObject);
    procedure RLDestinoOrigemDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure DBGridServicosKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure actAprovarSelecionadoExecute(Sender: TObject);
    procedure actCarregarDestinoExecute(Sender: TObject);
    procedure DBGridExecutantesKeyPress(Sender: TObject; var Key: Char);
    procedure actInfoExecute(Sender: TObject);
    procedure DBGridExecutantesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure actMatrizOrigemDestinoExecute(Sender: TObject);
    procedure actTransbordosExecute(Sender: TObject);
    procedure DateProgramacaoCloseUp(Sender: TObject);
    procedure PanelTituloMagicMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure ComboBoxExecutanteKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridServicosKeyPress(Sender: TObject; var Key: Char);
    procedure actContadorSolicitacaoExecute(Sender: TObject);
    procedure actCopiarKITPSExecute(Sender: TObject);
    procedure actCarregarDestinoJanelaExecute(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure RLAlterarDestinosSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ComboBoxDestinoMouseLeave(Sender: TObject);
    procedure ComboBoxDestinoCloseUp(Sender: TObject);
    procedure actCarregarDestinoProgramacaoExecute(Sender: TObject);
    procedure actExcluirProgramacaoExecute(Sender: TObject);
    procedure actexcluirExecutanteExecute(Sender: TObject);
    procedure actAjudaLimparExecute(Sender: TObject);
    procedure PanelTituloAjuda1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PanelTituloAjuda1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PanelTituloAjuda1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure actMudancaExecute(Sender: TObject);
    procedure actDuplicadosExecute(Sender: TObject);
    procedure DBGridContadorDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actConfigurarContadorExecute(Sender: TObject);
    procedure DBGridContadorCellClick(Column: TColumn);
    procedure actExcelProgramacaoExecute(Sender: TObject);
    procedure actCopyImageExecute(Sender: TObject);
    type
      TPlataforma = record
        Salvatagem,GD,BCI,Surfer,SOV,Aqua: String;
      end;
    type
      TMemoriaExecutante = record
        idProgramacaoExecutante: Integer;
        Origem, NomeExecutante, Funcao, Empresa,
        CodigoSAP, Documento, RequisitantePT: String
    end;
  private
    { Private declarations }
    abrirDestinoOrigem: Boolean;
    MatrizPlataforma,MatrizConsulta: array of array of String;
    DestinoMudanca: String;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure GravarStatus(StatusProgramacao: String);
    function TotalDestino(Destino,StatusProgramacao: String): Integer;
    function TotalOrigem(Origem,StatusProgramacao: String): Integer;
    function TotalPOB(Destino,StatusProgramacao: String): Integer;
    function Contar_TS_OP(TipoEtapaServico,Destino: String): Integer;
    function Contar_KIT_PS(Destino: String): Integer;
    function InfoPlataforma(Plataforma: String): TPlataforma;
    function SalvatagemPlataforma(Plataforma: String): String;
    procedure CopiarGridComoImagemParaClipboard(AGrid: TStringGrid);

  public
    procedure StatusLinhaSelecionada;
    { Public declarations }
  end;

var
  FrmGerenciarSolicitacoes: TFrmGerenciarSolicitacoes;

implementation
  uses untPrincipal,untDataModule, untFrmTabela, untConsultaExecutantesProgramados,
  untMotivoCancelamento;
{$R *.dfm}

type
  TWinControlAccess = class(TWinControl);

procedure TFrmGerenciarSolicitacoes.actAbrirServicoExecute(Sender: TObject);
  var
    codigoProgramacao: Integer;
begin
  if (DBGridServicos.Visible)AND(DBEditidProgramacaoExecutante.Text<>'') then
  begin
    codigoProgramacao:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('idProgramacaoDiaria').AsInteger;

    FrmDataModule.ADOQueryGerenciarServico.Active := false;
    FrmDataModule.ADOQueryGerenciarServico.Parameters.Items[0].Value:= codigoProgramacao;
    FrmDataModule.ADOQueryGerenciarServico.Active := true;
  end;
end;

procedure TFrmGerenciarSolicitacoes.actAjudaLimparExecute(Sender: TObject);
begin
  PanelAjuda.Visible:= false;
  PanelAlterarDestinos.Visible:= false;
  PanelConfigurarContador.Visible:= false;
end;

procedure TFrmGerenciarSolicitacoes.actAprovarSelecionadoExecute(
  Sender: TObject);
begin
  GravarStatus('Aprovado');
  actContadorSolicitacao.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actAutoFitExecute(Sender: TObject);
begin
  FrmPrincipal.AutoSizeDBGrid(DBGridExecutantes);
  FrmPrincipal.AutoSizeDBGrid(DBGridServicos);
  AutoFitGrade(RLDestinoOrigem);
  AutoFitStatusBar(StatusBarExecutantes);
end;

procedure TFrmGerenciarSolicitacoes.actOcultarExecute(Sender: TObject);
begin
  btnMostrarOcultar.Action:= actMostrar;
  PanelServicos.Height:= 150;
  DBGridServicos.Visible:= true;
  actAbrirServico.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actCancelarSelecionadoExecute(
  Sender: TObject);
begin
  if not Assigned(FrmMotivoCancelamento) then
    FrmMotivoCancelamento:= TFrmMotivoCancelamento.Create(Application)
  else
    FrmMotivoCancelamento.Show;

  FrmMotivoCancelamento.Iniciar(0);
end;

procedure TFrmGerenciarSolicitacoes.actCarregarDestinoExecute(Sender: TObject);
  var
    idProgramacaoExecutante,CodigoProgramacaoDiariaNova,CodigoProgramacaoDiariaAlterada,
    NumCancelados,NumAprovados,NumExecutantes,i: Integer;
    txtTipoEtapaServico,NovoDestino,DataProgramacao: String;
begin
  for i := 1 to RLAlterarDestinos.RowCount-1 do
  begin
    if RLAlterarDestinos.Cells[1,i] <> DestinoMudanca then
    begin
      NovoDestino:= RLAlterarDestinos.Cells[1,i];
      idProgramacaoExecutante:= StrToInt(RLAlterarDestinos.Cells[5,i]);
      txtTipoEtapaServico:= RLAlterarDestinos.Cells[3,i];
      DataProgramacao:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
      FieldByName('DataProgramacao').AsString;
      //Incluir Programação
      CodigoProgramacaoDiariaNova:= FrmPrincipal.incluirProgramacao(
      DataProgramacao,NovoDestino,txtTipoEtapaServico,FrmPrincipal.logChave,
      FrmPrincipal.logMaquina,now);
      //Abrir registro do executante a ser alterado
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
      idProgramacaoExecutante;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
      //Incluir Executante
      FrmDataModule.ADOQueryInserirExecutante1.Insert;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('CodigoProgramacaoDiaria').AsInteger:= CodigoProgramacaoDiariaNova;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('NomeExecutante').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('NomeExecutante').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('Origem').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('Origem').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('Funcao').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('Funcao').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('Empresa').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('Empresa').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('CodigoSAP').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('CodigoSAP').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('Documento').AsString:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('Documento').AsString;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('RequisitantePT').AsBoolean:= FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('RequisitantePT').AsBoolean;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('StatusProgramacao').AsString:= 'Aprovado';
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('AvaliadoPorProgramacao').AsString:= FrmPrincipal.logChave;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('DataAvaliacaoProgramacao').AsDateTime:= now;
      FrmDataModule.DataSourceInserirExecutante.DataSet.
      FieldByName('ComputadorProgramacao').AsString:= FrmPrincipal.logMaquina;
      FrmDataModule.ADOQueryInserirExecutante1.Post;
      //Incluir Serviço Padrão
      FrmDataModule.ADOQueryInserirServico.Insert;
      FrmDataModule.DataSourceInserirServico.DataSet.
      FieldByName('CodigoProgramacaoDiaria').AsInteger:= CodigoProgramacaoDiariaNova;
      FrmDataModule.DataSourceInserirServico.DataSet.
      FieldByName('TextoBreveOP').AsString:= FrmPrincipal.incluirServicoPadrao(txtTipoEtapaServico);;
      FrmDataModule.ADOQueryInserirServico.Post;
      //Excluir executante da programação alterada
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
      idProgramacaoExecutante;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Delete;
    end;
  end;
  //Altualiza busca
  PanelAjuda.Visible:= false;
  actDestinoOrigem.Execute;
  //Atualizar contagem de executantes da programação alterada
  CodigoProgramacaoDiariaAlterada:= StrToInt(RLAlterarDestinos.Cells[6,1]);
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:=CodigoProgramacaoDiariaAlterada;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;
  NumExecutantes:= FrmPrincipal.CalcNumExecutantes(CodigoProgramacaoDiariaAlterada);
  //Caso não tenha executantes, excluir a programação completa
  if NumExecutantes = 0 then
  begin
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Delete;
  end
  //Caso ainda tenha executantes, atualizar os numeros
  else
  begin
    NumCancelados:= FrmPrincipal.CalcNumCanceladosAprovado(CodigoProgramacaoDiariaAlterada,0);
    NumAprovados:= FrmPrincipal.CalcNumCanceladosAprovado(CodigoProgramacaoDiariaAlterada,1);
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumExecutantes').AsInteger:= NumExecutantes;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumCancelados').AsInteger:= NumCancelados;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumAprovados').AsInteger:= NumAprovados;
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Post;
  end;
end;

procedure TFrmGerenciarSolicitacoes.actCarregarDestinoJanelaExecute(
  Sender: TObject);
  var
    numLinhas: Integer;
    idProgramacaoDiaria,txtTipoEtapaServico: String;
begin
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'Plataforma',
  'SELECT Plataforma FROM tblPlataforma ORDER BY Plataforma;',ComboBoxDestino);
  DestinoMudanca:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
  FieldByName('txtDestino').AsString;
  idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
  FieldByName('idProgramacaoDiaria').AsString;
  txtTipoEtapaServico:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
  FieldByName('txtTipoEtapaServico').AsString;
  FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Active := false;
  FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Parameters.Items[0].Value:= idProgramacaoDiaria;
  FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Active := true;
  RLAlterarDestinos.FixedRows:= 0;
  RLAlterarDestinos.RowCount:= 1;
  RLAlterarDestinos.ColCount:= 7;
  RLAlterarDestinos.Cells[0,0]:= 'Origem';
  RLAlterarDestinos.Cells[1,0]:= 'Destino';
  RLAlterarDestinos.Cells[2,0]:= 'Nome do Executante';
  RLAlterarDestinos.Cells[3,0]:= 'Tipo de Etapa de Serviço';
  RLAlterarDestinos.Cells[4,0]:= 'Função';
  RLAlterarDestinos.Cells[5,0]:= 'Código Executante';
  RLAlterarDestinos.Cells[6,0]:= 'Código Programação';
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.RecordCount,
  'Carregando Executantes...');
  while not FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Eof do
  begin
    numLinhas:= RLAlterarDestinos.RowCount;
    RLAlterarDestinos.RowCount:= numLinhas+1;
    //Preencher valores
    RLAlterarDestinos.Cells[0,numLinhas]:= FrmDataModule.
    DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Origem').AsString;
    RLAlterarDestinos.Cells[1,numLinhas]:= DestinoMudanca;
    RLAlterarDestinos.Cells[2,numLinhas]:= FrmDataModule.
    DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('NomeExecutante').AsString;
    RLAlterarDestinos.Cells[3,numLinhas]:= txtTipoEtapaServico;
    RLAlterarDestinos.Cells[4,numLinhas]:= FrmDataModule.
    DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('Funcao').AsString;
    RLAlterarDestinos.Cells[5,numLinhas]:= FrmDataModule.
    DataSourceProgramacaoExecutante_Consulta.DataSet.FieldByName('idProgramacaoExecutante').AsString;
    RLAlterarDestinos.Cells[6,numLinhas]:= idProgramacaoDiaria;
    FrmDataModule.ADOQueryProgramacaoExecutante_Consulta.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;

  StatusBarAlterarDestinos.Panels[0].Text:= 'N° registros: '+
  intToStr(RLAlterarDestinos.RowCount-1);
  AutoFitGrade(RLAlterarDestinos);
  FrmPrincipal.ProgressBarAtualizar;
  RLAlterarDestinos.ColWidths[1] := 90;
  AutoFitStatusBar(StatusBarAlterarDestinos);
  try
    RLAlterarDestinos.FixedRows:= 1;
  except
    RLAlterarDestinos.FixedRows:= 0;
  end;
  actAjudaLimpar.Execute;
  PanelAlterarDestinos.Align:= alClient;
  PanelAlterarDestinos.Visible:= true;
  PanelAjuda.Visible:= true;
  PanelAjuda.Width:= 650;
  PanelAjuda.Height:= 300;
  PanelAjuda.Top:= 180;
  PanelAjuda.Left:= 400;
end;

procedure TFrmGerenciarSolicitacoes.actCarregarDestinoProgramacaoExecute(
  Sender: TObject);
  var
    CodigoProgramacaoDiaria,i: Integer;
    NovoDestino: String;
begin
  for i := 1 to RLAlterarDestinos.RowCount-1 do
  begin
    if RLAlterarDestinos.Cells[1,i] <> DestinoMudanca then
    begin
      NovoDestino:= RLAlterarDestinos.Cells[1,i];
      CodigoProgramacaoDiaria:= StrToInt(RLAlterarDestinos.Cells[6,1]);
      FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
      FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:=CodigoProgramacaoDiaria;
      FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;
      FrmDataModule.ADOQueryConsultaProgramacao_ID.Edit;
      FrmDataModule.ADOQueryConsultaProgramacao_ID.FieldByName('txtDestino').AsString:= NovoDestino;
      FrmDataModule.ADOQueryConsultaProgramacao_ID.Post;
      actDestinoOrigem.Execute;
      PanelAjuda.Visible:= false;
      Break
    end;
  end;
end;

procedure TFrmGerenciarSolicitacoes.actCopiarKITPSExecute(Sender: TObject);
  var
   DataProcura,SQLBase,NomeExecutante,txtDestino,Mensagem: String;
   ListaNaoEncontrado: TStringList;
   I: integer;
begin
  DataProcura:= FormatDateTime('dd/mm/yyyy',now);
  if InputQuery('Data da Cópia','Entre com a data da programação modelo de replicação',
  DataProcura) then
  begin
    if FrmPrincipal.isData(DataProcura) then
    begin
      ListaNaoEncontrado:= TStringList.Create;
      ListaNaoEncontrado.Clear;
      btnFiltroClearExecutantes.Click;
      SQLBase:=
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
      'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
      'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+') '+
      'AND (tblProgramacaoExecutante.Kit_PS LIKE TRUE)) '+
      'ORDER BY txtDestino,NomeExecutante,Origem;';
      FrmDataModule.ADOQueryTemporarioDBColibri.Close;
      FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
      FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
      FrmDataModule.ADOQueryTemporarioDBColibri.Open;
      while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
      begin
        txtDestino:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
        FieldByName('txtDestino').AsString;
        NomeExecutante:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
        FieldByName('NomeExecutante').AsString;
        FrmPrincipal.buscaFiledGrid1('Origem',FrmPrincipal.OrigemPlataformas,'Exato',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('NomeExecutante',NomeExecutante,'Exato',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('txtDestino',txtDestino,'Exato',ColunasLayoutExecutantes,4,false);
        actProcurar.Execute;
        if not FrmDataModule.ADOQueryGerenciarExecutante.IsEmpty then
        begin
          FrmDataModule.ADOqueryGerenciarExecutante.Edit;
          FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('Kit_PS').AsBoolean:= True;
          FrmDataModule.ADOqueryGerenciarExecutante.Post;
        end
        else
        begin
          ListaNaoEncontrado.Add(txtDestino+': '+NomeExecutante)
        end;
        FrmDataModule.ADOQueryTemporarioDBColibri.Next;
      end;
      btnFiltroClearExecutantes.Click;
      FrmPrincipal.buscaFiledGrid1('Kit_PS','True','Exato',ColunasLayoutExecutantes,4,false);
      actProcurar.Execute;
      if ListaNaoEncontrado.Count > 0 then
      begin
        Mensagem:= '';
        for I := 0 to ListaNaoEncontrado.Count -1 do
          Mensagem:= Mensagem+#13+IntToStr(i+1)+'. '+ListaNaoEncontrado[i];

        ShowMessage('NÃO ENCONTRADOS:'+#13+Mensagem);
      end;
    end;
  end;

end;

procedure TFrmGerenciarSolicitacoes.actCopyImageExecute(Sender: TObject);
begin
  CopiarGridComoImagemParaClipboard(RLDestinoOrigem);
  Application.MessageBox(
    'Imagem da grade copiada para a área de transferência.',
    'Colibri',
    MB_ICONINFORMATION
  );
end;

procedure TFrmGerenciarSolicitacoes.actConfigurarContadorExecute(
  Sender: TObject);
begin
  actAjudaLimpar.Execute;
  PanelTituloAjuda1.Caption:= 'Configuração de Contadores';
  PanelConfigurarContador.Align:= alClient;
  PanelConfigurarContador.Visible:= true;
  PanelAjuda.Visible:= true;
  PanelAjuda.Width:= 550;
  PanelAjuda.Height:= 400;
  PanelAjuda.Top:= 180;
  PanelAjuda.Left:= 400;
end;

procedure TFrmGerenciarSolicitacoes.actContadorSolicitacaoExecute(Sender: TObject);
  var
    i,Coluna,Total: Integer;
    TipoEtapaServico,Destino,Origem,StatusProgramacao,Descricao: String;
    MatrizTipoEtapaServico,MatrizDestino,MatrizOrigem,MatrizStatusProgramacao: String;
    BooleanOrigemDestino,MatrizBooleanOrigemDestino: Boolean;
begin
  FrmDataModule.ADOQueryContadorSolicitacao.Active:= false;
  FrmDataModule.ADOQueryContadorSolicitacao.Active:= true;
  FrmDataModule.DataSourceContadorSolicitacao.Enabled:= false;
  FrmDataModule.ADOQueryContadorSolicitacao.First;
  Coluna:= 0;
  RLContadorSolicitacao.ColCount:= 1;
  RLContadorSolicitacao.Rows[0].Clear;
  while not FrmDataModule.ADOQueryContadorSolicitacao.Eof do
  begin
    Descricao:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('Descricao').AsString;
    Origem:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('Origem').AsString;
    Destino:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('Destino').AsString;
    TipoEtapaServico:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('TipoEtapaServico').AsString;
    StatusProgramacao:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('StatusProgramacao').AsString;
    BooleanOrigemDestino:= FrmDataModule.DataSourceContadorSolicitacao.DataSet.FieldByName('BooleanOrigemDestino').AsBoolean;
    //==========================================================================
    Total:= 0;
    for i := 0 to High(MatrizConsulta[0]) do
    begin
      MatrizOrigem:= MatrizConsulta[0,i];
      MatrizDestino:= MatrizConsulta[1,i];
      MatrizStatusProgramacao:= MatrizConsulta[2,i];
      MatrizTipoEtapaServico:= MatrizConsulta[3,i];
      if ((BooleanOrigemDestino)AND(MatrizConsulta[0,i]<>MatrizConsulta[1,i]))
      OR((BooleanOrigemDestino=false)) then
        MatrizBooleanOrigemDestino:= true
      else
        MatrizBooleanOrigemDestino:= false;

      if (FrmPrincipal.buscarString(Origem,MatrizOrigem)AND
      FrmPrincipal.buscarString(Destino,MatrizDestino)AND
      FrmPrincipal.buscarString(StatusProgramacao,MatrizStatusProgramacao)AND
      FrmPrincipal.buscarString(TipoEtapaServico,MatrizTipoEtapaServico)AND(MatrizBooleanOrigemDestino)) then
        Total:= Total+1;
    end;
    RLContadorSolicitacao.Cells[Coluna,0]:= Descricao+': '+IntToStr(Total);
    Coluna:= RLContadorSolicitacao.ColCount;
    RLContadorSolicitacao.ColCount:= RLContadorSolicitacao.ColCount+1;
    FrmDataModule.ADOQueryContadorSolicitacao.Next;
  end;
  AutoFitGrade(RLContadorSolicitacao);
  FrmDataModule.ADOQueryContadorSolicitacao.First;
  FrmDataModule.DataSourceContadorSolicitacao.Enabled:= true;
end;

procedure TFrmGerenciarSolicitacoes.actMatrizOrigemDestinoExecute(Sender: TObject);
  var
    i,NumRegistros: Integer;
    DataProcura,SQLBase: String;
begin
  DataProcura:= (FormatDateTime('dd/mm/yyyy',DateProgramacao.Date));//DateToStr(DateProgramacao.Date);
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE "'+Dataprocura+'"))'+
  ' ORDER BY tblProgramacaoDiaria.txtDestino;';
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  NumRegistros:= FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount;
  //Configurar matriz
  MatrizConsulta:= nil;
  SetLength(MatrizConsulta, 5);
  for i := 0 to High(MatrizConsulta) do
    SetLength(MatrizConsulta[i], NumRegistros);
  //Carregar Matriz
  i:= 0;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    MatrizConsulta[0,i]:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('Origem').AsString;
    MatrizConsulta[1,i]:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('txtDestino').AsString;
    MatrizConsulta[2,i]:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('StatusProgramacao').AsString;
    MatrizConsulta[3,i]:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('txtTipoEtapaServico').AsString;
    MatrizConsulta[4,i]:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('Kit_PS').AsString;
    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
    i:= i+1;
  end;
end;

procedure TFrmGerenciarSolicitacoes.actTransbordosExecute(Sender: TObject);
begin
  if not Assigned(FrmTabela) then
    Application.CreateForm(TFrmTabela, FrmTabela);
  FrmTabela.Show;   // ou ShowModal conforme o caso

  TProgramacaoRTUtils.PainelTransbordos(DateToStr(DateProgramacao.Date));
end;

procedure TFrmGerenciarSolicitacoes.actDestinoOrigemExecute(Sender: TObject);
  var
    DataProcura,SQLBase: String;
begin
  FrmPrincipal.actMatrizForaOperacao.Execute;
  if btnDestino.Down then
  begin
    PanelTituloDestinoOrigem.Caption:= 'Destinos Programados';
    actDestino.Execute;
  end
  else
  begin
    PanelTituloDestinoOrigem.Caption:= 'Origens Programadas';
    actOrigem.Execute;
  end;
  actStatusTodos.Execute;
  DataProcura:= (FormatDateTime('dd/mm/yyyy',DateProgramacao.Date));
  SQLBase:= 'SELECT tblProgramacaoNotas.*  FROM tblProgramacaoNotas '+
  'WHERE ((DataProgramacao LIKE '+QuotedStr(Dataprocura)+'));';
  FrmDataModule.ADOQueryProgramacaoNotas.Close;
  FrmDataModule.ADOQueryProgramacaoNotas.SQL.Clear;
  FrmDataModule.ADOQueryProgramacaoNotas.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryProgramacaoNotas.Open;
  if FrmDataModule.ADOQueryProgramacaoNotas.IsEmpty then
  begin
    DataProcura:= FrmPrincipal.corrigirData(DateProgramacao.Date);
    FrmDataModule.ADOQueryProgramacaoNotas.Insert;
    FrmDataModule.ADOQueryProgramacaoNotas.FieldByName
    ('DataProgramacao').AsString:= DataProcura;
    FrmDataModule.ADOQueryProgramacaoNotas.Post;
  end;

  FrmDataModule.ADOQueryInserirServico.Active:= false;
  FrmDataModule.ADOQueryInserirServico.Active:= true;

  FrmDataModule.ADOQueryInserirExecutante1.Active:= false;
  FrmDataModule.ADOQueryInserirExecutante1.Active:= true;
end;

procedure TFrmGerenciarSolicitacoes.actDuplicadosExecute(Sender: TObject);
begin
  TProgramacaoRTUtils.PainelTransbordos(DateToStr(DateProgramacao.Date));
end;

procedure TFrmGerenciarSolicitacoes.actExcelProgramacaoExecute(Sender: TObject);
begin
  ExcelStringGrid(RLDestinoOrigem,'Programação',
  'Data: '+DateToStr(DateProgramacao.DateTime),'','',true,
  FrmPrincipal.ProgressBarPrincipal,FrmPrincipal.MemoPrincipal);
end;

procedure TFrmGerenciarSolicitacoes.actexcluirExecutanteExecute(
  Sender: TObject);
  var
    idProgramacaoExecutante: Integer;
begin
  idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
  FieldByName('idProgramacaoExecutante').AsInteger;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:= idProgramacaoExecutante;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= true;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Delete;
  actDestinoOrigem.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actExcluirProgramacaoExecute(
  Sender: TObject);
  var
    idProgramacaoDiaria: Integer;
begin
  idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.DataSet.
  FieldByName('idProgramacaoDiaria').AsInteger;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:= idProgramacaoDiaria;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Delete;
  actDestinoOrigem.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actInfoExecute(Sender: TObject);
begin
  ShowMessage('∑ Apro. = somatório de executantes aprovados'+#13+
  '∑ Canc. = somatório de executantes cancelados'+#13+
  'Saldo = executantes que entram (Destino) na plataforma menos os executantes '+
  'que saem da plataforma (Origem) somado aos executantes que (Permanecem) na plataforma  (Destino-Origem+Permanecem)'+#13+
  'Salv. = salvatagem da plataforma'+#13+
  'GD = Condição Operacional de Guindastes'+#13+
  'BCI = Condição Operacional de BCI'+#13+#13+
  'Ac. = Condição Operacional do Acesso a Plataforma'+#13+#13+
  'Limitação de Programação'+#13+
  '* 11 Plataformas para 2 Campos'+#13+
  '* 10 Plataformas para 3 Campos'+#13+
  '* 9 Plataformas para 4 Campos'+#13+
  '* 9 Plataformas para 2 Campos + PRB-01');
end;

procedure TFrmGerenciarSolicitacoes.actDestinoExecute(Sender: TObject);
  var
    DataProcura,Destino: String;
    numLinhas,NumRegistros,i: Integer;
begin
  DataProcura:= DateToStr(DateProgramacao.Date);
  //============================================================
  FrmDataModule.ADOQueryGerenciarSolicitacoes.Close;
  FrmDataModule.ADOQueryGerenciarSolicitacoes.SQL.Clear;
  FrmDataModule.ADOQueryGerenciarSolicitacoes.SQL.Add(
  'SELECT tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao '+
  'FROM tblProgramacaoDiaria '+
  'GROUP BY tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao '+
  'HAVING (tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(DataProcura)+');');
  FrmDataModule.ADOQueryGerenciarSolicitacoes.Open;
  //=============================================================
  RLDestinoOrigem.ColCount:= 11;
  RLDestinoOrigem.FixedRows:= 0;
  RLDestinoOrigem.RowCount:= 1;
  RLDestinoOrigem.Cells[0,0]:= 'Destino';
  RLDestinoOrigem.Cells[1,0]:= '∑ Apro.';
  RLDestinoOrigem.Cells[2,0]:= '∑ Canc.';
  RLDestinoOrigem.Cells[3,0]:= '∑ Mud.';
  RLDestinoOrigem.Cells[4,0]:= 'TS';
  RLDestinoOrigem.Cells[5,0]:= 'OP';
  RLDestinoOrigem.Cells[6,0]:= 'GD ';
  RLDestinoOrigem.Cells[7,0]:= 'BCI';
  {RLDestinoOrigem.Cells[8,0]:= 'Surfer';
  RLDestinoOrigem.Cells[9,0]:= 'SOV';
  RLDestinoOrigem.Cells[10,0]:= 'Áqua';}
  RLDestinoOrigem.Cells[8,0]:= 'Saldo';
  RLDestinoOrigem.Cells[9,0]:= 'Salv.';
  RLDestinoOrigem.Cells[10,0]:= 'KIT';
  //Configurar matriz de dados das plataformas
  NumRegistros:= FrmDataModule.ADOQueryGerenciarSolicitacoes.RecordCount;
  MatrizPlataforma:= nil;
  SetLength(MatrizPlataforma, 7);
  for i := 0 to High(MatrizPlataforma) do
    SetLength(MatrizPlataforma[i], NumRegistros);
  //==============================================
  FrmDataModule.ADOQueryGerenciarSolicitacoes.First;
  while not FrmDataModule.ADOQueryGerenciarSolicitacoes.Eof do
  begin
    NumLinhas:= RLDestinoOrigem.RowCount;
    RLDestinoOrigem.RowCount:= NumLinhas+1;
    //Incluir Executante que sobe
    Destino:= (FrmDataModule.DataSourceGerenciarSolicitacoes.DataSet.
    FieldByName('txtDestino').AsString);
    RLDestinoOrigem.Cells[0,NumLinhas]:= Destino;
    RLDestinoOrigem.Cells[1,NumLinhas]:='';
    RLDestinoOrigem.Cells[2,NumLinhas]:='';
    RLDestinoOrigem.Cells[3,NumLinhas]:='';
    RLDestinoOrigem.Cells[4,NumLinhas]:='';
    RLDestinoOrigem.Cells[5,NumLinhas]:='';
    RLDestinoOrigem.Cells[6,NumLinhas]:='';
    RLDestinoOrigem.Cells[7,NumLinhas]:='';
    {RLDestinoOrigem.Cells[8,NumLinhas]:='';
    RLDestinoOrigem.Cells[9,NumLinhas]:='';
    RLDestinoOrigem.Cells[10,NumLinhas]:='';}
    RLDestinoOrigem.Cells[8,NumLinhas]:='';
    RLDestinoOrigem.Cells[9,NumLinhas]:='';
    RLDestinoOrigem.Cells[10,NumLinhas]:='';
    //============================================
    MatrizPlataforma[0,NumLinhas-1]:= Destino;
    MatrizPlataforma[1,NumLinhas-1]:= SalvatagemPlataforma(Destino);
    MatrizPlataforma[2,NumLinhas-1]:= FrmPrincipal.ForaOperacao(Destino,1);//GD
    MatrizPlataforma[3,NumLinhas-1]:= FrmPrincipal.ForaOperacao(Destino,4);//BCI
    MatrizPlataforma[4,NumLinhas-1]:= FrmPrincipal.ForaOperacao(Destino,11);//Surfer
    MatrizPlataforma[5,NumLinhas-1]:= FrmPrincipal.ForaOperacao(Destino,12);//SOV
    MatrizPlataforma[6,NumLinhas-1]:= FrmPrincipal.ForaOperacao(Destino,13);//Áqua
    //============================================
    FrmDataModule.ADOQueryGerenciarSolicitacoes.Next;
  end;
  try
    RLDestinoOrigem.FixedRows:= 1;
  except
    RLDestinoOrigem.RowCount:= 2;
    RLDestinoOrigem.FixedRows:= 1;
    RLDestinoOrigem.Rows[1].Clear;
  end;
  AutoFitGrade(RLDestinoOrigem);
  RLDestinoOrigem.Row:= 1;
  Destino:= (RLDestinoOrigem.Cells[0,1]);
  FrmPrincipal.buscaFiledGrid1('txtDestino',Destino,'Contem',ColunasLayoutExecutantes,4,false);
  actProcurar.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actMostrarExecute(Sender: TObject);
begin
  btnMostrarOcultar.Action:= actOcultar;
  PanelServicos.Height:= 25;
  DBGridServicos.Visible:= false;
end;

procedure TFrmGerenciarSolicitacoes.actMudancaExecute(Sender: TObject);
begin
  if not Assigned(FrmMotivoCancelamento) then
    FrmMotivoCancelamento:= TFrmMotivoCancelamento.Create(Application)
  else
    FrmMotivoCancelamento.Show;

  FrmMotivoCancelamento.Iniciar(2);
end;

procedure TFrmGerenciarSolicitacoes.actOrigemExecute(Sender: TObject);
  var
    DataProcura,Origem: String;
    numLinhas,NumRegistros,i: Integer;
begin
  DataProcura:= DateToStr(DateProgramacao.Date);
  //============================================================
  FrmDataModule.ADOQueryGerenciarSolicitacoes.Close;
  FrmDataModule.ADOQueryGerenciarSolicitacoes.SQL.Clear;
  FrmDataModule.ADOQueryGerenciarSolicitacoes.SQL.Add(
  'SELECT tblProgramacaoExecutante.Origem, tblProgramacaoDiaria.DataProgramacao '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'GROUP BY tblProgramacaoExecutante.Origem, tblProgramacaoDiaria.DataProgramacao '+
  'HAVING (tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(DataProcura)+');');
  FrmDataModule.ADOQueryGerenciarSolicitacoes.Open;
  //=========================================================
  RLDestinoOrigem.ColCount:= 6;
  RLDestinoOrigem.FixedRows:= 0;
  RLDestinoOrigem.RowCount:= 1;
  RLDestinoOrigem.Cells[0,0]:= 'Origem';
  RLDestinoOrigem.Cells[1,0]:= '∑ Apro.';
  RLDestinoOrigem.Cells[2,0]:= '∑ Canc.';
  RLDestinoOrigem.Cells[3,0]:= '∑ Mud.';
  RLDestinoOrigem.Cells[4,0]:= 'TS';
  RLDestinoOrigem.Cells[5,0]:= 'OP';
  //Configurar matriz de dados das plataformas
  NumRegistros:= FrmDataModule.ADOQueryGerenciarSolicitacoes.RecordCount;
  MatrizPlataforma:= nil;
  SetLength(MatrizPlataforma, 4);
  for i := 0 to High(MatrizPlataforma) do
    SetLength(MatrizPlataforma[i], NumRegistros);

  FrmDataModule.ADOQueryGerenciarSolicitacoes.First;
  i:= 0;
  while not FrmDataModule.ADOQueryGerenciarSolicitacoes.Eof do
  begin
    NumLinhas:= RLDestinoOrigem.RowCount;
    RLDestinoOrigem.RowCount:= NumLinhas+1;
    //Incluir Executante que sobe
    Origem:= (FrmDataModule.DataSourceGerenciarSolicitacoes.DataSet.
    FieldByName('Origem').AsString);
    RLDestinoOrigem.Cells[0,NumLinhas]:= Origem;
    RLDestinoOrigem.Cells[1,NumLinhas]:='';
    RLDestinoOrigem.Cells[2,NumLinhas]:='';
    RLDestinoOrigem.Cells[3,NumLinhas]:='';
    RLDestinoOrigem.Cells[4,NumLinhas]:='';
    RLDestinoOrigem.Cells[5,NumLinhas]:='';
    //===================================
    RLDestinoOrigem.Cells[6,NumLinhas]:='';
    RLDestinoOrigem.Cells[7,NumLinhas]:='';
    //===================================
    {RLDestinoOrigem.Cells[8,NumLinhas]:='';
    RLDestinoOrigem.Cells[9,NumLinhas]:='';
    RLDestinoOrigem.Cells[10,NumLinhas]:='';}
    //===================================
    MatrizPlataforma[0,i]:= Origem;
    MatrizPlataforma[1,i]:= SalvatagemPlataforma(Origem);
    //===================================
    FrmDataModule.ADOQueryGerenciarSolicitacoes.Next;
    i:=i+1;
  end;
  try
    RLDestinoOrigem.FixedRows:= 1;
  except
    RLDestinoOrigem.RowCount:= 2;
    RLDestinoOrigem.FixedRows:= 1;
    RLDestinoOrigem.Rows[1].Clear;
  end;
  AutoFitGrade(RLDestinoOrigem);
  RLDestinoOrigem.Row:= 1;
  Origem:= (RLDestinoOrigem.Cells[0,1]);
  FrmPrincipal.buscaFiledGrid1('Origem',Origem,'Contem',ColunasLayoutExecutantes,4,false);
  actProcurar.Execute;
end;

procedure TFrmGerenciarSolicitacoes.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase,DataProcura: String;
begin
  if not FrmDataModule.ADOQueryGerenciarSolicitacoes.IsEmpty then
  begin
    DataProcura:= DateToStr(DateProgramacao.Date);
    SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutExecutantes,false);
    if SQLString <> '' then
      SQLString:= ' AND '+SQLString;
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+')'+
    SQLString+' ORDER BY txtTipoEtapaServico,NomeExecutante;';
    FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryGerenciarExecutante,StatusBarExecutantes);
  end;
end;

procedure TFrmGerenciarSolicitacoes.actSelecaoLimparExecute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryGerenciarExecutante.RecNo;
  FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
  FrmDataModule.ADOQueryGerenciarExecutante.First;
  while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryGerenciarExecutante.Next;
  end;
  FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
  FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
  FrmDataModule.ADOQueryGerenciarExecutante.RecNo:= selRegistro;
  FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
end;

procedure TFrmGerenciarSolicitacoes.actSelecaoTodosExecute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryGerenciarExecutante.RecNo;
  FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
  FrmDataModule.ADOQueryGerenciarExecutante.First;
  while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryGerenciarExecutante.Next;
  end;
  FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
  FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
  FrmDataModule.ADOQueryGerenciarExecutante.RecNo:= selRegistro;
  FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
end;

procedure TFrmGerenciarSolicitacoes.actStatusTodosExecute(Sender: TObject);
  var
    i,linhaSel: Integer;
begin
  linhaSel:= RLDestinoOrigem.Row;
  actMatrizOrigemDestino.Execute;
  try
    abrirDestinoOrigem:= false;
    //FrmPrincipal.ProgressBarIncializa(RLDestinoOrigem.RowCount-1,
    //'Carregando Status da programação de executantes...');
    for i := 1 to RLDestinoOrigem.RowCount-1 do
    begin
      RLDestinoOrigem.Row:= i;
      StatusLinhaSelecionada;
      //FrmPrincipal.ProgressBarIncremento(1);
    end;
    RLDestinoOrigem.Row:= 1;
    //FrmPrincipal.ProgressBarAtualizar;
  finally
    RLDestinoOrigem.Row:= linhaSel;
    abrirDestinoOrigem:= true;
    actContadorSolicitacao.Execute;
  end;
end;

procedure TFrmGerenciarSolicitacoes.actZoomMaisExecute(Sender: TObject);
begin
  DBGridExecutantes.Font.Size:= DBGridExecutantes.Font.Size+1;
  DBGridServicos.Font.Size:= DBGridServicos.Font.Size+1;
  //RLDestinoOrigem
  RLDestinoOrigem.Font.Size:= RLDestinoOrigem.Font.Size+1;
  RLDestinoOrigem.DefaultRowHeight:= RLDestinoOrigem.DefaultRowHeight+1;
  AutoFitGrade(RLDestinoOrigem);
  //Contador Solicitações
  RLContadorSolicitacao.Font.Size:= RLContadorSolicitacao.Font.Size+1;
  RLContadorSolicitacao.DefaultRowHeight:= RLContadorSolicitacao.DefaultRowHeight+1;
  AutoFitGrade(RLContadorSolicitacao);
  //StatusBarExecutantes
  StatusBarExecutantes.Height:= StatusBarExecutantes.Height+2;
  StatusBarExecutantes.Font.Size:= StatusBarExecutantes.Font.Size+1;
  //Notas
  DBMemoNotas.Font.Size:= DBMemoNotas.Font.Size+1;
  //AutoFit
  actAutoFit.Execute;
  DBGridExecutantes.Columns[0].Width:= 40;
  DBGridExecutantes.Columns[1].Width:= 40;
  DBGridExecutantes.Columns[2].Width:= 40;
end;

procedure TFrmGerenciarSolicitacoes.actZoomMenosExecute(Sender: TObject);
begin
  if DBGridExecutantes.Font.Size > 7 then
  begin
    DBGridExecutantes.Font.Size:= DBGridExecutantes.Font.Size-1;
    DBGridServicos.Font.Size:= DBGridServicos.Font.Size-1;
    //RLDestinoOrigem
    RLDestinoOrigem.Font.Size:= RLDestinoOrigem.Font.Size-1;
    RLDestinoOrigem.DefaultRowHeight:= RLDestinoOrigem.DefaultRowHeight-1;
    AutoFitGrade(RLDestinoOrigem);
    //Contador Solicitações
    RLContadorSolicitacao.Font.Size:= RLContadorSolicitacao.Font.Size-1;
    RLContadorSolicitacao.DefaultRowHeight:= RLContadorSolicitacao.DefaultRowHeight-1;
    AutoFitGrade(RLContadorSolicitacao);
    //StatusBarExecutantes
    StatusBarExecutantes.Font.Size:= StatusBarExecutantes.Font.Size-1;
    StatusBarExecutantes.Height:= StatusBarExecutantes.Height-2;
    //Notas
    DBMemoNotas.Font.Size:= DBMemoNotas.Font.Size-1;
    //AutoFit
    actAutoFit.Execute;
    DBGridExecutantes.Columns[0].Width:= 40;
    DBGridExecutantes.Columns[1].Width:= 40;
    DBGridExecutantes.Columns[2].Width:= 40;
  end;
end;

procedure TFrmGerenciarSolicitacoes.btnDestinoClick(Sender: TObject);
begin
  if btnDestino.Down then
  begin
    btnOrigem.Down := false;
  end
  else
  begin
    btnOrigem.Down := true;
  end;
  actDestinoOrigem.Execute;
end;

procedure TFrmGerenciarSolicitacoes.btnOrigemClick(Sender: TObject);
begin
  if btnOrigem.Down then
  begin
    btnDestino.Down := false;
  end
  else
  begin
    btnDestino.Down := true;
  end;
  actDestinoOrigem.Execute;
end;

procedure TFrmGerenciarSolicitacoes.ComboBoxDestinoCloseUp(Sender: TObject);
begin
  RLAlterarDestinos.Cells[1,RLAlterarDestinos.Row]:= (ComboBoxDestino.Text);
  ComboBoxDestino.Visible := False;
end;

procedure TFrmGerenciarSolicitacoes.ComboBoxDestinoMouseLeave(Sender: TObject);
begin
  ComboBoxDestino.Visible := False;
end;

procedure TFrmGerenciarSolicitacoes.ComboBoxExecutanteKeyPress(Sender: TObject;
  var Key: Char);
begin
  key:= #0;
end;

function TFrmGerenciarSolicitacoes.Contar_KIT_PS(Destino: String): Integer;
  var
    i,Contar: Integer;
begin
  Contar:= 0;
  for i := 0 to High(MatrizConsulta[0]) do
    if ((MatrizConsulta[1,i]=Destino)AND
    (uppercase(MatrizConsulta[4,i])= 'TRUE')AND
    (MatrizConsulta[2,i]='Aprovado')) then
      Contar:= Contar+1;
  //==================================
  Result:= Contar;

  {
  Contar:= 0;
  for i := 0 to High(MatrizKIT_PS[0]) do
  begin
    if MatrizKIT_PS[1,i] = Destino then
      Contar:= Contar+1;
  end;
  Result:= Contar;}
end;

function TFrmGerenciarSolicitacoes.Contar_TS_OP(TipoEtapaServico,
  Destino: String): Integer;
  var
    i,Contar: Integer;
begin
  Contar:= 0;
  for i := 0 to High(MatrizConsulta[0]) do
    if ((MatrizConsulta[1,i]=Destino)AND
    (MatrizConsulta[3,i]=TipoEtapaServico)AND
    (MatrizConsulta[2,i]='Aprovado')) then
      Contar:= Contar+1;
  //==================================
  Result:= Contar;
end;

procedure TFrmGerenciarSolicitacoes.DateProgramacaoCloseUp(Sender: TObject);
begin
  if DateProgramacao.DateTime < FrmPrincipal.carregaDataMinima(false) then
  begin
    btnAprovar.Enabled:= false;
    btnCancelar.Enabled:= false;
  end
  else
  begin
    btnAprovar.Enabled:= true;
    btnCancelar.Enabled:= true;
  end;
  actDestinoOrigem.Execute;
end;

procedure TFrmGerenciarSolicitacoes.DBGridContadorCellClick(Column: TColumn);
begin
  try
    if (Self.DBGridContador.SelectedField.DataType = ftBoolean)AND
    (Column.Field.FieldName = 'BooleanOrigemDestino') then
    begin
      DBGridContador.Options:=
      [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
      FrmDataModule.ADOQueryContadorSolicitacao.Edit;
      FrmDataModule.DataSourceContadorSolicitacao.DataSet.
      FieldByName('BooleanOrigemDestino').AsBoolean:= not Self.DBGridContador.SelectedField.AsBoolean;
      try
        FrmDataModule.ADOQueryContadorSolicitacao.Post;
      except
        MessageBox(0,'Banco de dados alterado, necessário atualização!','Banco de Dados',MB_ICONINFORMATION);
        FrmDataModule.ADOQueryContadorSolicitacao.Cancel;
        FrmDataModule.ADOQueryContadorSolicitacao.Refresh;
      end;
    end
    else
      DBGridContador.Options:=
      [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
  except
  end;
end;

procedure TFrmGerenciarSolicitacoes.DBGridContadorDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridContador,ColunasLayoutContador,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'BooleanOrigemDestino') then
  begin
    Self.DBGridContador.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 10;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridContador.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end
end;

procedure TFrmGerenciarSolicitacoes.DBGridExecutantesCellClick(Column: TColumn);
begin
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)) then
  begin
    try
      if (Self.DBGridExecutantes.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'booleanSelecao') then
      begin
        DBGridExecutantes.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOqueryGerenciarExecutante.Edit;
        FrmDataModule.DataSourceGerenciarExecutante.DataSet.
        FieldByName('booleanSelecao').AsBoolean:= not Self.DBGridExecutantes.SelectedField.AsBoolean;
        try
          FrmDataModule.ADOqueryGerenciarExecutante.Post;
        except
          MessageBox(0,'Banco de dados alterado, necessário atualização!','Banco de Dados',MB_ICONINFORMATION);
          FrmDataModule.ADOqueryGerenciarExecutante.Cancel;
          actDestinoOrigem.Execute;
        end;
      end
      else if (Self.DBGridExecutantes.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'Kit_PS') then
      begin
        DBGridExecutantes.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOqueryGerenciarExecutante.Edit;
        FrmDataModule.DataSourceGerenciarExecutante.DataSet.
        FieldByName('Kit_PS').AsBoolean:= not Self.DBGridExecutantes.SelectedField.AsBoolean;
        try
          FrmDataModule.ADOqueryGerenciarExecutante.Post;
        except
          MessageBox(0,'Banco de dados alterado, necessário atualização!','Banco de Dados',MB_ICONINFORMATION);
          FrmDataModule.ADOqueryGerenciarExecutante.Cancel;
          actDestinoOrigem.Execute;
        end;
      end
      else
        DBGridExecutantes.Options:=
        [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
    except
    end;
  end;
end;

procedure TFrmGerenciarSolicitacoes.DBGridExecutantesDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
begin
  FrmPrincipal.GridZebrado(DBGridExecutantes,ColunasLayoutExecutantes,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'StatusProgramacao')OR
  (Column.Field.FieldName = 'NomeExecutante')OR
  (Column.Field.FieldName = 'txtTipoEtapaServico') then
  begin
    if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Aprovado' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clLime;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Cancelado' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clRed;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Mudança' then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clYellow;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumCancelados') then
  begin
    if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('NumCancelados').AsInteger > 0 then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clRed;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumAprovados') then
  begin
    if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
    FieldByName('NumAprovados').AsInteger > 0 then
    begin
      DBGridExecutantes.Canvas.Brush.Color:= clLime;
      DBGridExecutantes.Font.Color:= clBlack;
      DBGridExecutantes.Canvas.FillRect(Rect);
      DBGridExecutantes.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmGerenciarSolicitacoes.DBGridExecutantesKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridExecutantes.DataSource.DataSet.RecNo=
  DBGridExecutantes.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmGerenciarSolicitacoes.DBGridExecutantesKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key:= #0;
end;

procedure TFrmGerenciarSolicitacoes.DBGridServicosDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridServicos,ColunasLayoutServico,State,Rect,DataCol,Column);
end;

procedure TFrmGerenciarSolicitacoes.DBGridServicosKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridServicos.DataSource.DataSet.RecNo=
  DBGridServicos.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmGerenciarSolicitacoes.DBGridServicosKeyPress(Sender: TObject;
  var Key: Char);
begin
  key:= #0;
end;

procedure TFrmGerenciarSolicitacoes.edtDestinoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    actProcurar.Execute;
end;

procedure TFrmGerenciarSolicitacoes.edtNomeExecutanteKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    actProcurar.Execute;
end;

procedure TFrmGerenciarSolicitacoes.edtTipoEtapaServicoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    actProcurar.Execute;
end;

procedure TFrmGerenciarSolicitacoes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmGerenciarSolicitacoes:=nil;
end;

procedure TFrmGerenciarSolicitacoes.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  FrmDataModule.setFilterDBGrid(DBGridExecutantes);
  FrmDataModule.setFilterDBGrid(DBGridServicos);
  //=====================================
  actTransbordos.ImageIndex:= 456;
  actTransbordos.Caption:= 'Transbordo';
  btnKITPS.Visible:= true;
  DateProgramacao.Date:= IncDay(now,1);
  FrmPrincipal.actMatrizForaOperacao.Execute;
  //=============================================
  FrmPrincipal.ProgressBarIncializa(7,'Inicializando....');
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)) then
  begin
    btnAprovar.Enabled:= true;
    btnCancelar.Enabled:= true;
    btnMudanca.Enabled:= true;
    actCarregarDestino.Enabled:= true;
    DBGridExecutantes.ReadOnly:= false;
    DBGridServicos.ReadOnly:= false;
    actCarregarDestinoJanela.Enabled:= true;
    actCopiarKITPS.Enabled:= true;
    actConfigurarContador.Enabled:= true;
  end
  else
  begin
    btnAprovar.Enabled:= false;
    btnCancelar.Enabled:= false;
    btnMudanca.Enabled:= false;
    actCarregarDestino.Enabled:= false;
    DBGridExecutantes.ReadOnly:= true;
    DBGridServicos.ReadOnly:= true;
    actCarregarDestinoJanela.Enabled:= false;
    actCopiarKITPS.Enabled:= false;
    actConfigurarContador.Enabled:= false;
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  //Incicialização
  FrmPrincipal.ProgressBarIncremento(1);
  abrirDestinoOrigem:= true;
  actDestinoOrigem.Execute;
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.SetupGridFilterPickListSQL(FrmDataModule.ADOConnectionConsulta,'txtDestino',
  'SELECT Plataforma FROM tblPlataforma GROUP BY Plataforma;',DBGridExecutantes,false);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.SetupGridFilterPickListSQL(FrmDataModule.ADOConnectionConsulta,'Origem',
  'SELECT Plataforma FROM tblPlataforma GROUP BY Plataforma;',DBGridExecutantes,false);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmDataModule.ADOQueryInserirServico.Active:= false;
  FrmDataModule.ADOQueryInserirServico.Active:= true;
  FrmPrincipal.ProgressBarIncremento(1);
  FrmDataModule.ADOQueryInserirExecutante1.Active:= false;
  FrmDataModule.ADOQueryInserirExecutante1.Active:= true;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmGerenciarSolicitacoes.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmGerenciarSolicitacoes.GravarStatus(StatusProgramacao: String);
  var
    idProgramacaoExecutante,CodigoProgramacaoDiaria,NumRegistros: Integer;
begin
  if not FrmDataModule.ADOQueryGerenciarExecutante.IsEmpty then
  begin
    FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
    NumRegistros:= FrmDataModule.ADOQueryGerenciarExecutante.RecordCount;
    FrmDataModule.ADOQueryGerenciarExecutante.First;
    FrmPrincipal.ProgressBarIncializa(NumRegistros,
    'Atribuindo status para programação...');
    while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
    begin
      if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
      FieldByName('booleanSelecao').AsBoolean then
      begin
        idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
        DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
        FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,0,
        StatusProgramacao,'');
      end;
      FrmDataModule.ADOQueryGerenciarExecutante.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
    FrmPrincipal.ProgressBarAtualizar;
    //Gravar N° de Cancelados e Aprovados
    FrmDataModule.ADOQueryGerenciarExecutante.First;
    FrmPrincipal.ProgressBarIncializa(NumRegistros,
    'Calculando: Aprovados, Cancelados e N° de Executantes...');
    while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
    begin
      if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
      FieldByName('booleanSelecao').AsBoolean then
      begin
        CodigoProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.
        DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger;
        FrmPrincipal.GravarCanceladoAprovado(CodigoProgramacaoDiaria);
      end;
      FrmDataModule.ADOQueryGerenciarExecutante.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
    FrmPrincipal.ProgressBarAtualizar;
    FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
    FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
    FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
    //Carregar Numero de Aprovados,Cancelados
    actMatrizOrigemDestino.Execute;
    StatusLinhaSelecionada;
  end;
end;

procedure TFrmGerenciarSolicitacoes.PanelTituloAjuda1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmGerenciarSolicitacoes.PanelTituloAjuda1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmGerenciarSolicitacoes.PanelTituloAjuda1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    FrmPrincipal.Capturing := false;
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmGerenciarSolicitacoes.PanelTituloAjudaMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmGerenciarSolicitacoes.PanelTituloMagicMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmGerenciarSolicitacoes.RLAlterarDestinosSelectCell(Sender: TObject;
  ACol, ARow: Integer; var CanSelect: Boolean);
  var
    R: TRect;
begin
  if ((ACol = 1) AND(ARow > 0)) then
  begin
    R := RLAlterarDestinos.CellRect(ACol, ARow);
    R.Left := R.Left + RLAlterarDestinos.Left;
    R.Right := R.Right + RLAlterarDestinos.Left;
    R.Top := R.Top + RLAlterarDestinos.Top;
    R.Bottom := R.Bottom + RLAlterarDestinos.Top;
    ComboBoxDestino.Left := R.Left + 1;
    ComboBoxDestino.Top := R.Top + 1;
    ComboBoxDestino.Width := (R.Right + 1) - R.Left;
    ComboBoxDestino.Height := (R.Bottom + 1) - R.Top;
    ComboBoxDestino.Visible := True;
    FrmPrincipal.selComboBox(ComboBoxDestino,(RLAlterarDestinos.Cells[1, ARow]));
    ComboBoxDestino.SetFocus;
  end;
  CanSelect := True;
end;

procedure TFrmGerenciarSolicitacoes.RLDestinoOrigemDrawCell(Sender: TObject;
  ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  R: TRect;
  KIT_PS: Integer;
begin
  //GD, BCI ou Balsa
  if ((ACol=6)OR(ACol=7)OR(ACol=8)OR(ACol=9)OR(ACol=10)) then
  begin
    R := Rect;
    R.Left := R.Left - 4;
    //Fora Operação
    if RLDestinoOrigem.Cells[ACol,ARow] = 'FO' then
    begin
      RLDestinoOrigem.Canvas.Brush.Color := clRed;
      RLDestinoOrigem.Canvas.Font.Style := [];
      RLDestinoOrigem.Canvas.FillRect(R);
      DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
      -1, R, DT_Center);
      //FrmPrincipal.ImageList1.Draw(RLDestinoOrigem.Canvas,Rect.Left+4,Rect.Top+4,93);
    end
    //OK
    else if RLDestinoOrigem.Cells[ACol,ARow] = 'OK' then
    begin
      RLDestinoOrigem.Canvas.Brush.Color := clLime;
      RLDestinoOrigem.Canvas.Font.Style := [];
      RLDestinoOrigem.Canvas.FillRect(R);
      DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
      -1, R, DT_Center);
      //FrmPrincipal.ImageList1.Draw(RLDestinoOrigem.Canvas,Rect.Left+4,Rect.Top+4,92);
    end
    //Restrição Operacional
    else if RLDestinoOrigem.Cells[ACol,ARow] = 'RO' then
    begin
      RLDestinoOrigem.Canvas.Brush.Color := clYellow;
      RLDestinoOrigem.Canvas.Font.Style := [];
      RLDestinoOrigem.Canvas.FillRect(R);
      DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
      -1, R, DT_Center);
      //FrmPrincipal.ImageList1.Draw(RLDestinoOrigem.Canvas,Rect.Left+4,Rect.Top+4,454);
    end
    //N/A
    else if RLDestinoOrigem.Cells[ACol,ARow] = 'N/A' then
    begin
      RLDestinoOrigem.Canvas.Brush.Color := clSilver;
      RLDestinoOrigem.Canvas.Font.Style := [];
      RLDestinoOrigem.Canvas.FillRect(R);
      DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
      -1, R, DT_Center);
      //FrmPrincipal.ImageList1.Draw(RLDestinoOrigem.Canvas,Rect.Left+4,Rect.Top+4,454);
    end;
  end;
  if (ARow > 0) then
  begin
    R := Rect;
    R.Left := R.Left - 4;
    try
      //N° Aprovado
      if (StrToInt(RLDestinoOrigem.Cells[ACol,ARow]) > 0)AND(ACol=1) then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clLime;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      //N° Cancelado
      else if (StrToInt(RLDestinoOrigem.Cells[ACol,ARow]) > 0)AND(ACol=2) then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clRed;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      //N° Mudança
      else if (StrToInt(RLDestinoOrigem.Cells[ACol,ARow]) > 0)AND(ACol=3) then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clYellow;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      //TS
      else if (StrToInt(RLDestinoOrigem.Cells[ACol,ARow]) > 0)AND(ACol=4) then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clAqua;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      //OP
      else if (StrToInt(RLDestinoOrigem.Cells[ACol,ARow]) > 0)AND(ACol=5) then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clMenuHighlight;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      //KIT_PS
      else if (ACol=13) then
      begin
        try
          KIT_PS:= StrToInt(RLDestinoOrigem.Cells[ACol,ARow]);
          if (KIT_PS = 1) then
          begin
            RLDestinoOrigem.Canvas.Brush.Color := clLime;
            RLDestinoOrigem.Canvas.Font.Style := [];
            RLDestinoOrigem.Canvas.FillRect(R);
            DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
            -1, R, DT_Center);
          end
          else
          begin
            RLDestinoOrigem.Canvas.Brush.Color := clRed;
            RLDestinoOrigem.Canvas.Font.Style := [];
            RLDestinoOrigem.Canvas.FillRect(R);
            DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
            -1, R, DT_Center);
          end;
        except
        end;
      end
      else if gdSelected in State then
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clSkyBlue;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end
      else
      begin
        RLDestinoOrigem.Canvas.Brush.Color := clWhite;
        RLDestinoOrigem.Canvas.Font.Style := [];
        RLDestinoOrigem.Canvas.FillRect(R);
        DrawText(RLDestinoOrigem.Canvas.Handle, PChar(RLDestinoOrigem.Cells[ACol, ARow]),
        -1, R, DT_Center);
      end;
    except
    end;
  end;
end;

procedure TFrmGerenciarSolicitacoes.RLDestinoOrigemSelectCell(Sender: TObject;
  ACol, ARow: Integer; var CanSelect: Boolean);
  var
    Origem,Destino: String;
begin
  if abrirDestinoOrigem then
  begin
    try
      if btnDestino.Down then
      begin
        Destino:= (RLDestinoOrigem.Cells[0, ARow]);
        FrmPrincipal.buscaFiledGrid1('txtDestino',Destino,'Contem',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('Origem','','',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('CodigoSAP','','',ColunasLayoutExecutantes,4,false);
        actProcurar.Execute;
      end
      else if btnOrigem.Down then
      begin
        Origem:= (RLDestinoOrigem.Cells[0, ARow]);
        FrmPrincipal.buscaFiledGrid1('txtDestino','','',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('Origem',Origem,'Contem',ColunasLayoutExecutantes,4,false);
        FrmPrincipal.buscaFiledGrid1('CodigoSAP','','',ColunasLayoutExecutantes,4,false);
        actProcurar.Execute;
      end;
    except

    end;
  end;
end;

procedure TFrmGerenciarSolicitacoes.SpeedButton4Click(Sender: TObject);
begin
  PanelAjuda.Visible:= false;
end;

procedure TFrmGerenciarSolicitacoes.CopiarGridComoImagemParaClipboard(
  AGrid: TStringGrid);
var
  Bitmap: TBitmap;
begin
  if (AGrid = nil) or (not AGrid.HandleAllocated) then
    Exit;

  Bitmap := TBitmap.Create;
  try
    AGrid.Update;

    Bitmap.PixelFormat := pf24bit;
    Bitmap.SetSize(AGrid.Width, AGrid.Height);
    Bitmap.Canvas.Brush.Color := clWhite;
    Bitmap.Canvas.FillRect(Rect(0, 0, Bitmap.Width, Bitmap.Height));
    TWinControlAccess(AGrid).PaintTo(Bitmap.Canvas.Handle, 0, 0);

    Clipboard.Assign(Bitmap);
  finally
    Bitmap.Free;
  end;
end;

procedure TFrmGerenciarSolicitacoes.StatusLinhaSelecionada;
  var
    DestinoOrigem: String;
    NumContar,NumAprovados,Saldo,QtdeSAI,POB: Integer;
    DadosPlataforma: TPlataforma;
begin
  DestinoOrigem:= (RLDestinoOrigem.Cells[0, RLDestinoOrigem.Row]);
  if btnDestino.Down then
  begin
    //NumAprovados
    NumAprovados:= TotalDestino(DestinoOrigem,'Aprovado');
    RLDestinoOrigem.Cells[1,RLDestinoOrigem.Row]:= IntToStr(NumAprovados);
    //NumCancelados
    NumContar:= TotalDestino(DestinoOrigem,'Cancelado');
    RLDestinoOrigem.Cells[2,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    //NumMudancas
    NumContar:= TotalDestino(DestinoOrigem,'Mudança');
    RLDestinoOrigem.Cells[3,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    //Saldo
    QtdeSAI:= TotalOrigem(DestinoOrigem,'Aprovado');
    POB:= TotalPOB(DestinoOrigem,'Aprovado');
    Saldo:= NumAprovados-QtdeSAI+POB;
    //TS e OP
    NumContar:= Contar_TS_OP('SEGURANÇA INDUSTRIAL',DestinoOrigem);
    RLDestinoOrigem.Cells[4,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    NumContar:= Contar_TS_OP('OPERAÇÃO',DestinoOrigem);
    RLDestinoOrigem.Cells[5,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    //Dados da Plataforma
    DadosPlataforma:= InfoPlataforma(DestinoOrigem);
    RLDestinoOrigem.Cells[6,RLDestinoOrigem.Row]:= DadosPlataforma.GD;
    RLDestinoOrigem.Cells[7,RLDestinoOrigem.Row]:= DadosPlataforma.BCI;
    {RLDestinoOrigem.Cells[8,RLDestinoOrigem.Row]:= DadosPlataforma.Surfer;
    RLDestinoOrigem.Cells[9,RLDestinoOrigem.Row]:= DadosPlataforma.SOV;
    RLDestinoOrigem.Cells[10,RLDestinoOrigem.Row]:= DadosPlataforma.Aqua;}
    RLDestinoOrigem.Cells[8,RLDestinoOrigem.Row]:= IntToStr(Saldo);
    RLDestinoOrigem.Cells[9,RLDestinoOrigem.Row]:= DadosPlataforma.Salvatagem;
    //KIT_PS
    NumContar:= Contar_KIT_PS(DestinoOrigem);
    RLDestinoOrigem.Cells[10,RLDestinoOrigem.Row]:= IntToStr(NumContar);
  end
  else if btnOrigem.Down then
  begin
    //NumAprovados
    NumAprovados:= TotalOrigem(DestinoOrigem,'Aprovado');
    RLDestinoOrigem.Cells[1,RLDestinoOrigem.Row]:= IntToStr(NumAprovados);
    //NumCancelados
    NumContar:= TotalOrigem(DestinoOrigem,'Cancelado');
    RLDestinoOrigem.Cells[2,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    //NumMudancas
    NumContar:= TotalOrigem(DestinoOrigem,'Mudança');
    RLDestinoOrigem.Cells[3,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    //TS e OP
    NumContar:= Contar_TS_OP('SEGURANÇA INDUSTRIAL',DestinoOrigem);
    RLDestinoOrigem.Cells[4,RLDestinoOrigem.Row]:= IntToStr(NumContar);
    NumContar:= Contar_TS_OP('OPERAÇÃO',DestinoOrigem);
    RLDestinoOrigem.Cells[5,RLDestinoOrigem.Row]:= IntToStr(NumContar);
  end;
end;

function TFrmGerenciarSolicitacoes.TotalDestino(Destino,
  StatusProgramacao: String): Integer;
  var
    i,Contar: Integer;
begin
  Contar:= 0;
  for i := 0 to High(MatrizConsulta[0]) do
    if ((MatrizConsulta[1,i]=Destino)AND(MatrizConsulta[0,i]<>Destino)AND
    (MatrizConsulta[2,i]=StatusProgramacao)AND
    (MatrizConsulta[0,i]<>'HOTEL')) then
      Contar:= Contar+1;
  //==================================
  Result:= Contar;
end;

function TFrmGerenciarSolicitacoes.TotalOrigem(Origem,
  StatusProgramacao: String): Integer;
  var
    i,Contar: Integer;
begin
  Contar:= 0;
  for i := 0 to High(MatrizConsulta[0]) do
    if ((MatrizConsulta[0,i]=Origem)AND(MatrizConsulta[1,i]<>Origem)AND
    (MatrizConsulta[2,i]=StatusProgramacao)) then
      Contar:= Contar+1;
  //==================================
  Result:= Contar;
end;

function TFrmGerenciarSolicitacoes.TotalPOB(Destino,
  StatusProgramacao: String): Integer;
  var
    i,Contar: Integer;
begin
  Contar:= 0;
  for i := 0 to High(MatrizConsulta[0]) do
    if ((MatrizConsulta[1,i]=Destino)AND
    (MatrizConsulta[0,i]=MatrizConsulta[1,i])AND
    (MatrizConsulta[2,i]=StatusProgramacao)) then
      Contar:= Contar+1;
  //==================================
  Result:= Contar;
end;

function TFrmGerenciarSolicitacoes.SalvatagemPlataforma(Plataforma: String): String;
  var
    i: Integer;
begin
  for i := 0 to High(FrmPrincipal.MatrizSalvatagem[0]) do
  begin
    if FrmPrincipal.MatrizSalvatagem[0,i] = Plataforma  then
    begin
      Result:= FrmPrincipal.MatrizSalvatagem[1,i];
      break
    end;
  end;
end;

function TFrmGerenciarSolicitacoes.InfoPlataforma(Plataforma: String): TPlataforma;
  var
    i: Integer;
begin
  Result.Salvatagem:= '0';
  for i := 0 to High(MatrizPlataforma[0]) do
    if (MatrizPlataforma[0,i] = Plataforma) then
    begin
      Result.Salvatagem:= (MatrizPlataforma[1,i]);
      Result.GD:= (MatrizPlataforma[2,i]);
      Result.BCI:= (MatrizPlataforma[3,i]);
      Result.Surfer:= (MatrizPlataforma[4,i]);
      Result.SOV:= (MatrizPlataforma[5,i]);
      Result.Aqua:= (MatrizPlataforma[6,i]);
      break
    end;
end;

procedure TFrmGerenciarSolicitacoes.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
