unit untFrmManifestoEmbarque;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Data.DB, System.Actions,
  Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, Vcl.DBGrids,
  untDBGridFilter, Vcl.ComCtrls, Vcl.Buttons, Vcl.DBCtrls, Vcl.ToolWin,
  Data.Win.ADODB, Vcl.Grids, Vcl.StdCtrls, Vcl.CheckLst, uZucchi,
  System.DateUtils;

type
  TFrmManifestoEmbarque = class(TForm)
    PanelTitulo: TPanel;
    StatusBarManifestoEmbarque: TStatusBar;
    ToolBar1: TToolBar;
    DBNavigatorManifestoEmbarque: TDBNavigator;
    btnClearFiltro: TToolButton;
    btnLayout: TToolButton;
    btnExcel: TToolButton;
    DBGridManifestoEmbarque: TFilterDBGrid;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    actExcelImportar: TAction;
    actAtualizarExecutante: TAction;
    actCriarProgramacao: TAction;
    actExcluirFiltrados: TAction;
    BitBtn11: TBitBtn;
    BitBtn1: TBitBtn;
    BitBtnExcluir: TBitBtn;
    ToolButton1: TToolButton;
    Panel3: TPanel;
    dataInicio: TDateTimePicker;
    Panel4: TPanel;
    dataFim: TDateTimePicker;
    Panel8: TPanel;
    PanelFiltroRapido: TPanel;
    GroupBoxTipoEmbarque: TGroupBox;
    clbTipoEmbarque: TCheckListBox;
    GroupBoxSentido: TGroupBox;
    clbEmbarqueDesembarque: TCheckListBox;
    PanelFiltroRapidoAcoes: TPanel;
    btnRestaurarFiltrosRapidos: TButton;

    procedure actProcurarExecute(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure actExcelImportarExecute(Sender: TObject);
    procedure actAtualizarExecutanteExecute(Sender: TObject);
    procedure actCriarProgramacaoExecute(Sender: TObject);
    procedure actExcluirFiltradosExecute(Sender: TObject);
    procedure dataInicioCloseUp(Sender: TObject);
    procedure dataFimCloseUp(Sender: TObject);
    procedure dataFimKeyPress(Sender: TObject; var Key: Char);
    procedure dataInicioKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridManifestoEmbarqueDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure clbFiltroRapidoClickCheck(Sender: TObject);
    procedure btnRestaurarFiltrosRapidosClick(Sender: TObject);
  private
    { Private declarations }
    FAtualizandoFiltroRapido: Boolean;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure InicializarFiltrosRapidos;
    procedure MarcarTodosFiltroRapido;
    function SQLFiltroRapidoLista(ACheckList: TCheckListBox;
      const ACampo: string): string;
    function SQLFiltroRapido: string;
  public
    { Public declarations }
  end;

var
  FrmManifestoEmbarque: TFrmManifestoEmbarque;

implementation
  uses untPrincipal,untFrmLogin,untDataModule;


{$R *.dfm}

const
  VALOR_TIPO_EMBARQUE_FIXO = 'FIXO';
  VALOR_TIPO_EMBARQUE_SOV = 'SOV';
  VALOR_TIPO_EMBARQUE_BATE_VOLTA = 'BATE-VOLTA';
  VALOR_TIPO_EMBARQUE_AEREO = 'A' + #201 + 'REO';
  VALOR_EMBARQUE = 'EMBARQUE';
  VALOR_DESEMBARQUE = 'DESEMBARQUE';

procedure TFrmManifestoEmbarque.actExcelImportarExecute(Sender: TObject);
begin
  ImportarExcelSimplesPerguntaInsercao(DBGridManifestoEmbarque,
    FrmPrincipal.ProgressBarPrincipal,FrmPrincipal.MemoPrincipal);
end;

procedure TFrmManifestoEmbarque.actExcluirFiltradosExecute(Sender: TObject);
var
  DS: TDataSet;
begin
  DS := FrmDataModule.DataSourceManifestoEmbarque.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    MessageBox(0, 'Nenhum registro filtrado encontrado para excluir.', 'Colibri', MB_OK or MB_ICONINFORMATION);
    Exit;
  end;

  if DS.State in dsEditModes then
    DS.Post;

  if Application.MessageBox(
    'Deseja excluir todos os registros filtrados do manifesto?',
    '.::ATENCAO::.',
    MB_ICONQUESTION or MB_YESNO
  ) <> IDYES then
    Exit;

  deleteRegistros(FrmDataModule.DataSourceManifestoEmbarque,
    FrmPrincipal.ProgressBarPrincipal, FrmPrincipal.MemoPrincipal);

  if FrmDataModule.ADOQueryManifestoEmbarque.Active then
    FrmDataModule.ADOQueryManifestoEmbarque.Requery();
end;

procedure TFrmManifestoEmbarque.actAtualizarExecutanteExecute(Sender: TObject);
var
  Bmk: TBookmark;
  DS: TDataSet;
  HasBookmark: Boolean;
  QAtualizar: TADOQuery;
  CodigoSAP, CPFNumerico, CPFFormatado, OutroDocumento: string;
  AtualizadosCodigoSAP, AtualizadosCPF, AtualizadosOutroDocumento: Integer;
  Ignorados: Integer;

  function NormalizarCodigoSAPLocal(const ACodigo: string): string;
  var
    I: Integer;
    ApenasNumeros: Boolean;
  begin
    Result := Trim(ACodigo);
    if Result = '' then
      Exit;

    ApenasNumeros := True;
    for I := 1 to Length(Result) do
    begin
      if not CharInSet(Result[I], ['0'..'9']) then
      begin
        ApenasNumeros := False;
        Break;
      end;
    end;

    if ApenasNumeros then
    begin
      while (Length(Result) > 1) and (Result[1] = '0') do
        Delete(Result, 1, 1);
    end;

    if Result = '0' then
      Result := '';
  end;

  function CodigoSAPEhNumericoLocal(const ACodigo: string): Boolean;
  begin
    Result := NormalizarCodigoSAPLocal(ACodigo) <> '';
  end;

  function CPFValidoParaBusca(const ACPF: string; out ACPFNumerico,
    ACPFFormatado: string): Boolean;
  begin
    ACPFNumerico := FrmPrincipal.SomenteNumero(Trim(ACPF));
    ACPFFormatado := FrmPrincipal.VerificaCPF(ACPF);
    Result := ACPFFormatado <> '';
  end;

  procedure PrepararParametrosRateio;
  begin
    QAtualizar.Parameters.ParamByName('pCentroCusto').Value :=
      Trim(DS.FieldByName('CentroCusto').AsString);
    QAtualizar.Parameters.ParamByName('pDiagramaRede').Value :=
      Trim(DS.FieldByName('DiagramaRede').AsString);
    QAtualizar.Parameters.ParamByName('pOperRede').Value :=
      Trim(DS.FieldByName('OperRede').AsString);
    QAtualizar.Parameters.ParamByName('pElementoPEP').Value :=
      Trim(DS.FieldByName('ElementoPEP').AsString);
  end;

  function AtualizarPorCodigoSAP(const ACodigoSAP: string): Boolean;
  begin
    QAtualizar.Close;
    QAtualizar.SQL.Text :=
      'UPDATE tblExecutante ' +
      'SET CentroCusto = :pCentroCusto, ' +
      '    DiagramaRede = :pDiagramaRede, ' +
      '    OperRede = :pOperRede, ' +
      '    ElementoPEP = :pElementoPEP ' +
      'WHERE Trim(CodigoSAP) = :pCodigoSAP';
    PrepararParametrosRateio;
    QAtualizar.Parameters.ParamByName('pCodigoSAP').Value := ACodigoSAP;
    QAtualizar.ExecSQL;
    Result := QAtualizar.RowsAffected > 0;
  end;

  function AtualizarPorCPF(const ACPFNumerico, ACPFFormatado: string): Boolean;
  begin
    QAtualizar.Close;
    QAtualizar.SQL.Text :=
      'UPDATE tblExecutante ' +
      'SET CentroCusto = :pCentroCusto, ' +
      '    DiagramaRede = :pDiagramaRede, ' +
      '    OperRede = :pOperRede, ' +
      '    ElementoPEP = :pElementoPEP ' +
      'WHERE Trim(CPF) = :pCPF OR Trim(CPF) = :pCPF2';
    PrepararParametrosRateio;
    QAtualizar.Parameters.ParamByName('pCPF').Value := ACPFNumerico;
    QAtualizar.Parameters.ParamByName('pCPF2').Value := ACPFFormatado;
    QAtualizar.ExecSQL;
    Result := QAtualizar.RowsAffected > 0;
  end;

  function AtualizarPorOutroDocumento(const AOutroDocumento: string): Boolean;
  begin
    QAtualizar.Close;
    QAtualizar.SQL.Text :=
      'UPDATE tblExecutante ' +
      'SET CentroCusto = :pCentroCusto, ' +
      '    DiagramaRede = :pDiagramaRede, ' +
      '    OperRede = :pOperRede, ' +
      '    ElementoPEP = :pElementoPEP ' +
      'WHERE Trim(OutroDocumento) = :pOutroDocumento';
    PrepararParametrosRateio;
    QAtualizar.Parameters.ParamByName('pOutroDocumento').Value := Trim(AOutroDocumento);
    QAtualizar.ExecSQL;
    Result := QAtualizar.RowsAffected > 0;
  end;

begin
  DS := FrmDataModule.DataSourceManifestoEmbarque.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    MessageBox(0, 'Nenhum registro filtrado encontrado para atualizar.', 'Colibri', MB_OK or MB_ICONINFORMATION);
    Exit;
  end;

  if DS.State in dsEditModes then
    DS.Post;

  if Application.MessageBox(
    'Deseja atualizar a tblExecutante com os dados de CentroCusto, DiagramaRede, OperRede e ElementoPEP de todos os registros filtrados no manifesto?',
    '.::ATENCAO::.',
    MB_ICONQUESTION or MB_YESNO
  ) <> IDYES then
    Exit;

  HasBookmark := False;
  if not DS.IsEmpty then
  begin
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  QAtualizar := TADOQuery.Create(nil);
  try
    QAtualizar.Connection := FrmDataModule.ADOConnectionConsulta;
    QAtualizar.ParamCheck := True;

    AtualizadosCodigoSAP := 0;
    AtualizadosCPF := 0;
    AtualizadosOutroDocumento := 0;
    Ignorados := 0;

    FrmPrincipal.ProgressBarIncializa(
      DS.RecordCount,
      'Atualizando executantes com base nos registros filtrados do manifesto.'
    );

    FrmDataModule.ADOConnectionConsulta.BeginTrans;
    try
      DS.DisableControls;
      try
        DS.First;
        while not DS.Eof do
        begin
          CodigoSAP := NormalizarCodigoSAPLocal(DS.FieldByName('CodigoSAP').AsString);
          OutroDocumento := Trim(DS.FieldByName('OutroDocumento').AsString);

          if CodigoSAPEhNumericoLocal(CodigoSAP) then
          begin
            if AtualizarPorCodigoSAP(CodigoSAP) then
              Inc(AtualizadosCodigoSAP)
            else
              Inc(Ignorados);
          end
          else if CPFValidoParaBusca(DS.FieldByName('CPF').AsString,
            CPFNumerico, CPFFormatado) then
          begin
            if AtualizarPorCPF(CPFNumerico, CPFFormatado) then
              Inc(AtualizadosCPF)
            else
              Inc(Ignorados);
          end
          else if OutroDocumento <> '' then
          begin
            if AtualizarPorOutroDocumento(OutroDocumento) then
              Inc(AtualizadosOutroDocumento)
            else
              Inc(Ignorados);
          end
          else
            Inc(Ignorados);

          DS.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
      finally
        if HasBookmark then
        begin
          try
            if DS.BookmarkValid(Bmk) then
              DS.GotoBookmark(Bmk);
          finally
            DS.FreeBookmark(Bmk);
          end;
        end;

        DS.EnableControls;
      end;

      FrmDataModule.ADOConnectionConsulta.CommitTrans;
    except
      FrmDataModule.ADOConnectionConsulta.RollbackTrans;
      raise;
    end;
  finally
    FrmPrincipal.ProgressBarAtualizar;
    QAtualizar.Free;
  end;

  if FrmDataModule.ADOQueryExecutante.Active then
    FrmDataModule.ADOQueryExecutante.Requery();

  MessageBox(
    0,
    PChar(
      'Atualizacao concluida.' + sLineBreak +
      'Rateios atualizados por CodigoSAP: ' + IntToStr(AtualizadosCodigoSAP) + sLineBreak +
      'Rateios atualizados por CPF: ' + IntToStr(AtualizadosCPF) + sLineBreak +
      'Rateios atualizados por OutroDocumento: ' + IntToStr(AtualizadosOutroDocumento) + sLineBreak +
      'Registros ignorados sem chave valida ou sem correspondencia: ' + IntToStr(Ignorados)
    ),
    'Colibri',
    MB_OK or MB_ICONINFORMATION
  );
end;

procedure TFrmManifestoEmbarque.actCriarProgramacaoExecute(Sender: TObject);
var
  Bmk: TBookmark;
  DS: TDataSet;
  HasBookmark: Boolean;
  QExecutanteCadastro: TADOQuery;
  QInserirExecutante: TADOQuery;
  QInserirServico: TADOQuery;
  QAtualizarProgramacao: TADOQuery;
  ListaProgramacoes: TStringList;
  ListaQuantidadeGrupo: TStringList;
  ListaExecutantesGrupo: TStringList;
  CodigoSAP, CPFManifesto, CPFFormatadoManifesto, OutroDocumentoManifesto: String;
  IdentificadorExecutante, NomeExecutante, OrigemManifesto, OrigemExecutante: String;
  FuncaoExecutante, EmpresaExecutante, DocumentoExecutante: String;
  OutroDocumentoExecutante, TipoEtapaServico, Destino, DataProgramacao: String;
  ChaveGrupo, ChaveExecutanteGrupo, ItemIgnorado: String;
  IdxGrupo, CodigoProgramacaoDiaria, QuantidadeGrupo: Integer;
  QuantidadeCriadas, QuantidadeInseridas, QuantidadeDuplicadas: Integer;
  QuantidadeSemIdentificador, QuantidadeSemCadastro, QuantidadeSemTipoEtapa: Integer;
  QuantidadeSemDestino, QuantidadeSemData: Integer;
  RequisitantePT: Boolean;

  function CampoManifesto(const ANomeCampo: String): String;
  begin
    Result := '';
    if Assigned(DS.FindField(ANomeCampo)) then
      Result := Trim(DS.FieldByName(ANomeCampo).AsString);
  end;

  function CampoCadastro(const ANomeCampo: String): String;
  begin
    Result := '';
    if Assigned(QExecutanteCadastro.FindField(ANomeCampo)) then
      Result := Trim(QExecutanteCadastro.FieldByName(ANomeCampo).AsString);
  end;

  function NormalizarCodigoSAP(const ACodigoSAP: String): String;
  begin
    Result := FrmPrincipal.SomenteNumero(Trim(ACodigoSAP));
    while (Length(Result) > 1) and (Result[1] = '0') do
      Delete(Result, 1, 1);

    if Result = '0' then
      Result := '';
  end;

  function CPFValido(const ACPF: String; out ACPFNumerico,
    ACPFFormatado: String): Boolean;
  begin
    ACPFNumerico := FrmPrincipal.SomenteNumero(Trim(ACPF));
    ACPFFormatado := FrmPrincipal.VerificaCPF(ACPF);
    Result := (ACPFFormatado <> '') and (ACPFFormatado <> '000.000.000-00');
  end;

  function IdentificadorExecutanteManifesto(out AIdentificador, ACodigoSAP,
    ACPFNumerico, ACPFFormatado, AOutroDocumento: String): Boolean;
  begin
    ACodigoSAP := NormalizarCodigoSAP(CampoManifesto('CodigoSAP'));
    ACPFNumerico := '';
    ACPFFormatado := '';
    AOutroDocumento := '';
    AIdentificador := '';

    if ACodigoSAP <> '' then
    begin
      AIdentificador := 'SAP=' + ACodigoSAP;
      Exit(True);
    end;

    if CPFValido(CampoManifesto('CPF'), ACPFNumerico, ACPFFormatado) then
    begin
      AIdentificador := 'CPF=' + ACPFFormatado;
      Exit(True);
    end;

    AOutroDocumento := Trim(CampoManifesto('OutroDocumento'));
    if AOutroDocumento <> '' then
    begin
      AIdentificador := 'DOC=' + AOutroDocumento;
      Exit(True);
    end;

    Result := False;
  end;

  procedure AbrirCadastroExecutante(const ACodigoSAP, ACPFNumerico,
    ACPFFormatado, AOutroDocumento: String);
  begin
    QExecutanteCadastro.Close;

    if ACodigoSAP <> '' then
    begin
      QExecutanteCadastro.SQL.Text :=
        'SELECT tblExecutante.* ' +
        'FROM tblExecutante ' +
        'WHERE (CodigoSAP = :pCodigoSAP);';
      QExecutanteCadastro.Parameters.ParamByName('pCodigoSAP').Value := ACodigoSAP;
    end
    else if ACPFFormatado <> '' then
    begin
      QExecutanteCadastro.SQL.Text :=
        'SELECT tblExecutante.* ' +
        'FROM tblExecutante ' +
        'WHERE (CPF = :pCPF OR CPF = :pCPF2);';
      QExecutanteCadastro.Parameters.ParamByName('pCPF').Value := ACPFNumerico;
      QExecutanteCadastro.Parameters.ParamByName('pCPF2').Value := ACPFFormatado;
    end
    else
    begin
      QExecutanteCadastro.SQL.Text :=
        'SELECT tblExecutante.* ' +
        'FROM tblExecutante ' +
        'WHERE (OutroDocumento = :pOutroDocumento);';
      QExecutanteCadastro.Parameters.ParamByName('pOutroDocumento').Value := AOutroDocumento;
    end;

    QExecutanteCadastro.Open;
  end;

  function MontarChaveGrupo(const AData, ADestino, ATipoEtapaServico: String): String;
  begin
    Result := Trim(AData) + '|' + Trim(ADestino) + '|' + Trim(ATipoEtapaServico);
  end;

  procedure IncrementarQuantidadeGrupo(const AChaveGrupo: String);
  begin
    QuantidadeGrupo := StrToIntDef(ListaQuantidadeGrupo.Values[AChaveGrupo], 0) + 1;
    ListaQuantidadeGrupo.Values[AChaveGrupo] := IntToStr(QuantidadeGrupo);
  end;

begin
  DS := FrmDataModule.DataSourceManifestoEmbarque.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    MessageBox(0, 'Nenhum registro filtrado encontrado para programar.', 'Colibri', MB_OK or MB_ICONINFORMATION);
    Exit;
  end;

  if DS.State in dsEditModes then
    DS.Post;

  if Application.MessageBox(
    'Deseja criar programacoes com base nos registros filtrados do manifesto, agrupando por DataEmbarque, Destino e TipoEtapaServico?',
    '.::ATENCAO::.',
    MB_ICONQUESTION or MB_YESNO
  ) <> IDYES then
    Exit;

  HasBookmark := False;
  if not DS.IsEmpty then
  begin
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  QExecutanteCadastro := TADOQuery.Create(nil);
  QInserirExecutante := TADOQuery.Create(nil);
  QInserirServico := TADOQuery.Create(nil);
  QAtualizarProgramacao := TADOQuery.Create(nil);
  ListaProgramacoes := TStringList.Create;
  ListaQuantidadeGrupo := TStringList.Create;
  ListaExecutantesGrupo := TStringList.Create;
  try
    ListaProgramacoes.NameValueSeparator := '=';
    ListaQuantidadeGrupo.NameValueSeparator := '=';
    ListaExecutantesGrupo.CaseSensitive := False;
    ListaExecutantesGrupo.Sorted := False;

    QExecutanteCadastro.Connection := FrmDataModule.ADOConnectionConsulta;
    QExecutanteCadastro.ParamCheck := True;
    QExecutanteCadastro.SQL.Text := 'SELECT tblExecutante.* FROM tblExecutante WHERE (1 = 0);';

    QInserirExecutante.Connection := FrmDataModule.ADOConnectionColibri;
    QInserirExecutante.CursorType := ctStatic;
    QInserirExecutante.LockType := ltPessimistic;
    QInserirExecutante.SQL.Text :=
      'SELECT tblProgramacaoExecutante.* ' +
      'FROM tblProgramacaoExecutante ' +
      'WHERE (CodigoProgramacaoDiaria = 0);';
    QInserirExecutante.Open;

    QInserirServico.Connection := FrmDataModule.ADOConnectionColibri;
    QInserirServico.CursorType := ctStatic;
    QInserirServico.LockType := ltPessimistic;
    QInserirServico.SQL.Text :=
      'SELECT tblProgramacaoServico.* ' +
      'FROM tblProgramacaoServico ' +
      'WHERE (CodigoProgramacaoDiaria = 0);';
    QInserirServico.Open;

    QAtualizarProgramacao.Connection := FrmDataModule.ADOConnectionColibri;
    QAtualizarProgramacao.CursorType := ctStatic;
    QAtualizarProgramacao.LockType := ltPessimistic;
    QAtualizarProgramacao.ParamCheck := True;
    QAtualizarProgramacao.SQL.Text :=
      'SELECT tblProgramacaoDiaria.* ' +
      'FROM tblProgramacaoDiaria ' +
      'WHERE (idProgramacaoDiaria = :pCodigoProgramacaoDiaria);';

    QuantidadeCriadas := 0;
    QuantidadeInseridas := 0;
    QuantidadeDuplicadas := 0;
    QuantidadeSemIdentificador := 0;
    QuantidadeSemCadastro := 0;
    QuantidadeSemTipoEtapa := 0;
    QuantidadeSemDestino := 0;
    QuantidadeSemData := 0;

    FrmPrincipal.ProgressBarIncializa(
      DS.RecordCount,
      'Criando programacoes com base nos registros filtrados do manifesto.'
    );

    FrmDataModule.ADOConnectionColibri.BeginTrans;
    try
      DS.DisableControls;
      try
        DS.First;
        while not DS.Eof do
        begin
          if not IdentificadorExecutanteManifesto(IdentificadorExecutante,
            CodigoSAP, CPFManifesto, CPFFormatadoManifesto, OutroDocumentoManifesto) then
          begin
            Inc(QuantidadeSemIdentificador);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          if Assigned(DS.FindField('DataEmbarque')) and DS.FieldByName('DataEmbarque').IsNull then
          begin
            Inc(QuantidadeSemData);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          if Assigned(DS.FindField('DataEmbarque')) then
            DataProgramacao := FrmPrincipal.corrigirData(DS.FieldByName('DataEmbarque').AsDateTime)
          else
            DataProgramacao := '';

          if Trim(DataProgramacao) = '' then
          begin
            Inc(QuantidadeSemData);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          Destino := CampoManifesto('Destino');
          if Destino = '' then
          begin
            Inc(QuantidadeSemDestino);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;
          AbrirCadastroExecutante(CodigoSAP, CPFManifesto,
            CPFFormatadoManifesto, OutroDocumentoManifesto);
          if QExecutanteCadastro.IsEmpty then
          begin
            Inc(QuantidadeSemCadastro);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          TipoEtapaServico := CampoCadastro('txtTipoEtapaServico');
          if TipoEtapaServico = '' then
          begin
            Inc(QuantidadeSemTipoEtapa);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          ChaveGrupo := MontarChaveGrupo(DataProgramacao, Destino, TipoEtapaServico);
          ChaveExecutanteGrupo := ChaveGrupo + '|' + IdentificadorExecutante;
          if ListaExecutantesGrupo.IndexOf(ChaveExecutanteGrupo) >= 0 then
          begin
            Inc(QuantidadeDuplicadas);
            DS.Next;
            FrmPrincipal.ProgressBarIncremento(1);
            Continue;
          end;

          IdxGrupo := ListaProgramacoes.IndexOfName(ChaveGrupo);
          if IdxGrupo < 0 then
          begin
            CodigoProgramacaoDiaria := FrmPrincipal.incluirProgramacao(
              DataProgramacao,
              Destino,
              TipoEtapaServico,
              FrmPrincipal.logChave,
              FrmPrincipal.logMaquina,
              Now
            );

            if CodigoProgramacaoDiaria = 0 then
              raise Exception.Create('Nao foi possivel criar a programacao para o grupo ' + ChaveGrupo + '.');

            QInserirServico.Insert;
            QInserirServico.FieldByName('CodigoProgramacaoDiaria').AsInteger := CodigoProgramacaoDiaria;
            QInserirServico.FieldByName('TextoBreveOP').AsString :=
              FrmPrincipal.incluirServicoPadrao(TipoEtapaServico);
            QInserirServico.Post;

            ListaProgramacoes.Values[ChaveGrupo] := IntToStr(CodigoProgramacaoDiaria);
            ListaQuantidadeGrupo.Values[ChaveGrupo] := '0';
            Inc(QuantidadeCriadas);
          end
          else
            CodigoProgramacaoDiaria := StrToIntDef(ListaProgramacoes.ValueFromIndex[IdxGrupo], 0);

          NomeExecutante := CampoManifesto('NomeExecutante');
          if NomeExecutante = '' then
            NomeExecutante := CampoCadastro('NomeExecutante');

          OrigemManifesto := CampoManifesto('Origem');
          OrigemExecutante := OrigemManifesto;
          if OrigemExecutante = '' then
            OrigemExecutante := CampoCadastro('OrigemCadastro');

          FuncaoExecutante := CampoCadastro('txtFuncao');
          EmpresaExecutante := CampoCadastro('txtEmpresa');
          DocumentoExecutante := CampoCadastro('Documento');
          OutroDocumentoExecutante := CampoCadastro('OutroDocumento');

          RequisitantePT := False;
          if Assigned(QExecutanteCadastro.FindField('RequisitantePT')) then
          begin
            try
              RequisitantePT := QExecutanteCadastro.FieldByName('RequisitantePT').AsBoolean;
            except
              RequisitantePT := SameText(CampoCadastro('RequisitantePT'), 'TRUE');
            end;
          end;

          QInserirExecutante.Insert;
          QInserirExecutante.FieldByName('CodigoProgramacaoDiaria').AsInteger := CodigoProgramacaoDiaria;
          QInserirExecutante.FieldByName('NomeExecutante').AsString := NomeExecutante;
          QInserirExecutante.FieldByName('Origem').AsString := OrigemExecutante;
          QInserirExecutante.FieldByName('Funcao').AsString := FuncaoExecutante;
          QInserirExecutante.FieldByName('Empresa').AsString := EmpresaExecutante;
          QInserirExecutante.FieldByName('CodigoSAP').AsString := CodigoSAP;
          QInserirExecutante.FieldByName('Documento').AsString := DocumentoExecutante;
          QInserirExecutante.FieldByName('OutroDocumento').AsString := OutroDocumentoExecutante;
          QInserirExecutante.FieldByName('RequisitantePT').AsBoolean := RequisitantePT;
          QInserirExecutante.FieldByName('StatusProgramacao').AsString := 'Aprovado';
          QInserirExecutante.FieldByName('MotivoProgramacao').AsString := '';
          QInserirExecutante.FieldByName('PalavraChaveProgramacao').AsString := '';
          QInserirExecutante.FieldByName('AvaliadoPorProgramacao').AsString := FrmPrincipal.logChave;
          QInserirExecutante.FieldByName('DataAvaliacaoProgramacao').AsDateTime := Now;
          QInserirExecutante.FieldByName('ComputadorProgramacao').AsString := FrmPrincipal.logMaquina;
          QInserirExecutante.Post;

          ListaExecutantesGrupo.Add(ChaveExecutanteGrupo);
          IncrementarQuantidadeGrupo(ChaveGrupo);
          Inc(QuantidadeInseridas);

          DS.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;

        for IdxGrupo := 0 to ListaProgramacoes.Count - 1 do
        begin
          ChaveGrupo := ListaProgramacoes.Names[IdxGrupo];
          CodigoProgramacaoDiaria := StrToIntDef(ListaProgramacoes.ValueFromIndex[IdxGrupo], 0);
          QuantidadeGrupo := StrToIntDef(ListaQuantidadeGrupo.Values[ChaveGrupo], 0);
          if (CodigoProgramacaoDiaria > 0) and (QuantidadeGrupo > 0) then
          begin
            QAtualizarProgramacao.Close;
            QAtualizarProgramacao.Parameters.ParamByName('pCodigoProgramacaoDiaria').Value := CodigoProgramacaoDiaria;
            QAtualizarProgramacao.Open;
            if not QAtualizarProgramacao.IsEmpty then
            begin
              QAtualizarProgramacao.Edit;
              QAtualizarProgramacao.FieldByName('NumExecutantes').AsInteger := QuantidadeGrupo;
              QAtualizarProgramacao.FieldByName('NumCancelados').AsInteger := 0;
              QAtualizarProgramacao.FieldByName('NumAprovados').AsInteger := QuantidadeGrupo;
              QAtualizarProgramacao.Post;
            end;
          end;
        end;
      finally
        if HasBookmark then
        begin
          try
            if DS.BookmarkValid(Bmk) then
              DS.GotoBookmark(Bmk);
          finally
            DS.FreeBookmark(Bmk);
          end;
        end;

        DS.EnableControls;
      end;

      FrmDataModule.ADOConnectionColibri.CommitTrans;
    except
      FrmDataModule.ADOConnectionColibri.RollbackTrans;
      raise;
    end;

    if FrmDataModule.ADOQueryManifestoEmbarque.Active then
      FrmDataModule.ADOQueryManifestoEmbarque.Requery();

    ItemIgnorado :=
      'Programacoes criadas: ' + IntToStr(QuantidadeCriadas) + sLineBreak +
      'Executantes inseridos: ' + IntToStr(QuantidadeInseridas) + sLineBreak +
      'Duplicados no mesmo grupo: ' + IntToStr(QuantidadeDuplicadas) + sLineBreak +
      'Sem CodigoSAP/CPF/OutroDocumento valido: ' + IntToStr(QuantidadeSemIdentificador) + sLineBreak +
      'Sem cadastro em tblExecutante: ' + IntToStr(QuantidadeSemCadastro) + sLineBreak +
      'Sem TipoEtapaServico no cadastro: ' + IntToStr(QuantidadeSemTipoEtapa) + sLineBreak +
      'Sem Destino: ' + IntToStr(QuantidadeSemDestino) + sLineBreak +
      'Sem DataEmbarque: ' + IntToStr(QuantidadeSemData);

    MessageBox(0, PChar(ItemIgnorado), 'Colibri', MB_OK or MB_ICONINFORMATION);
  finally
    FrmPrincipal.ProgressBarAtualizar;
    ListaExecutantesGrupo.Free;
    ListaQuantidadeGrupo.Free;
    ListaProgramacoes.Free;
    QAtualizarProgramacao.Free;
    QInserirServico.Free;
    QInserirExecutante.Free;
    QExecutanteCadastro.Free;
  end;
end;

procedure TFrmManifestoEmbarque.InicializarFiltrosRapidos;
begin
  FAtualizandoFiltroRapido := True;
  try
    MarcarTodosFiltroRapido;
  finally
    FAtualizandoFiltroRapido := False;
  end;
end;

procedure TFrmManifestoEmbarque.MarcarTodosFiltroRapido;
var
  I: Integer;
begin
  for I := 0 to clbTipoEmbarque.Items.Count - 1 do
    clbTipoEmbarque.Checked[I] := True;

  for I := 0 to clbEmbarqueDesembarque.Items.Count - 1 do
    clbEmbarqueDesembarque.Checked[I] := True;
end;

function TFrmManifestoEmbarque.SQLFiltroRapidoLista(ACheckList: TCheckListBox;
  const ACampo: string): string;
var
  I: Integer;
  QuantidadeMarcados: Integer;
  ListaValores: string;
  ValorItem: string;
begin
  Result := '';
  if (ACheckList = nil) or (ACheckList.Items.Count = 0) then
    Exit;

  QuantidadeMarcados := 0;
  ListaValores := '';
  for I := 0 to ACheckList.Items.Count - 1 do
  begin
    if not ACheckList.Checked[I] then
      Continue;

    Inc(QuantidadeMarcados);
    ValorItem := AnsiUpperCase(Trim(ACheckList.Items[I]));
    ValorItem := StringReplace(ValorItem, '''', '''''', [rfReplaceAll]);
    if ListaValores <> '' then
      ListaValores := ListaValores + ', ';
    ListaValores := ListaValores + QuotedStr(ValorItem);
  end;

  if QuantidadeMarcados = 0 then
  begin
    Result := '1 = 0';
    Exit;
  end;

  if QuantidadeMarcados = ACheckList.Items.Count then
    Exit;

  Result := 'UCase(Trim(' + ACampo + ')) IN (' + ListaValores + ')';
end;

function TFrmManifestoEmbarque.SQLFiltroRapido: string;

  procedure AdicionarFiltro(var ADestino: string; const AFiltro: string);
  begin
    if AFiltro = '' then
      Exit;

    if ADestino <> '' then
      ADestino := ADestino + ' AND ';
    ADestino := ADestino + '(' + AFiltro + ')';
  end;

begin
  Result := '';
  AdicionarFiltro(Result, SQLFiltroRapidoLista(clbTipoEmbarque, '[TipoEmbarque]'));
  AdicionarFiltro(Result, SQLFiltroRapidoLista(clbEmbarqueDesembarque,
    '[EmbarqueDesembarque]'));
end;

procedure TFrmManifestoEmbarque.clbFiltroRapidoClickCheck(Sender: TObject);
begin
  if FAtualizandoFiltroRapido then
    Exit;

  actProcurar.Execute;
end;

procedure TFrmManifestoEmbarque.btnRestaurarFiltrosRapidosClick(
  Sender: TObject);
begin
  FAtualizandoFiltroRapido := True;
  try
    MarcarTodosFiltroRapido;
  finally
    FAtualizandoFiltroRapido := False;
  end;

  actProcurar.Execute;
end;
procedure TFrmManifestoEmbarque.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase,SQLFiltroRapidoAtual,
    DataProcuraIncio,DataProcuraFim: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  //======================================================
  SQLString:= BuildFilterSQL(DBGridManifestoEmbarque,false);
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;

  SQLFiltroRapidoAtual := SQLFiltroRapido;
  if SQLFiltroRapidoAtual <> '' then
    SQLString := SQLString + ' AND ' + SQLFiltroRapidoAtual;

  SQLBase:= 'SELECT tblManifestoEmbarque.* FROM tblManifestoEmbarque '+
  'WHERE (DataEmbarque BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'+
  SQLString+' ORDER BY DataEmbarque, Destino, NomeExecutante;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryManifestoEmbarque,StatusBarManifestoEmbarque);
end;

procedure TFrmManifestoEmbarque.dataFimCloseUp(Sender: TObject);
begin
  actProcurar.Execute;
end;

procedure TFrmManifestoEmbarque.dataFimKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    actProcurar.Execute;
end;

procedure TFrmManifestoEmbarque.dataInicioCloseUp(Sender: TObject);
begin
  actProcurar.Execute;
end;

procedure TFrmManifestoEmbarque.dataInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    actProcurar.Execute;
end;

procedure TFrmManifestoEmbarque.DBGridManifestoEmbarqueDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  with DBGridManifestoEmbarque do
  begin
    if (Column.Field.FieldName = 'EmbarqueDesembarque') then
    begin
      if SameText(Column.Field.AsString, VALOR_EMBARQUE) then
        Canvas.Brush.Color := clSkyBlue
      else if SameText(Column.Field.AsString, VALOR_DESEMBARQUE) then
        Canvas.Brush.Color := clWebBurlywood;

      Font.Color := clBlack;
      Canvas.FillRect(Rect);
      DefaultDrawColumnCell(Rect, DataCol, Column, State);
    end;
    if (Column.Field.FieldName = 'TipoEmbarque') then
    begin
      if SameText(Column.Field.AsString, VALOR_TIPO_EMBARQUE_FIXO) then
        Canvas.Brush.Color := clWebDarkOrange
      else if SameText(Column.Field.AsString, VALOR_TIPO_EMBARQUE_SOV) then
        Canvas.Brush.Color := clWebGreenYellow
      else if SameText(Column.Field.AsString, VALOR_TIPO_EMBARQUE_BATE_VOLTA) then
        Canvas.Brush.Color := clWebDeepPink
      else if SameText(Column.Field.AsString, VALOR_TIPO_EMBARQUE_AEREO) then
        Canvas.Brush.Color := clWebCornFlowerBlue;

      Font.Color := clBlack;
      Canvas.FillRect(Rect);
      DefaultDrawColumnCell(Rect, DataCol, Column, State);
    end;
  end;

end;

procedure TFrmManifestoEmbarque.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmManifestoEmbarque:=nil;
end;

procedure TFrmManifestoEmbarque.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  //Incicializa��o
  actAtualizarExecutante.Hint := 'Atualizar tabela Executante com base nos registros filtrados do manifesto';
  actExcluirFiltrados.Hint := 'Excluir todos os registros filtrados do manifesto';
  btnRestaurarFiltrosRapidos.Hint := 'Restaurar todas as opcoes dos filtros rapidos';
  InicializarFiltrosRapidos;

  dataInicio.Date:= IncDay(now,1);
  dataFim.Date:= IncDay(now,1);
  actProcurar.Execute;

  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)) then
  begin
    actCriarProgramacao.Enabled:= true;
    actAtualizarExecutante.Enabled:= true;
    actExcelImportar.Enabled:= true;
    DBGridManifestoEmbarque.ReadOnly:= false;
    DBNavigatorManifestoEmbarque.Enabled:= true;
  end
  else if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT)) then
  begin
    actCriarProgramacao.Enabled:= false;
    actAtualizarExecutante.Enabled:= true;
    actExcelImportar.Enabled:= true;
    DBGridManifestoEmbarque.ReadOnly:= false;
    DBNavigatorManifestoEmbarque.Enabled:= true;
  end
  else
  begin
    actCriarProgramacao.Enabled:= false;
    actAtualizarExecutante.Enabled:= false;
    actExcelImportar.Enabled:= false;
    DBGridManifestoEmbarque.ReadOnly:= true;
    DBNavigatorManifestoEmbarque.Enabled:= false;
  end;
end;

procedure TFrmManifestoEmbarque.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmManifestoEmbarque.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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


