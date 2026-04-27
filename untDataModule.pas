unit untDataModule;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Data.Win.ADODB,Winapi.Windows,
  ActiveX, System.Variants, ADOInt, untDBGridFilter, Vcl.ComCtrls, uZucchi, uMensagens,
  uFuncaoExecutanteUtils, uEmpresaExecutanteUtils;

type
  TFrmDataModule = class(TDataModule)
    ADOQueryProgramacaoServico_Cadastro: TADOQuery;
    ADOQueryProgramacaoExecutante_Cadastro: TADOQuery;
    ADOQueryProgramacaoDiaria_Consulta: TADOQuery;
    DataSourceProgramacaoDiaria_Consulta: TDataSource;
    DataSourceProgramacaoExecutante_Cadastro: TDataSource;
    DataSourceProgramacaoServico_Cadastro: TDataSource;
    ADOConnectionColibri: TADOConnection;
    ADOQueryPlataforma: TADOQuery;
    DataSourcePlataforma: TDataSource;
    ADOQueryExecutante: TADOQuery;
    DataSourceExecutante: TDataSource;
    ADOQueryUsuario: TADOQuery;
    DataSourceUsuario: TDataSource;
    DataSourceProgramacaoDiaria_Cadastro: TDataSource;
    ADOQueryProgramacaoDiaria_Cadastro: TADOQuery;
    ADOQueryConsultaExecutante_DataCodigoSAP: TADOQuery;
    DataSourceConsultaExecutante_DataCodigoSAP: TDataSource;
    DataSourceInserirExecutante: TDataSource;
    ADOQueryInserirExecutante1: TADOQuery;
    DataSourceInserirServico: TDataSource;
    ADOQueryInserirServico: TADOQuery;
    DataSourceInserirProgramacao: TDataSource;
    ADOQueryInserirProgramacao: TADOQuery;
    ADOQueryConsultaProgramacao_ID: TADOQuery;
    DataSourceConsultaProgramacao_ID: TDataSource;
    DataSourceConsultaPlataforma: TDataSource;
    ADOQueryConsultaPlataforma: TADOQuery;
    DataSourceInserirExecutanteCadastro: TDataSource;
    ADOQueryInserirExecutanteCadastro: TADOQuery;
    DataSourceConsultaExecutante_TipoEtapaServico: TDataSource;
    ADOQueryConsultaExecutante_TipoEtapaServico: TADOQuery;
    ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico: TADOQuery;
    DataSourceConsultaCodigoTipoEtapaServico_TipoEtapaServico: TDataSource;
    ADOQueryTipoEtapaServico: TADOQuery;
    DataSourceTipoEtapaServico: TDataSource;
    ADOQueryConsultaTipoEtapaServico_ID: TADOQuery;
    DataSourceConsultaTipoEtapaServico_ID: TDataSource;
    ADOQueryGerenciarSolicitacoes: TADOQuery;
    DataSourceGerenciarSolicitacoes: TDataSource;
    ADOQueryConsultaPlataforma_Nome: TADOQuery;
    DataSourceConsultaPlataforma_Nome: TDataSource;
    DataSourceGerenciarExecutante: TDataSource;
    ADOQueryGerenciarExecutante: TADOQuery;
    DataSourceGerenciarServico: TDataSource;
    ADOQueryGerenciarServico: TADOQuery;
    DataSourceGerenciarProgramacao: TDataSource;
    ADOQueryGerenciarProgramacao: TADOQuery;
    DataSourceConsultaProgramacao_Destino_DataProgramacao: TDataSource;
    ADOQueryConsultaProgramacao_Destino_DataProgramacao: TADOQuery;
    DataSourceConsultaProgramacaoExecutante_ID: TDataSource;
    ADOQueryConsultaProgramacaoExecutante_ID: TADOQuery;
    DataSourceEmbarcacoes: TDataSource;
    ADOQueryEmbarcacoes: TADOQuery;
    ADOQueryTM_Embarcacao: TADOQuery;
    DataSourceTM_Embarcacao: TDataSource;
    ADOQueryRoteamento: TADOQuery;
    DataSourceRoteamento: TDataSource;
    ADOQueryTM_Plataformas: TADOQuery;
    DataSourceTM_Plataformas: TDataSource;
    ADOQueryTM_RotaExecutantes: TADOQuery;
    DataSourceTM_RotaExecutantes: TDataSource;
    DataSourceExecutantesDisponiveis: TDataSource;
    ADOQueryExecutantesDisponiveis: TADOQuery;
    ADOQueryConsultaExecutantes_DataProgramacao: TADOQuery;
    DataSourceConsultaExecutantes_DataProgramacao: TDataSource;
    ADOQueryAux_Rota_Distribuicao: TADOQuery;
    DataSourceAux_Rota_Distribuicao: TDataSource;
    ADOQueryExcluirVinculoExecutante: TADOQuery;
    DataSourceExcluirVinculoExecutante: TDataSource;
    ADOQueryTemporarioDBColibri: TADOQuery;
    DataSourceTemporarioDBColibri: TDataSource;
    ADOQueryConsultaServicosProgramados: TADOQuery;
    DataSourceConsultaServicosProgramados: TDataSource;
    ADOQueryConsultaEmbarcacao: TADOQuery;
    DataSourceConsultaEmbarcacao: TDataSource;
    ADOQueryConsultaExecutantesProgramados: TADOQuery;
    DataSourceConsultaExecutantesProgramados: TDataSource;
    DataSourceColibri: TDataSource;
    DataSourceGeradores: TDataSource;
    ADOQueryColibri: TADOQuery;
    ADOQueryGeradores: TADOQuery;
    ADOQueryConsultaExecutante_CodigoSAP: TADOQuery;
    DataSourceConsultaExecutante_CodigoSAP: TDataSource;
    DataSourceConsultaProgramacaoServico_ID: TDataSource;
    ADOQueryConsultaProgramacaoServico_ID: TADOQuery;
    DataSourceContarExecCancelados_ID: TDataSource;
    ADOQueryNumAprovado: TADOQuery;
    ADOQueryNumCancelado: TADOQuery;
    DataSourceContarExecTotal_ID: TDataSource;
    ADOQueryProgramacaoExecutante_Consulta: TADOQuery;
    DataSourceProgramacaoExecutante_Consulta: TDataSource;
    DataSourceProgramacaoServico_Consulta: TDataSource;
    ADOQueryProgramacaoServico_Consulta: TADOQuery;
    ADOQueryNumExecutantes: TADOQuery;
    DataSourceNumExecutantes: TDataSource;
    ADOQueryProgramados: TADOQuery;
    DataSourceProgramados: TDataSource;
    DataSourceCentroTrabalho_Descricao: TDataSource;
    ADOQueryCentroTrabalho_Descricao: TADOQuery;
    DataSourceTipoEtapaServico_Carteira: TDataSource;
    ADOQueryTipoEtapaServico_Carteira: TADOQuery;
    DataSourceCadastroUsuario: TDataSource;
    ADOQueryCadastroUsuario: TADOQuery;
    DataSourceTemporarioDBConsulta1: TDataSource;
    ADOQueryTemporarioDBConsulta1: TADOQuery;
    ADOQueryTotalCampo: TADOQuery;
    DataSourceTotalCampo: TDataSource;
    ADOQueryTotalOrigem: TADOQuery;
    DataSourceTotalOrigem: TDataSource;
    ADOQueryPassageirosTM_SIM: TADOQuery;
    DataSourcePassageirosTM_SIM: TDataSource;
    ADOQueryPassageirosTM_NAO: TADOQuery;
    DataSourcePassageirosTM_NAO: TDataSource;
    DataSourceConsultaExecutante_Documento_Data: TDataSource;
    ADOQueryConsultaExecutante_Documento_Data: TADOQuery;
    ADOConnectionImportar: TADOConnection;
    DataSourceProgramacaoDiaria_Importar: TDataSource;
    ADOQueryProgramacaoDiaria_Importar: TADOQuery;
    DataSourceProgramacaoExecutante_Importar: TDataSource;
    ADOQueryProgramacaoExecutante_Importar: TADOQuery;
    DataSourceProgramacaoServico_Importar: TDataSource;
    ADOQueryProgramacaoServico_Importar: TADOQuery;
    DataSourceTemporarioDBImportar: TDataSource;
    ADOQueryTemporarioDBImportar: TADOQuery;
    ADOConnectionConsulta: TADOConnection;
    DataSourceExecutante_TipoEtapaServico: TDataSource;
    ADOQueryExecutante_TipoEtapaServico: TADOQuery;
    DataSourceMovimentacaoCarga: TDataSource;
    ADOQueryMovimentacaoCarga: TADOQuery;
    DataSourceCondicaoMar: TDataSource;
    ADOQueryCondicaoMar: TADOQuery;
    DataSourceCondicaoEmbarcacao: TDataSource;
    ADOQueryCondicaoEmbarcacao: TADOQuery;
    DataSourceProgramacaoNotas: TDataSource;
    ADOQueryProgramacaoNotas: TADOQuery;
    ADOQueryContadorSolicitacao: TADOQuery;
    DataSourceContadorSolicitacao: TDataSource;
    DataSourceProgramacaoCalendario: TDataSource;
    ADOQueryProgramacaoCalendario1: TADOQuery;
    ADOQueryPlataformaControle: TADOQuery;
    DataSourcePlataformaControle: TDataSource;
    ADOQueryPalavraChave: TADOQuery;
    DataSourcePalavraChave: TDataSource;
    ADOQueryAuxTabelaRT: TADOQuery;
    DataSourceAuxTabelaRT: TDataSource;
    ADOQueryCancelarRT: TADOQuery;
    DataSourceCancelarRT: TDataSource;
    ADOQueryGestaoRT: TADOQuery;
    DataSourceGestaoRT: TDataSource;
    ADOConnectionRT: TADOConnection;
    DataSourceTemporarioRT: TDataSource;
    ADOQueryTemporarioRT: TADOQuery;
    ADOQueryConfigRT: TADOQuery;
    DataSourceConfigRT: TDataSource;
    DataSourceSAPImport: TDataSource;
    ADOQuerySAPImport: TADOQuery;
    ADOQueryRegrasRecolhimento: TADOQuery;
    DataSourceRegrasRecolhimento: TDataSource;
    ADOQueryFrequenciaResumo: TADOQuery;
    DataSourceFrequenciaResumo: TDataSource;
    ADOQueryFrequenciaDetalhe: TADOQuery;
    DataSourceFrequenciaDetalhe: TDataSource;
    ADOQueryManifestoEmbarque: TADOQuery;
    DataSourceManifestoEmbarque: TDataSource;
    ADOConnectionProntidao: TADOConnection;
    procedure ADOQueryEmbarcacoesBeforePost(DataSet: TDataSet);
    procedure ADOQueryRoteamentoBeforePost(DataSet: TDataSet);
    procedure ADOQueryGeradoresBeforePost(DataSet: TDataSet);
    procedure ADOQueryExecutanteBeforePost(DataSet: TDataSet);
    procedure ADOQueryTipoEtapaServicoBeforePost(DataSet: TDataSet);
    procedure ADOQueryCadastroUsuarioBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaAfterRefresh(DataSet: TDataSet);
    procedure ADOQueryMovimentacaoCargaBeforePost(DataSet: TDataSet);
    procedure ADOQueryCondicaoMarBeforePost(DataSet: TDataSet);
    procedure ADOQueryCondicaoEmbarcacaoBeforePost(DataSet: TDataSet);
    procedure ADOQueryProgramacaoNotasBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaControleBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaAfterPost(DataSet: TDataSet);
    procedure ADOQueryManifestoEmbarqueBeforePost(DataSet: TDataSet);
  private

    { Private declarations }
  public
    naoGravar: Boolean;
    procedure setFilterDBGrid(DBGrid: TFilterDBGrid);
    procedure ProcuraQueryCompleta(
      const SQLBase   : String;
            sourceQuery: TADOQuery;
            StatusBar : TStatusBar);
    { Public declarations }
  end;

var
  FrmDataModule: TFrmDataModule;

implementation
  uses untPrincipal,untProgramacaoDiaria,untGerenciarSolicitacoes,
  untExecutante,untCondicaoEmbarcacao,untFrmSobre, untFrmMagicFiltro, untFrmManifestoEmbarque;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TFrmDataModule }

procedure TFrmDataModule.setFilterDBGrid(DBGrid: TFilterDBGrid);
begin
  // Mantido apenas por compatibilidade. O TFilterDBGrid agora usa
  // automaticamente o proprio DataSource quando FilterDataSource nao e informado.
end;

procedure TFrmDataModule.ProcuraQueryCompleta(
  const SQLBase   : String;
        sourceQuery: TADOQuery;
        StatusBar : TStatusBar);
var
  NumeroLinhas: Integer;
begin
  try
    sourceQuery.Close;
    sourceQuery.SQL.Text := SQLBase;
    sourceQuery.Open;
    // Vincula ao DataSource e atualiza o StatusBar
    NumeroLinhas := sourceQuery.RecordCount;
    StatusBar.Panels[0].Text := 'N� Registros: ' + IntToStr(NumeroLinhas);
    AutoFitStatusBar(StatusBar);

    if Assigned(FrmMagicFiltro) then
      FrmMagicFiltro.MemoSQLMagicFiltro.Text := SQLBase;

  except
    on E: Exception do
    begin
      TMensagens.Erro(MSG_FALHA_CONSULTA + sLineBreak + E.Message);
      Abort;
    end;
  end;
end;

procedure TFrmDataModule.ADOQueryCadastroUsuarioBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceCadastroUsuario.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceCadastroUsuario.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceCadastroUsuario.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryCondicaoEmbarcacaoBeforePost(
  DataSet: TDataSet);
begin
  FrmDataModule.DataSourceCondicaoEmbarcacao.DataSet.
  FieldByName('AvaliadoPor').AsString:=FrmPrincipal.logChave;
  FrmDataModule.DataSourceCondicaoEmbarcacao.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceCondicaoEmbarcacao.DataSet.
  FieldByName('Computador').AsString:=FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryCondicaoMarBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceCondicaoMar.DataSet.
  FieldByName('AvaliadoPor').AsString:=FrmPrincipal.logChave;
  FrmDataModule.DataSourceCondicaoMar.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:=now;
  FrmDataModule.DataSourceCondicaoMar.DataSet.
  FieldByName('Computador').AsString:=FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryEmbarcacoesBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceEmbarcacoes.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceEmbarcacoes.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceEmbarcacoes.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryExecutanteBeforePost(DataSet: TDataSet);
  var
    CodigoSAP, CPF: String;
begin
  if naoGravar = false then
  begin
    if Assigned(FrmExecutante) then
    begin
      CPF:=  FrmDataModule.ADOQueryExecutante.FieldByName('CPF').AsString;
      if FrmPrincipal.VerificaCPF(CPF) = '' then
      begin
        FrmDataModule.ADOQueryExecutante.FieldByName('CPF').AsString:= '000.000.000-00';
      end
      else
      begin
        FrmDataModule.ADOQueryExecutante.FieldByName('CPF').AsString:= FrmPrincipal.FormatarCPF(CPF);
      end;
    end;
    CodigoSAP:= FrmDataModule.DataSourceExecutante.DataSet.FieldByName('CodigoSAP').AsString;
    CodigoSAP:= FrmPrincipal.SomenteNumero(CodigoSAP);
    if ((CodigoSAP = '')OR(CodigoSAP='0')) then
      CodigoSAP:= 'NA';
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('DataAtualizacao').AsDateTime:= now;
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('CodigoSAP').AsString:= CodigoSAP;
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('txtFuncao').AsString:= NormalizarFuncaoExecutante(
      FrmDataModule.DataSourceExecutante.DataSet.FieldByName('txtFuncao').AsString);
    FrmDataModule.DataSourceExecutante.DataSet.
    FieldByName('txtEmpresa').AsString:= NormalizarEmpresaExecutante(
      FrmDataModule.DataSourceExecutante.DataSet.FieldByName('txtEmpresa').AsString);
  end;
end;

procedure TFrmDataModule.ADOQueryGeradoresBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceGeradores.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceGeradores.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceGeradores.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryManifestoEmbarqueBeforePost(DataSet: TDataSet);
  var
    CodigoSAP, CPF, CPFNumerico, OutroDocumento: String;
    QExecutante: TADOQuery;

    function CampoEstaVazio(const ANomeCampo: String): Boolean;
    var
      F: TField;
    begin
      F := DataSet.FindField(ANomeCampo);
      Result := Assigned(F) and (Trim(F.AsString) = '');
    end;

    procedure PreencherCampoSeVazio(const ACampoDestino, ACampoOrigem: String);
    var
      CampoDestino, CampoOrigem: TField;
      ValorOrigem: String;
    begin
      CampoDestino := DataSet.FindField(ACampoDestino);
      CampoOrigem := QExecutante.FindField(ACampoOrigem);
      if (CampoDestino = nil) or (CampoOrigem = nil) then
        Exit;

      if Trim(CampoDestino.AsString) <> '' then
        Exit;

      ValorOrigem := Trim(CampoOrigem.AsString);
      if ValorOrigem = '' then
        Exit;

      CampoDestino.AsString := ValorOrigem;
    end;

    procedure PreencherCampoSePadrao(const ACampoDestino, ACampoOrigem,
      AValorPadraoDestino: String);
    var
      CampoDestino, CampoOrigem: TField;
      ValorDestino, ValorOrigem: String;
    begin
      CampoDestino := DataSet.FindField(ACampoDestino);
      CampoOrigem := QExecutante.FindField(ACampoOrigem);
      if (CampoDestino = nil) or (CampoOrigem = nil) then
        Exit;

      ValorDestino := Trim(CampoDestino.AsString);
      if ValorDestino <> AValorPadraoDestino then
        Exit;

      ValorOrigem := Trim(CampoOrigem.AsString);
      if (ValorOrigem = '') or (ValorOrigem = AValorPadraoDestino) then
        Exit;

      CampoDestino.AsString := ValorOrigem;
    end;
begin
  if Assigned(FrmManifestoEmbarque) then
  begin
    CPF := DataSet.FieldByName('CPF').AsString;
    OutroDocumento := Trim(DataSet.FieldByName('OutroDocumento').AsString);
    CodigoSAP := FrmPrincipal.SomenteNumero(DataSet.FieldByName('CodigoSAP').AsString);

    if FrmPrincipal.VerificaCPF(CPF) = '' then
      CPF := '000.000.000-00'
    else
      CPF := FrmPrincipal.FormatarCPF(CPF);

    DataSet.FieldByName('CPF').AsString := CPF;
    CPFNumerico := FrmPrincipal.SomenteNumero(CPF);

    if ((CodigoSAP = '') OR (CodigoSAP = '0')) then
      CodigoSAP := 'NA';
    DataSet.FieldByName('CodigoSAP').AsString := CodigoSAP;


    if CampoEstaVazio('NomeExecutante') or
       CampoEstaVazio('txtTipoEtapaServico') or
       CampoEstaVazio('txtEmpresa') or
       CampoEstaVazio('txtFuncao') or
       (CPF = '000.000.000-00') or
       (CodigoSAP = 'NA') or
       CampoEstaVazio('OutroDocumento') then
    begin
      QExecutante := TADOQuery.Create(nil);
      try
        QExecutante.Connection := FrmDataModule.ADOConnectionConsulta;
        QExecutante.ParamCheck := True;

        if CodigoSAP <> 'NA' then
        begin
          QExecutante.SQL.Text :=
            'SELECT tblExecutante.* ' +
            'FROM tblExecutante ' +
            'WHERE (CodigoSAP = :pCodigoSAP);';
          QExecutante.Parameters.ParamByName('pCodigoSAP').Value := CodigoSAP;
        end
        else if CPF <> '000.000.000-00' then
        begin
          QExecutante.SQL.Text :=
            'SELECT tblExecutante.* ' +
            'FROM tblExecutante ' +
            'WHERE (CPF = :pCPF OR CPF = :pCPF2);';
          QExecutante.Parameters.ParamByName('pCPF').Value := CPFNumerico;
          QExecutante.Parameters.ParamByName('pCPF2').Value := CPF;
        end
        else if OutroDocumento <> '' then
        begin
          QExecutante.SQL.Text :=
            'SELECT tblExecutante.* ' +
            'FROM tblExecutante ' +
            'WHERE (OutroDocumento = :pOutroDocumento);';
          QExecutante.Parameters.ParamByName('pOutroDocumento').Value := OutroDocumento;
        end
        else
          Exit;

        QExecutante.Open;
        if not QExecutante.IsEmpty then
        begin
          PreencherCampoSeVazio('NomeExecutante', 'NomeExecutante');
          PreencherCampoSeVazio('txtTipoEtapaServico', 'txtTipoEtapaServico');
          PreencherCampoSeVazio('txtEmpresa', 'txtEmpresa');
          PreencherCampoSeVazio('txtFuncao', 'txtFuncao');
          PreencherCampoSePadrao('CPF', 'CPF', '000.000.000-00');
          PreencherCampoSePadrao('CodigoSAP', 'CodigoSAP', 'NA');
          PreencherCampoSeVazio('OutroDocumento', 'OutroDocumento');
        end;
      finally
        QExecutante.Free;
      end;
    end;
  end;
end;

procedure TFrmDataModule.ADOQueryMovimentacaoCargaBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceMovimentacaoCarga.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceMovimentacaoCarga.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceMovimentacaoCarga.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryPlataformaAfterRefresh(DataSet: TDataSet);
begin
  ADOQueryPlataforma.Active:= false;
  ADOQueryPlataforma.Active:= true;
end;

procedure TFrmDataModule.ADOQueryPlataformaBeforePost(DataSet: TDataSet);
var
  PlataformaAtual: string;
begin
  if DataSet.FindField('PrioridadeDistribuicao') <> nil then
    if DataSet.FieldByName('PrioridadeDistribuicao').IsNull then
      DataSet.FieldByName('PrioridadeDistribuicao').AsInteger := 99;

  if DataSet.FindField('booleanHubPrincipal') <> nil then
    if DataSet.FieldByName('booleanHubPrincipal').IsNull then
      DataSet.FieldByName('booleanHubPrincipal').AsBoolean := False;

  if DataSet.FindField('booleanGangwayAqua') <> nil then
    if DataSet.FieldByName('booleanGangwayAqua').IsNull then
      DataSet.FieldByName('booleanGangwayAqua').AsBoolean := False;

  if DataSet.FindField('booleanGangwaySOV') <> nil then
    if DataSet.FieldByName('booleanGangwaySOV').IsNull then
      DataSet.FieldByName('booleanGangwaySOV').AsBoolean := False;

  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;

  if (DataSet.FindField('booleanHubPrincipal') <> nil) and
     DataSet.FieldByName('booleanHubPrincipal').AsBoolean then
  begin
    PlataformaAtual := Trim(DataSet.FieldByName('Plataforma').AsString);
    if PlataformaAtual <> '' then
      FrmDataModule.ADOConnectionConsulta.Execute(
        'UPDATE tblPlataforma SET booleanHubPrincipal = False ' +
        'WHERE Plataforma <> ' + QuotedStr(PlataformaAtual)
      );
  end;
end;

procedure TFrmDataModule.ADOQueryPlataformaAfterPost(DataSet: TDataSet);
begin
  ADOQueryPlataforma.Requery;
end;

procedure TFrmDataModule.ADOQueryPlataformaControleBeforePost(
  DataSet: TDataSet);
begin
  FrmDataModule.DataSourcePlataformaControle.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourcePlataformaControle.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourcePlataformaControle.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryProgramacaoNotasBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceProgramacaoNotas.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceProgramacaoNotas.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceProgramacaoNotas.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryRoteamentoBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

procedure TFrmDataModule.ADOQueryTipoEtapaServicoBeforePost(DataSet: TDataSet);
begin
  FrmDataModule.DataSourceTipoEtapaServico.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceTipoEtapaServico.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourceTipoEtapaServico.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
end;

end.



