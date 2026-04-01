unit untDataModule;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Data.Win.ADODB,Winapi.Windows,uExcelSapRunner,
  ActiveX, System.Variants, ADOInt, untDBGridFilter, Vcl.ComCtrls, uZucchi, uMensagens;

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
    ADOQueryCarregaExecutanteAPLAT: TADOQuery;
    DataSourceCarregaExecutanteAPLAT: TDataSource;
    DataSourceProgramacaoDiaria_Cadastro: TDataSource;
    ADOQueryProgramacaoDiaria_Cadastro: TADOQuery;
    ADOQueryImportarExecutanteAPLAT: TADOQuery;
    DataSourceImportarExecutanteAPLAT: TDataSource;
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
    ADOConnectionMemoria: TADOConnection;
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
    DataSourceAnalisarTipoEtapaServico: TDataSource;
    ADOQueryAnalisarTipoEtapaServico: TADOQuery;
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
    DataSourceTemporarioDBMemoria: TDataSource;
    ADOQueryTemporarioDBMemoria: TADOQuery;
    ADOConnectionGantt: TADOConnection;
    ADOQueryGantt: TADOQuery;
    DataSourceGantt: TDataSource;
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
    DataSourceImportarMemoria: TDataSource;
    ADOQueryImportarMemoria: TADOQuery;
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
    procedure ADOQueryEmbarcacoesBeforePost(DataSet: TDataSet);
    procedure ADOQueryRoteamentoBeforePost(DataSet: TDataSet);
    procedure ADOQueryGeradoresBeforePost(DataSet: TDataSet);
    procedure ADOQueryExecutanteBeforePost(DataSet: TDataSet);
    procedure ADOQueryTipoEtapaServicoBeforePost(DataSet: TDataSet);
    procedure ADOQueryCadastroUsuarioBeforePost(DataSet: TDataSet);
    procedure ADOQueryImportarExecutanteAPLATBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaAfterRefresh(DataSet: TDataSet);
    procedure ADOQueryMovimentacaoCargaBeforePost(DataSet: TDataSet);
    procedure ADOQueryCondicaoMarBeforePost(DataSet: TDataSet);
    procedure ADOQueryCondicaoEmbarcacaoBeforePost(DataSet: TDataSet);
    procedure ADOQueryProgramacaoNotasBeforePost(DataSet: TDataSet);
    procedure ADOQueryPlataformaControleBeforePost(DataSet: TDataSet);
  private

    { Private declarations }
  public
    naoGravar: Boolean;
    procedure AtualizarProgramacaoExecutanteComRetornoRT(
      const Rows: TArray<TRtR3Row>);
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
  untExecutante,untCondicaoEmbarcacao,untFrmSobre, untFrmMagicFiltro;

{%CLASSGROUP 'Vcl.Controls.TControl'}

{$R *.dfm}

{ TFrmDataModule }

procedure TFrmDataModule.setFilterDBGrid(DBGrid: TFilterDBGrid);
begin
  // 2) aponta o DataSource do grid
  DBGrid.FilterDataSource := DBGrid.DataSource;
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
    StatusBar.Panels[0].Text := 'N° Registros: ' + IntToStr(NumeroLinhas);
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

procedure TFrmDataModule.AtualizarProgramacaoExecutanteComRetornoRT(
  const Rows: TArray<TRtR3Row>);
var
  Q: TADOQuery;
  i: Integer;
  LogRT, RT, Status, Erro: string;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := ADOConnectionColibri;

    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante ' +
      'SET logRT = :logRT, RT = :RT, StatusRT = :StatusRT, ErroRT = :ErroRT ' +
      'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    // (opcional) melhora um pouco performance
    // Q.Prepared := True;

    for i := 0 to High(Rows) do
    begin
      RT     := Rows[i].RtGerada;
      Status := Rows[i].Status;
      Erro   := Rows[i].Erro;

      if Status = '' then
      begin
        Status := 'ERRO';
        if Erro = '' then
          Erro := 'Sem retorno da macro (Excel/SAP).';
      end;

      LogRT := Format('%s | %s', [RT, Status]);

      Q.Parameters.ParamByName('logRT').Value := LogRT;
      Q.Parameters.ParamByName('RT').Value := RT;
      Q.Parameters.ParamByName('StatusRT').Value := Status;
      Q.Parameters.ParamByName('ErroRT').Value := Erro;
      Q.Parameters.ParamByName('idProgramacaoExecutante').Value := Rows[i].IdProgramacaoExecutante;

      Q.ExecSQL;
    end;

  finally
    Q.Free;
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
    CodigoSAP: String;
begin
  if naoGravar = false then
  begin
    if Assigned(FrmExecutante) then
    begin
      if FrmPrincipal.VerificaCPF(FrmDataModule.DataSourceExecutante.DataSet.FieldByName('CPF').AsString) = '' then
      begin
        FrmDataModule.DataSourceExecutante.DataSet.FieldByName('CPF').AsString:= '000.000.000-00';
        FrmExecutante.MaskEditCPF.Text:= '000.000.000-00';
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

procedure TFrmDataModule.ADOQueryImportarExecutanteAPLATBeforePost(
  DataSet: TDataSet);
begin
  //Usuário
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('ImportadoPor').asString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('DataImportacao').AsDateTime:= now;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('ComputadorImportacao').AsString:= FrmPrincipal.logMaquina;
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
begin
  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('DataAtualizacao').AsDateTime:= now;
  FrmDataModule.DataSourcePlataforma.DataSet.
  FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
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

