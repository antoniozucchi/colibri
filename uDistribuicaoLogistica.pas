unit uDistribuicaoLogistica;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Data.Win.ADODB, System.Generics.Collections,
  System.Generics.Defaults, System.Math, System.DateUtils;

type
  // Estrutura para retorno de validações
  TValidacaoResult = record
    Valido: Boolean;
    Mensagem: string;
    constructor Create(AValido: Boolean; const AMensagem: string);
  end;

  // Estrutura para representar uma rota avaliada pelo algoritmo
  TRotaAvaliada = record
    IdRoteamento: Integer;
    NomeRota: string;
    NomeEmbarcacao: string;
    HoraPartida: TTime;
    ScoreTotal: Double;
    TempoViagemMinutos: Integer;
    OcupacaoPercentual: Double;
    MotivoRejeicao: string;
    Valida: Boolean;
  end;

  // Classe principal de regras de negócio para distribuição logística
  TDistribuicaoLogistica = class
  private
    FConnColibri: TADOConnection;
    FConnConsulta: TADOConnection;
    
    // Métodos auxiliares internos
    function ObterCapacidadeEmbarcacao(const ANomeEmbarcacao: string): Integer;
    function ObterOcupacaoAtualRota(AIdRoteamento: Integer): Integer;
    function ObterSequenciaRota(AIdRoteamento: Integer): TStringList;
    function CalcularTempoViagem(AIdRoteamento: Integer; const AOrigem, ADestino: string): Integer;
    
  public
    constructor Create(AConnColibri, AConnConsulta: TADOConnection);
    destructor Destroy; override;

    // 1. Validações de Capacidade
    function ValidarCapacidadeEmbarcacao(AIdRoteamento: Integer; ANovosPassageiros: Integer = 1): TValidacaoResult;
    
    // 2. Validações de Rota (Origem/Destino)
    function ValidarDestinoNaRota(AIdRoteamento: Integer; const ADestinoExecutante: string): TValidacaoResult;
    function ValidarOrigemNaRota(AIdRoteamento: Integer; const AOrigemExecutante: string): TValidacaoResult;
    
    // 3. Operações de Vinculação
    function VincularExecutanteARota(AIdExecutante: Integer; AIdRoteamento: Integer; out AMensagemErro: string): Boolean;
    function DesvincularExecutanteDeRota(AIdAuxRotaDistribuicao: Integer; out AMensagemErro: string): Boolean;
    
    // 4. Otimização e IA
    function CalcularOcupacaoPercentual(AIdRoteamento: Integer): Double;
    function AvaliarRotasParaExecutante(AIdExecutante: Integer; const ADataProgramacao: string): TArray<TRotaAvaliada>;
    function SugerirMelhorRotaParaExecutante(AIdExecutante: Integer; const ADataProgramacao: string): Integer;
    function AlocarLoteAutomaticamente(ADataProgramacao: string; out AMensagem: string): Integer;
  end;

implementation

{ TValidacaoResult }

constructor TValidacaoResult.Create(AValido: Boolean; const AMensagem: string);
begin
  Valido := AValido;
  Mensagem := AMensagem;
end;

{ TDistribuicaoLogistica }

constructor TDistribuicaoLogistica.Create(AConnColibri, AConnConsulta: TADOConnection);
begin
  FConnColibri := AConnColibri;
  FConnConsulta := AConnConsulta;
end;

destructor TDistribuicaoLogistica.Destroy;
begin
  inherited;
end;

// =============================================================================
// MÉTODOS AUXILIARES
// =============================================================================

function TDistribuicaoLogistica.ObterCapacidadeEmbarcacao(const ANomeEmbarcacao: string): Integer;
var
  Qry: TADOQuery;
begin
  Result := 0;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text := 'SELECT CapacidadePAX FROM tblEmbarcacao WHERE NomeEmbarcacao = :Nome';
    Qry.Parameters.ParamByName('Nome').Value := ANomeEmbarcacao;
    Qry.Open;
    
    if not Qry.IsEmpty then
      Result := Qry.FieldByName('CapacidadePAX').AsInteger;
  finally
    Qry.Free;
  end;
end;

function TDistribuicaoLogistica.ObterOcupacaoAtualRota(AIdRoteamento: Integer): Integer;
var
  Qry: TADOQuery;
begin
  Result := 0;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text := 'SELECT COUNT(*) AS Total FROM tblAux_Rota_Distribuicao WHERE CodigoRota = :IdRota';
    Qry.Parameters.ParamByName('IdRota').Value := AIdRoteamento;
    Qry.Open;
    
    if not Qry.IsEmpty then
      Result := Qry.FieldByName('Total').AsInteger;
  finally
    Qry.Free;
  end;
end;

function TDistribuicaoLogistica.ObterSequenciaRota(AIdRoteamento: Integer): TStringList;
var
  Qry: TADOQuery;
  SequenciaStr: string;
begin
  Result := TStringList.Create;
  Result.Delimiter := ';';
  Result.StrictDelimiter := True;
  
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text := 'SELECT RotaSequencia FROM tblRoteamento WHERE idRoteamento = :IdRota';
    Qry.Parameters.ParamByName('IdRota').Value := AIdRoteamento;
    Qry.Open;
    
    if not Qry.IsEmpty then
    begin
      SequenciaStr := Qry.FieldByName('RotaSequencia').AsString;
      Result.DelimitedText := SequenciaStr;
    end;
  finally
    Qry.Free;
  end;
end;

function TDistribuicaoLogistica.CalcularTempoViagem(AIdRoteamento: Integer; const AOrigem, ADestino: string): Integer;
var
  Sequencia: TStringList;
  I, PosOrigem, PosDestino: Integer;
  QryRota: TADOQuery;
  TempoTransbordo: Integer;
begin
  // Estimativa simplificada: 15 min por trecho + tempo de transbordo
  // Em uma implementação real, buscaria as distâncias reais entre as plataformas
  Result := 9999; // Valor alto para rotas inválidas
  
  Sequencia := ObterSequenciaRota(AIdRoteamento);
  QryRota := TADOQuery.Create(nil);
  try
    QryRota.Connection := FConnColibri;
    QryRota.SQL.Text := 'SELECT TempoTransbordo FROM tblRoteamento WHERE idRoteamento = :IdRota';
    QryRota.Parameters.ParamByName('IdRota').Value := AIdRoteamento;
    QryRota.Open;
    
    TempoTransbordo := 0;
    if not QryRota.IsEmpty then
      TempoTransbordo := QryRota.FieldByName('TempoTransbordo').AsInteger;
      
    PosOrigem := -1;
    PosDestino := -1;
    
    // Encontrar posições na sequência
    for I := 0 to Sequencia.Count - 1 do
    begin
      if SameText(Trim(Sequencia[I]), Trim(AOrigem)) then PosOrigem := I;
      if SameText(Trim(Sequencia[I]), Trim(ADestino)) then PosDestino := I;
    end;
    
    // Se a origem não está na sequência, assumimos que é o ponto de partida (índice -1)
    if PosOrigem = -1 then PosOrigem := -1;
    
    if PosDestino > PosOrigem then
    begin
      // Calcula tempo: (número de trechos * 15 min) + (paradas intermediárias * tempo transbordo)
      Result := ((PosDestino - PosOrigem) * 15) + (Max(0, PosDestino - PosOrigem - 1) * TempoTransbordo);
    end;
  finally
    Sequencia.Free;
    QryRota.Free;
  end;
end;

// =============================================================================
// VALIDAÇÕES DE CAPACIDADE
// =============================================================================

function TDistribuicaoLogistica.ValidarCapacidadeEmbarcacao(AIdRoteamento: Integer; ANovosPassageiros: Integer = 1): TValidacaoResult;
var
  QryRota: TADOQuery;
  NomeEmbarcacao: string;
  CapacidadeMaxima, OcupacaoAtual, NovaOcupacao: Integer;
begin
  QryRota := TADOQuery.Create(nil);
  try
    QryRota.Connection := FConnColibri;
    QryRota.SQL.Text := 'SELECT NomeEmbarcacao FROM tblRoteamento WHERE idRoteamento = :IdRota';
    QryRota.Parameters.ParamByName('IdRota').Value := AIdRoteamento;
    QryRota.Open;
    
    if QryRota.IsEmpty then
      Exit(TValidacaoResult.Create(False, 'Rota não encontrada.'));
      
    NomeEmbarcacao := QryRota.FieldByName('NomeEmbarcacao').AsString;
  finally
    QryRota.Free;
  end;

  CapacidadeMaxima := ObterCapacidadeEmbarcacao(NomeEmbarcacao);
  if CapacidadeMaxima <= 0 then
    Exit(TValidacaoResult.Create(True, 'Aviso: Capacidade da embarcação não definida no cadastro.'));

  OcupacaoAtual := ObterOcupacaoAtualRota(AIdRoteamento);
  NovaOcupacao := OcupacaoAtual + ANovosPassageiros;

  if NovaOcupacao > CapacidadeMaxima then
    Result := TValidacaoResult.Create(False, 
      Format('Capacidade excedida! A embarcação %s suporta %d passageiros. Ocupação atual: %d. Tentando adicionar: %d.', 
      [NomeEmbarcacao, CapacidadeMaxima, OcupacaoAtual, ANovosPassageiros]))
  else
    Result := TValidacaoResult.Create(True, 
      Format('Capacidade OK. Ocupação passará para %d/%d.', [NovaOcupacao, CapacidadeMaxima]));
end;

// =============================================================================
// VALIDAÇÕES DE ROTA (ORIGEM/DESTINO)
// =============================================================================

function TDistribuicaoLogistica.ValidarDestinoNaRota(AIdRoteamento: Integer; const ADestinoExecutante: string): TValidacaoResult;
var
  Sequencia: TStringList;
  I: Integer;
  Encontrou: Boolean;
begin
  if Trim(ADestinoExecutante) = '' then
    Exit(TValidacaoResult.Create(False, 'Destino do executante não informado.'));

  Sequencia := ObterSequenciaRota(AIdRoteamento);
  try
    if Sequencia.Count = 0 then
      Exit(TValidacaoResult.Create(False, 'A rota não possui uma sequência de plataformas definida.'));

    Encontrou := False;
    for I := 0 to Sequencia.Count - 1 do
    begin
      if SameText(Trim(Sequencia[I]), Trim(ADestinoExecutante)) then
      begin
        Encontrou := True;
        Break;
      end;
    end;

    if Encontrou then
      Result := TValidacaoResult.Create(True, 'Destino compatível com a rota.')
    else
      Result := TValidacaoResult.Create(False, 
        Format('O destino %s não faz parte da sequência desta rota.', [ADestinoExecutante]));
  finally
    Sequencia.Free;
  end;
end;


function TDistribuicaoLogistica.ValidarOrigemNaRota(AIdRoteamento: Integer;
  const AOrigemExecutante: string): TValidacaoResult;
var
  Sequencia: TStringList;
begin
  if Trim(AOrigemExecutante) = '' then
    Exit(TValidacaoResult.Create(False, 'Origem do executante não informada.'));

  Sequencia := ObterSequenciaRota(AIdRoteamento);
  try
    if Sequencia.Count = 0 then
      Exit(TValidacaoResult.Create(False, 'A rota não possui sequência definida.'));

    if SameText(Trim(Sequencia[0]), Trim(AOrigemExecutante)) then
      Result := TValidacaoResult.Create(True, 'Origem compatível com a rota.')
    else
      Result := TValidacaoResult.Create(False,
        Format('Origem %s incompatível com a rota. Origem da rota: %s.',
          [AOrigemExecutante, Trim(Sequencia[0])]));
  finally
    Sequencia.Free;
  end;
end;

// =============================================================================
// OPERAÇÕES DE VINCULAÇÃO
// =============================================================================

function TDistribuicaoLogistica.VincularExecutanteARota(AIdExecutante: Integer;
  AIdRoteamento: Integer; out AMensagemErro: string): Boolean;
var
  Qry, QryExec, QryDup: TADOQuery;
  Origem, Destino: string;
  ValidacaoCapacidade, ValidacaoOrigem, ValidacaoDestino: TValidacaoResult;
begin
  Result := False;
  AMensagemErro := '';

  QryDup := TADOQuery.Create(nil);
  QryExec := TADOQuery.Create(nil);
  Qry := TADOQuery.Create(nil);
  try
    // 1. Evita vínculo duplicado
    QryDup.Connection := FConnColibri;
    QryDup.SQL.Text :=
      'SELECT COUNT(*) AS Total ' +
      'FROM tblAux_Rota_Distribuicao ' +
      'WHERE CodigoProgramacaoExecutante = :IdExec';
    QryDup.Parameters.ParamByName('IdExec').Value := AIdExecutante;
    QryDup.Open;

    if QryDup.FieldByName('Total').AsInteger > 0 then
    begin
      AMensagemErro := 'Executante já possui rota vinculada.';
      Exit;
    end;

    // 2. Busca origem/destino do executante
    QryExec.Connection := FConnColibri;
    QryExec.SQL.Text :=
      'SELECT pe.Origem, pd.txtDestino AS Destino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pe.idProgramacaoExecutante = :IdExec';
    QryExec.Parameters.ParamByName('IdExec').Value := AIdExecutante;
    QryExec.Open;

    if QryExec.IsEmpty then
    begin
      AMensagemErro := 'Executante não encontrado.';
      Exit;
    end;

    Origem := Trim(QryExec.FieldByName('Origem').AsString);
    Destino := Trim(QryExec.FieldByName('Destino').AsString);

    // 3. Validações da rota
    ValidacaoOrigem := ValidarOrigemNaRota(AIdRoteamento, Origem);
    if not ValidacaoOrigem.Valido then
    begin
      AMensagemErro := ValidacaoOrigem.Mensagem;
      Exit;
    end;

    ValidacaoDestino := ValidarDestinoNaRota(AIdRoteamento, Destino);
    if not ValidacaoDestino.Valido then
    begin
      AMensagemErro := ValidacaoDestino.Mensagem;
      Exit;
    end;

    ValidacaoCapacidade := ValidarCapacidadeEmbarcacao(AIdRoteamento, 1);
    if not ValidacaoCapacidade.Valido then
    begin
      AMensagemErro := ValidacaoCapacidade.Mensagem;
      Exit;
    end;
    //------------------------------------------------
    try
      // 4. Inserção
      Qry.Connection := FConnColibri;
      Qry.SQL.Text :=
        'INSERT INTO tblAux_Rota_Distribuicao (CodigoProgramacaoExecutante, CodigoRota) ' +
        'VALUES (:IdExecutante, :IdRota)';
      Qry.Parameters.ParamByName('IdExecutante').Value := AIdExecutante;
      Qry.Parameters.ParamByName('IdRota').Value := AIdRoteamento;

      Qry.ExecSQL;
      Result := True;
    except
      on E: Exception do
        AMensagemErro := 'Erro ao vincular executante: ' + E.Message;
    end;
  finally
    QryDup.Free;
    QryExec.Free;
    Qry.Free;
  end;
end;

function TDistribuicaoLogistica.DesvincularExecutanteDeRota(AIdAuxRotaDistribuicao: Integer; out AMensagemErro: string): Boolean;
var
  Qry: TADOQuery;
begin
  Result := False;
  AMensagemErro := '';

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text := 'DELETE FROM tblAux_Rota_Distribuicao WHERE idAux_Rota_Distribuicao = :IdAux';
    Qry.Parameters.ParamByName('IdAux').Value := AIdAuxRotaDistribuicao;
    
    try
      Qry.ExecSQL;
      Result := True;
    except
      on E: Exception do
        AMensagemErro := 'Erro ao desvincular executante: ' + E.Message;
    end;
  finally
    Qry.Free;
  end;
end;

// =============================================================================
// OTIMIZAÇÃO E IA
// =============================================================================

function TDistribuicaoLogistica.CalcularOcupacaoPercentual(AIdRoteamento: Integer): Double;
var
  QryRota: TADOQuery;
  NomeEmbarcacao: string;
  CapacidadeMaxima, OcupacaoAtual: Integer;
begin
  Result := 0.0;
  
  QryRota := TADOQuery.Create(nil);
  try
    QryRota.Connection := FConnColibri;
    QryRota.SQL.Text := 'SELECT NomeEmbarcacao FROM tblRoteamento WHERE idRoteamento = :IdRota';
    QryRota.Parameters.ParamByName('IdRota').Value := AIdRoteamento;
    QryRota.Open;
    
    if QryRota.IsEmpty then Exit;
    NomeEmbarcacao := QryRota.FieldByName('NomeEmbarcacao').AsString;
  finally
    QryRota.Free;
  end;

  CapacidadeMaxima := ObterCapacidadeEmbarcacao(NomeEmbarcacao);
  if CapacidadeMaxima <= 0 then Exit;

  OcupacaoAtual := ObterOcupacaoAtualRota(AIdRoteamento);
  
  Result := (OcupacaoAtual / CapacidadeMaxima) * 100.0;
end;

function TDistribuicaoLogistica.AvaliarRotasParaExecutante(AIdExecutante: Integer; const ADataProgramacao: string): TArray<TRotaAvaliada>;
var
  QryExec, QryRotas: TADOQuery;
  Origem, Destino: string;
  ListaRotas: TList<TRotaAvaliada>;
  Rota: TRotaAvaliada;
  ValidacaoDestino, ValidacaoCapacidade: TValidacaoResult;
  ScoreTempo, ScoreOcupacao, ScorePermanencia: Double;
  HoraPartidaStr: string;
  ValidacaoOrigem: TValidacaoResult;
begin
  ListaRotas := TList<TRotaAvaliada>.Create;
  QryExec := TADOQuery.Create(nil);
  QryRotas := TADOQuery.Create(nil);
  try
    // 1. Obter dados do executante (Origem de tblProgramacaoExecutante, Destino de tblProgramacaoDiaria)
    QryExec.Connection := FConnColibri;
    QryExec.SQL.Text := 
      'SELECT pe.Origem, pd.txtDestino AS Destino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pe.idProgramacaoExecutante = :IdExec';
    QryExec.Parameters.ParamByName('IdExec').Value := AIdExecutante;
    QryExec.Open;
    
    if QryExec.IsEmpty then Exit;
    
    Origem := QryExec.FieldByName('Origem').AsString;
    Destino := QryExec.FieldByName('Destino').AsString;
    
    // 2. Buscar todas as rotas do dia
    QryRotas.Connection := FConnColibri;
    QryRotas.SQL.Text := 'SELECT idRoteamento, NomeRota, NomeEmbarcacao, HoraRoteamento ' +
                         'FROM tblRoteamento WHERE DataRoteamento = :DataProg';
    QryRotas.Parameters.ParamByName('DataProg').Value := ADataProgramacao;
    QryRotas.Open;
    
    while not QryRotas.Eof do
    begin
      Rota.IdRoteamento := QryRotas.FieldByName('idRoteamento').AsInteger;
      Rota.NomeRota := QryRotas.FieldByName('NomeRota').AsString;
      Rota.NomeEmbarcacao := QryRotas.FieldByName('NomeEmbarcacao').AsString;
      
      HoraPartidaStr := QryRotas.FieldByName('HoraRoteamento').AsString;
      if Trim(HoraPartidaStr) <> '' then
        Rota.HoraPartida := StrToTimeDef(HoraPartidaStr, 0)
      else
        Rota.HoraPartida := 0;
        
      Rota.Valida := True;
      Rota.MotivoRejeicao := '';
      
      // 3. Aplicar Restrições (Filtros Hard)
      
      // Restrição 1: Destino na rota
      ValidacaoDestino := ValidarDestinoNaRota(Rota.IdRoteamento, Destino);
      if not ValidacaoDestino.Valido then
      begin
        Rota.Valida := False;
        Rota.MotivoRejeicao := 'Destino fora da rota';
      end;
      
      // Restrição 2: Origem compatível
      if Rota.Valida then
      begin
        ValidacaoOrigem := ValidarOrigemNaRota(Rota.IdRoteamento, Origem);
        if not ValidacaoOrigem.Valido then
        begin
          Rota.Valida := False;
          Rota.MotivoRejeicao := 'Origem incompatível';
        end;
      end;

      // Restrição 3: Capacidade
      if Rota.Valida then
      begin
        ValidacaoCapacidade := ValidarCapacidadeEmbarcacao(Rota.IdRoteamento, 1);
        if not ValidacaoCapacidade.Valido then
        begin
          Rota.Valida := False;
          Rota.MotivoRejeicao := 'Capacidade excedida';
        end;
      end;
      
      // 4. Calcular Scores (se válida)
      if Rota.Valida then
      begin
        Rota.TempoViagemMinutos := CalcularTempoViagem(Rota.IdRoteamento, Origem, Destino);
        Rota.OcupacaoPercentual := CalcularOcupacaoPercentual(Rota.IdRoteamento);
        
        // Algoritmo de Otimização (Opção D: Combinação com pesos)
        // Peso 1: Tempo de viagem (menor é melhor) -> 40%
        ScoreTempo := Max(0, 100 - (Rota.TempoViagemMinutos / 2)); 
        
        // Peso 2: Permanência (partir mais cedo é melhor) -> 40%
        // Assumindo que 05:00 é o ideal (score 100) e cada hora a mais perde pontos
        ScorePermanencia := Max(0, 100 - (HourOf(Rota.HoraPartida) - 5) * 10);
        
        // Peso 3: Balanceamento de carga (menor ocupação é melhor) -> 20%
        ScoreOcupacao := 100 - Rota.OcupacaoPercentual;
        
        // Score Total (0 a 100)
        Rota.ScoreTotal := (ScoreTempo * 0.4) + (ScorePermanencia * 0.4) + (ScoreOcupacao * 0.2);
      end
      else
      begin
        Rota.ScoreTotal := 0;
        Rota.TempoViagemMinutos := 0;
        Rota.OcupacaoPercentual := 0;
      end;
      
      ListaRotas.Add(Rota);
      QryRotas.Next;
    end;
    
    // 5. Ordenar rotas válidas por Score (decrescente) e HoraPartida (crescente - desempate)
    ListaRotas.Sort(TComparer<TRotaAvaliada>.Construct(
      function(const L, R: TRotaAvaliada): Integer
      begin
        // Primeiro critério: Validade
        if L.Valida and not R.Valida then Exit(-1);
        if not L.Valida and R.Valida then Exit(1);
        
        // Segundo critério: Score Total
        if L.ScoreTotal > R.ScoreTotal then Exit(-1);
        if L.ScoreTotal < R.ScoreTotal then Exit(1);
        
        // Terceiro critério (Desempate): Hora de Partida (mais cedo é melhor)
        if L.HoraPartida < R.HoraPartida then Exit(-1);
        if L.HoraPartida > R.HoraPartida then Exit(1);
        
        Result := 0;
      end));
      
    Result := ListaRotas.ToArray;
  finally
    ListaRotas.Free;
    QryExec.Free;
    QryRotas.Free;
  end;
end;

function TDistribuicaoLogistica.SugerirMelhorRotaParaExecutante(AIdExecutante: Integer; const ADataProgramacao: string): Integer;
var
  RotasAvaliadas: TArray<TRotaAvaliada>;
begin
  Result := 0;
  RotasAvaliadas := AvaliarRotasParaExecutante(AIdExecutante, ADataProgramacao);
  
  if (Length(RotasAvaliadas) > 0) and RotasAvaliadas[0].Valida then
    Result := RotasAvaliadas[0].IdRoteamento;
end;

function TDistribuicaoLogistica.AlocarLoteAutomaticamente(ADataProgramacao: string; out AMensagem: string): Integer;
var
  QryExec: TADOQuery;
  IdExecutante, MelhorRotaId: Integer;
  Sucessos, Falhas: Integer;
  MsgErro: string;
begin
  Sucessos := 0;
  Falhas := 0;
  
  QryExec := TADOQuery.Create(nil);
  try
    QryExec.Connection := FConnColibri;
    // Buscar executantes não alocados para a data
    QryExec.SQL.Text := 
      'SELECT pe.idProgramacaoExecutante ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      'AND (pe.InseridoProgramacaoTransporte = False OR pe.InseridoProgramacaoTransporte IS NULL)';
    QryExec.Parameters.ParamByName('DataProg').Value := ADataProgramacao;
    QryExec.Open;
    
    while not QryExec.Eof do
    begin
      IdExecutante := QryExec.FieldByName('idProgramacaoExecutante').AsInteger;
      
      // Sugerir melhor rota
      MelhorRotaId := SugerirMelhorRotaParaExecutante(IdExecutante, ADataProgramacao);
      
      if MelhorRotaId > 0 then
      begin
        // Tentar vincular
        if VincularExecutanteARota(IdExecutante, MelhorRotaId, MsgErro) then
        begin
          // Atualizar status
          FConnColibri.Execute(Format('UPDATE tblProgramacaoExecutante SET InseridoProgramacaoTransporte = True WHERE idProgramacaoExecutante = %d', [IdExecutante]));
          Inc(Sucessos);
        end
        else
          Inc(Falhas);
      end
      else
        Inc(Falhas); // Nenhuma rota válida encontrada
        
      QryExec.Next;
    end;
    
    AMensagem := Format('Alocação automática concluída: %d executantes alocados, %d falharam (sem rota compatível ou capacidade excedida).', [Sucessos, Falhas]);
    Result := Sucessos;
  finally
    QryExec.Free;
  end;
end;

end.
