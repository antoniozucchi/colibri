unit uCriacaoRotas;

interface

uses
  System.SysUtils, System.Classes, Data.DB, Data.Win.ADODB, System.Generics.Collections,
  System.Generics.Defaults, System.Math, System.DateUtils, uDistribuicaoLogistica, Vcl.Dialogs;

type
  // Estrutura para representar um executante pendente
  TExecutantePendente = record
    IdExecutante: Integer;
    Origem: string;
    Destino: string;
  end;

  // Estrutura para representar uma plataforma e suas coordenadas
  TPlataformaCoord = record
    Nome: string;
    Latitude: Double;
    Longitude: Double;
  end;

  TSugestaoNovaRota = record
    Origem: string;
    Executantes: TArray<TExecutantePendente>;
    Destinos: TArray<string>;
    SequenciaOtimizada: string;
    QtdExecutantes: Integer;
    ExecutantesIds: TArray<Integer>;
    HoraPartidaSugerida: TTime;
    NomeRotaSugerido: string;
  end;

  TEstadoEmbarcacao = record
    Nome: string;
    Capacidade: Integer;
    QtRotasDia: Integer;
    DistanciaAcumuladaKm: Double;
    UltimaPosicao: string;
    ProximaHoraLivre: TDateTime;
  end;

  TParObrigatorio = record
    PlataformaA: string;
    PlataformaB: string;
  end;

  // Classe principal para criação automática de rotas
  TProgressoRotasEvent = procedure(ATotal, AAtual: Integer; const AMensagem: string) of object;

  TCriacaoRotas = class
  private
    FConnColibri: TADOConnection;
    FConnConsulta: TADOConnection;
    FCacheCoord: TDictionary<string, TPlataformaCoord>; // Cache de coordenadas
    FCacheHoraSaida: TDictionary<string, TTime>; // Cache de horários
    FCachePrioridade: TDictionary<string, Integer>; // Plataforma -> PrioridadeDistribuicao
    FCachePares: TArray<TParObrigatorio>; // pares obrigatórios ativos
    FHubPlataforma: string; // plataforma hub (booleanHubPrincipal = True)
    FPendentesHub: TList<TExecutantePendente>; // executantes cuja origem é o hub
    FOnProgresso: TProgressoRotasEvent;
    FCacheGrupoFisico: TDictionary<string, string>; // Nome operacional -> NomeSAP
    FCacheModal: TDictionary<string, string>; // Plataforma -> RT_Modal


    function ObterModalPlataforma(const APlataforma: string): string;
    function DeveConsiderarDistribuicaoMaritima(const AOrigem, ADestino: string): Boolean;
    function ObterDataBaseProg(const ADataProg: string): TDateTime;
    function SplitSequencia(const ASequenciaTexto: string): TArray<string>;
    function CalcularDistanciaSequenciaArray(const ASequencia: TArray<string>): Double;
    function CalcularDistanciaSequenciaTexto(const ASequenciaTexto: string): Double;
    function ObterUltimaPosicaoDaSequencia(const ASequenciaTexto: string): string;
    function CalcularCustoReposicionamento(const AUltimaPosicao, ANovaOrigem: string): Double;

    function CalcularDuracaoEstimadaMinPorSequencia(const ASequenciaTexto: string): Integer;
    function CalcularDuracaoEstimadaMin(const ASugestao: TSugestaoNovaRota): Integer;

    function CalcularScoreEmbarcacao(const AEstado: TEstadoEmbarcacao;
      const ASugestao: TSugestaoNovaRota; const ADataProg: string): Double;

    function EscolherMelhorEmbarcacaoParaPendentes(
      var AEstados: TArray<TEstadoEmbarcacao>;
      const APendentes: TList<TExecutantePendente>;
      const AOrigem, ADataProg: string;
      out ASugestaoEscolhida: TSugestaoNovaRota): Integer;

    procedure AtualizarEstadoEmbarcacaoAposRota(var AEstado: TEstadoEmbarcacao;
      const ASugestao: TSugestaoNovaRota; const ADataProg: string);

    procedure ReportarProgresso(AAtual, ATotal: Integer; const AMensagem: string);
    // Métodos auxiliares
    function ObterCoordenadasPlataforma(const ANome: string): TPlataformaCoord;
    function CalcularDistancia(const P1, P2: TPlataformaCoord): Double;
    function ObterPrioridade(const APlataforma: string): Integer;
    procedure CarregarParesObrigatorios;
    procedure CarregarHubPlataforma;
    function InserirParadaHubNaSequencia(const ASeqArray: TArray<string>): TArray<string>;
    function AbsorverExecutantesHub(var ASugestao: TSugestaoNovaRota): Integer;
    function OtimizarSequencia(const AOrigem: string; const ADestinos: TArray<string>): TArray<string>;
    function ObterHoraSaidaOrigem(const AOrigem: string): TTime;
    function ObterProximoNomeRota(const AOrigem: string; const ADataProg: string): string;
    function ObterProximaHoraPartida(const AOrigem: string;
      const ADestinos: TArray<string>; const ADataProg: string): TTime;
    function CarregarTodasAsCoordenadasDeUmaVez: Boolean;
    function ValidarOrigemPlataforma(const AOrigem: string; const AOrigemPlataformasValidas: string): Boolean;

    function ObterCapacidadeEmbarcacao(const ANomeEmbarcacao: string): Integer;
    function CarregarEstadoEmbarcacoes(const ADataProg: string;
      const AEmbarcacoesDisponiveis: TArray<string>): TArray<TEstadoEmbarcacao>;
    function MontarSugestaoParaLote(const AOrigem: string;
      const AExecutantes: TArray<TExecutantePendente>;
      const ADataProg: string): TSugestaoNovaRota;
    function RemoverDuplicadosMantendoOrdem(const AItens: TArray<string>;
      const AOrigem: string = ''): TArray<string>;
    function NormalizarPlataforma(const AValor: string): string;
    function NormalizarSequenciaTexto(const ASequencia: string): string;
    function ObterIndiceDestinoNaSequencia(const ADestino: string;
      const ASequencia: TArray<string>): Integer;
    procedure OrdenarExecutantesPorSequencia(
      var AExecutantes: TArray<TExecutantePendente>; const AOrigem: string);

    function ExtrairDestinosDaSequencia(
      const ASequencia: string): TArray<string>;

    function ExtrairTurnoPlataforma(const APlataforma: string): string;
    function GrupoHorarioDoLote(const ADestinos: TArray<string>): string;
    function ObterGrupoFisico(const APlataforma: string): string;
    function ObterGrupoHorario(const APlataforma: string): string;
    function ContarRotasMesmoGrupoHorario(const ADataProg,
      AGrupoHorario: string): Integer;
    function MontarLoteRespeitandoGrupoHorario(
      const APendentes: TList<TExecutantePendente>;
      ACapacidade: Integer): TArray<TExecutantePendente>;
    function OrigemUsaBridge(const AOrigem: string): Boolean;
    function MontarLoteBridgeMesmoGrupo(const AOrigem: string;
      const APendentes: TList<TExecutantePendente>;
      ACapacidade: Integer): TArray<TExecutantePendente>;
    procedure RemoverLoteDaLista(APendentes: TList<TExecutantePendente>;
      const ALote: TArray<TExecutantePendente>);
    function EhExecutanteBridge(const AOrigem, ADestino: string): Boolean;
    function ObterNomeEmbarcacaoBridge(const AOrigem: string): string;
    function ObterHoraBridgePorTurno(const ATurno: string): TTime;
    function CarregarRotasExclusivasDoDia(
      const ADataProg: string): TDictionary<Integer, Byte>;
    function CriarOuObterRotaBridge(const ADataProg, AOrigem, ADestino,
      ANomeEmbarcacaoBridge: string;
      const ARotasExclusivas: TDictionary<Integer, Byte>;
      out AIdRota: Integer;
      out AMensagemErro: string): Boolean;
    function ProcessarRotasBridgeDoDia(const ADataProg: string;
      DistLogistica: TDistribuicaoLogistica;
      const ARotasExclusivas: TDictionary<Integer, Byte>;
      out AMensagem: string;
      var ASucessos, AFalhas, ARotasCriadas: Integer): Boolean;

  public
    constructor Create(AConnColibri, AConnConsulta: TADOConnection);
    destructor Destroy; override;

    property OnProgresso: TProgressoRotasEvent read FOnProgresso write FOnProgresso;

    // 1. Análise e Agrupamento
    function AnalisarExecutantesPendentes(const ADataProg: string;
      const AOrigemPlataformasValidas: string): TArray<TSugestaoNovaRota>;

    
    // 2. Criação de Rota
    function CriarRota(const ASugestao: TSugestaoNovaRota; const ADataProg: string; 
                       const ANomeEmbarcacao: string; out AMensagemErro: string): Integer;
                       
    // 3. Processo Completo (Batch)
    function CriarEAlocarLote(const ADataProg: string; const AEmbarcacoesDisponiveis: TArray<string>;
                              const AOrigemPlataformasValidas: string; out AMensagem: string): Boolean;

  end;

implementation

{ TCriacaoRotas }

function TCriacaoRotas.EhExecutanteBridge(const AOrigem, ADestino: string): Boolean;
var
  TurnoDestino: string;
begin
  Result := False;

  if not OrigemUsaBridge(AOrigem) then
    Exit;

  if not SameText(ObterGrupoFisico(AOrigem), ObterGrupoFisico(ADestino)) then
    Exit;

  TurnoDestino := ExtrairTurnoPlataforma(ADestino);
  Result := (TurnoDestino = 'D') or (TurnoDestino = 'N');
end;

function TCriacaoRotas.ObterNomeEmbarcacaoBridge(const AOrigem: string): string;
var
  Qry: TADOQuery;
begin
  Result := '';

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;

    // 1) prioridade: "<Origem> BRIDGE"
    Qry.SQL.Text :=
      'SELECT TOP 1 NomeEmbarcacao ' +
      'FROM tblEmbarcacao ' +
      'WHERE UsaBridgeMesmoGrupo = True ' +
      '  AND NomeEmbarcacao = :Nome';
    Qry.Parameters.ParamByName('Nome').Value := Trim(AOrigem) + ' BRIDGE';
    Qry.Open;

    if not Qry.IsEmpty then
      Exit(Trim(Qry.FieldByName('NomeEmbarcacao').AsString));

    // 2) fallback: própria origem marcada
    Qry.Close;
    Qry.SQL.Text :=
      'SELECT TOP 1 NomeEmbarcacao ' +
      'FROM tblEmbarcacao ' +
      'WHERE UsaBridgeMesmoGrupo = True ' +
      '  AND NomeEmbarcacao = :Nome';
    Qry.Parameters.ParamByName('Nome').Value := Trim(AOrigem);
    Qry.Open;

    if not Qry.IsEmpty then
      Result := Trim(Qry.FieldByName('NomeEmbarcacao').AsString);
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.ObterHoraBridgePorTurno(const ATurno: string): TTime;
begin
  if SameText(ATurno, 'N') then
    Result := EncodeTime(18, 0, 0, 0)
  else
    Result := EncodeTime(6, 0, 0, 0);
end;

function TCriacaoRotas.CarregarRotasExclusivasDoDia(
  const ADataProg: string): TDictionary<Integer, Byte>;
var
  Qry: TADOQuery;
  IdRota: Integer;
begin
  Result := TDictionary<Integer, Byte>.Create;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT DISTINCT r.idRoteamento ' +
      'FROM tblRoteamento r ' +
      'INNER JOIN tblAux_Rota_Distribuicao ard ON ard.CodigoRota = r.idRoteamento ' +
      'WHERE r.DataRoteamento = :DataProg';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Open;

    while not Qry.Eof do
    begin
      IdRota := Qry.FieldByName('idRoteamento').AsInteger;
      Result.AddOrSetValue(IdRota, 0);
      Qry.Next;
    end;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.CriarOuObterRotaBridge(const ADataProg, AOrigem, ADestino,
  ANomeEmbarcacaoBridge: string;
  const ARotasExclusivas: TDictionary<Integer, Byte>;
  out AIdRota: Integer; out AMensagemErro: string): Boolean;
var
  Qry: TADOQuery;
  DescricaoRota, Sequencia: string;
  HoraBridge: TTime;
  Turno: string;
  Capacidade: Integer;
begin
  Result := False;
  AIdRota := 0;
  AMensagemErro := '';

  Turno := ExtrairTurnoPlataforma(ADestino);
  if SameText(Turno, 'N') then
    DescricaoRota := 'SOV - NOITE'
  else
    DescricaoRota := 'SOV - DIA';

  HoraBridge := ObterHoraBridgePorTurno(Turno);
  Sequencia := AOrigem + ';' + ADestino;
  Capacidade := ObterCapacidadeEmbarcacao(ANomeEmbarcacaoBridge);

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;

    // tenta encontrar rota já criada no dia
    Qry.SQL.Text :=
      'SELECT idRoteamento ' +
      'FROM tblRoteamento ' +
      'WHERE DataRoteamento = :DataProg ' +
      '  AND NomeEmbarcacao = :Embarcacao ' +
      '  AND NomeRota = :NomeRota ' +
      'ORDER BY idRoteamento';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Parameters.ParamByName('Embarcacao').Value := ANomeEmbarcacaoBridge;
    Qry.Parameters.ParamByName('NomeRota').Value := DescricaoRota;
    Qry.Open;

    while not Qry.Eof do
    begin
      AIdRota := Qry.FieldByName('idRoteamento').AsInteger;
      if (ARotasExclusivas = nil) or
         (not ARotasExclusivas.ContainsKey(AIdRota)) then
      begin
        Result := True;
        Exit;
      end;
      Qry.Next;
    end;

    Qry.Close;
    Qry.SQL.Text :=
      'INSERT INTO tblRoteamento ' +
      '  (NomeRota, NomeEmbarcacao, DataRoteamento, HoraRoteamento, RotaSequencia, CapacidadePAX) ' +
      'VALUES ' +
      '  (:NomeRota, :Embarcacao, :DataProg, :Hora, :Sequencia, :Capacidade)';
    Qry.Parameters.ParamByName('NomeRota').Value := DescricaoRota;
    Qry.Parameters.ParamByName('Embarcacao').Value := ANomeEmbarcacaoBridge;
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Parameters.ParamByName('Hora').Value := FormatDateTime('hh:nn', HoraBridge);
    Qry.Parameters.ParamByName('Sequencia').Value := Sequencia;
    Qry.Parameters.ParamByName('Capacidade').Value := Capacidade;
    Qry.ExecSQL;
    try
      Qry.Close;
      Qry.SQL.Text := 'SELECT @@IDENTITY AS NovoID';
      Qry.Open;

      AIdRota := Qry.FieldByName('NovoID').AsInteger;
      Result := AIdRota > 0;
    except
      on E: Exception do
        AMensagemErro := E.Message;
    end;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.ProcessarRotasBridgeDoDia(const ADataProg: string;
  DistLogistica: TDistribuicaoLogistica;
  const ARotasExclusivas: TDictionary<Integer, Byte>;
  out AMensagem: string;
  var ASucessos, AFalhas, ARotasCriadas: Integer): Boolean;
var
  Qry: TADOQuery;
  IdExec, IdRota: Integer;
  Origem, Destino, NomeBridge, MsgErro: string;
  RotasCriadasNoMetodo: TDictionary<string, Integer>;
  ChaveRota: string;
begin
  Result := True;
  AMensagem := '';

  RotasCriadasNoMetodo := TDictionary<string, Integer>.Create;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT pe.idProgramacaoExecutante, pe.Origem, pd.txtDestino AS Destino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      '  AND (pe.InseridoProgramacaoTransporte = False OR pe.InseridoProgramacaoTransporte IS NULL) ' +
      '  AND NOT EXISTS (' +
      '    SELECT 1 ' +
      '    FROM tblAux_Rota_Distribuicao ard ' +
      '    WHERE ard.CodigoProgramacaoExecutante = pe.idProgramacaoExecutante' +
      '  )';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Open;

    while not Qry.Eof do
    begin
      IdExec := Qry.FieldByName('idProgramacaoExecutante').AsInteger;
      Origem := Trim(Qry.FieldByName('Origem').AsString);
      Destino := Trim(Qry.FieldByName('Destino').AsString);

      if not DeveConsiderarDistribuicaoMaritima(Origem, Destino) then
      begin
        Qry.Next;
        Continue;
      end;

      if EhExecutanteBridge(Origem, Destino) then
      begin
        NomeBridge := ObterNomeEmbarcacaoBridge(Origem);

        if NomeBridge = '' then
        begin
          Inc(AFalhas);
          AMensagem := AMensagem +
            Format('Sem embarcação BRIDGE configurada para a origem %s.', [Origem]) + sLineBreak;
          Qry.Next;
          Continue;
        end;

        ChaveRota := Origem + '|' + Destino;

        if not RotasCriadasNoMetodo.TryGetValue(ChaveRota, IdRota) then
        begin
          if CriarOuObterRotaBridge(
               ADataProg,
               Origem,
               Destino,
               NomeBridge,
               ARotasExclusivas,
               IdRota,
               MsgErro
             ) then
          begin
            RotasCriadasNoMetodo.AddOrSetValue(ChaveRota, IdRota);
            Inc(ARotasCriadas);
          end
          else
          begin
            Inc(AFalhas);
            AMensagem := AMensagem +
              Format('Falha ao criar rota BRIDGE %s -> %s: %s', [Origem, Destino, MsgErro]) + sLineBreak;
            Qry.Next;
            Continue;
          end;
        end;

        if DistLogistica.VincularExecutanteARota(IdExec, IdRota, MsgErro) then
        begin
          FConnColibri.Execute(
            Format(
              'UPDATE tblProgramacaoExecutante ' +
              'SET InseridoProgramacaoTransporte = True ' +
              'WHERE idProgramacaoExecutante = %d',
              [IdExec]
            )
          );
          Inc(ASucessos);
        end
        else
        begin
          Inc(AFalhas);
          AMensagem := AMensagem +
            Format('Falha ao vincular executante %d na rota BRIDGE %d: %s',
              [IdExec, IdRota, MsgErro]) + sLineBreak;
        end;
      end;

      Qry.Next;
    end;
  finally
    RotasCriadasNoMetodo.Free;
    Qry.Free;
  end;
end;

function TCriacaoRotas.CalcularDuracaoEstimadaMin(
  const ASugestao: TSugestaoNovaRota): Integer;
begin
  Result := CalcularDuracaoEstimadaMinPorSequencia(ASugestao.SequenciaOtimizada);
end;

function TCriacaoRotas.CalcularDuracaoEstimadaMinPorSequencia(
  const ASequenciaTexto: string): Integer;
var
  DistKm: Double;
begin
  DistKm := CalcularDistanciaSequenciaTexto(ASequenciaTexto);

  // ajuste fino depois conforme sua operação
  Result := Round((DistKm / 18) * 60) + 15;

  if Result < 15 then
    Result := 15;
end;

function TCriacaoRotas.SplitSequencia(const ASequenciaTexto: string): TArray<string>;
var
  SL: TStringList;
begin
  SL := TStringList.Create;
  try
    SL.StrictDelimiter := True;
    SL.Delimiter := ';';
    SL.DelimitedText := ASequenciaTexto;
    Result := RemoverDuplicadosMantendoOrdem(SL.ToStringArray);
  finally
    SL.Free;
  end;
end;

function TCriacaoRotas.CalcularDistanciaSequenciaArray(
  const ASequencia: TArray<string>): Double;
var
  I: Integer;
  P1, P2: TPlataformaCoord;
  A, B: string;
begin
  Result := 0;

  if Length(ASequencia) < 2 then
    Exit;

  for I := 0 to High(ASequencia) - 1 do
  begin
    A := NormalizarPlataforma(ASequencia[I]);
    B := NormalizarPlataforma(ASequencia[I + 1]);

    if (A = '') or (B = '') then
      Continue;

    if SameText(A, B) then
      Continue;

    P1 := ObterCoordenadasPlataforma(A);
    P2 := ObterCoordenadasPlataforma(B);
    Result := Result + CalcularDistancia(P1, P2);
  end;
end;

function TCriacaoRotas.CalcularDistanciaSequenciaTexto(
  const ASequenciaTexto: string): Double;
begin
  Result := CalcularDistanciaSequenciaArray(SplitSequencia(ASequenciaTexto));
end;

function TCriacaoRotas.ObterUltimaPosicaoDaSequencia(
  const ASequenciaTexto: string): string;
var
  Arr: TArray<string>;
begin
  Result := '';
  Arr := SplitSequencia(ASequenciaTexto);
  if Length(Arr) > 0 then
    Result := NormalizarPlataforma(Arr[High(Arr)]);
end;

function TCriacaoRotas.CalcularCustoReposicionamento(
  const AUltimaPosicao, ANovaOrigem: string): Double;
var
  P1, P2: TPlataformaCoord;
  UltPos, NovaOrigem: string;
begin
  Result := 0;

  UltPos := NormalizarPlataforma(AUltimaPosicao);
  NovaOrigem := NormalizarPlataforma(ANovaOrigem);

  if (UltPos = '') or (NovaOrigem = '') then
    Exit;

  if SameText(UltPos, NovaOrigem) then
    Exit;

  P1 := ObterCoordenadasPlataforma(UltPos);
  P2 := ObterCoordenadasPlataforma(NovaOrigem);
  Result := CalcularDistancia(P1, P2);
end;

function TCriacaoRotas.CalcularScoreEmbarcacao(
  const AEstado: TEstadoEmbarcacao;
  const ASugestao: TSugestaoNovaRota; const ADataProg: string): Double;
var
  DistNovaRota, DistReposicionamento, Aproveitamento: Double;
  EsperaMin: Double;
  HoraDesejada, DataBase: TDateTime;
begin
  DistNovaRota := CalcularDistanciaSequenciaTexto(ASugestao.SequenciaOtimizada);
  DistReposicionamento := CalcularCustoReposicionamento(
    AEstado.UltimaPosicao,
    ASugestao.Origem
  );

  if AEstado.Capacidade > 0 then
    Aproveitamento := ASugestao.QtdExecutantes / AEstado.Capacidade
  else
    Aproveitamento := 0;

  DataBase := ObterDataBaseProg(ADataProg);
  HoraDesejada := DataBase + ASugestao.HoraPartidaSugerida;

  if AEstado.ProximaHoraLivre > HoraDesejada then
    EsperaMin := (AEstado.ProximaHoraLivre - HoraDesejada) * 1440
  else
    EsperaMin := 0;

  Result :=
      (AEstado.DistanciaAcumuladaKm * 1.0)
    + (DistReposicionamento * 1.5)
    + (DistNovaRota * 1.0)
    + (AEstado.QtRotasDia * 10.0)
    + (EsperaMin * 0.8)
    - (Aproveitamento * 8.0);
end;

function TCriacaoRotas.EscolherMelhorEmbarcacaoParaPendentes(
  var AEstados: TArray<TEstadoEmbarcacao>;
  const APendentes: TList<TExecutantePendente>;
  const AOrigem, ADataProg: string;
  out ASugestaoEscolhida: TSugestaoNovaRota): Integer;
var
  I, QtdCandidata: Integer;
  LoteTemp: TArray<TExecutantePendente>;
  SugTemp: TSugestaoNovaRota;
  ScoreAtual, MelhorScore: Double;
begin
  Result := -1;
  MelhorScore := MaxDouble;

  if APendentes.Count = 0 then
    Exit;

  for I := 0 to High(AEstados) do
  begin
    QtdCandidata := AEstados[I].Capacidade;
    if QtdCandidata <= 0 then
      QtdCandidata := APendentes.Count;

    if QtdCandidata > APendentes.Count then
      QtdCandidata := APendentes.Count;

    if QtdCandidata <= 0 then
      Continue;

    LoteTemp := MontarLoteRespeitandoGrupoHorario(APendentes, QtdCandidata);
    if Length(LoteTemp) = 0 then
      Continue;

    SugTemp := MontarSugestaoParaLote(AOrigem, LoteTemp, ADataProg);
    ScoreAtual := CalcularScoreEmbarcacao(AEstados[I], SugTemp, ADataProg);

    if (Result = -1) or (ScoreAtual < MelhorScore) then
    begin
      MelhorScore := ScoreAtual;
      Result := I;
      ASugestaoEscolhida := SugTemp;
    end;
  end;
end;

procedure TCriacaoRotas.AtualizarEstadoEmbarcacaoAposRota(
  var AEstado: TEstadoEmbarcacao;
  const ASugestao: TSugestaoNovaRota; const ADataProg: string);
var
  HoraInicio, DataBase, NovaHoraLivre: TDateTime;
  DuracaoMin: Integer;
begin
  Inc(AEstado.QtRotasDia);

  AEstado.DistanciaAcumuladaKm :=
    AEstado.DistanciaAcumuladaKm +
    CalcularDistanciaSequenciaTexto(ASugestao.SequenciaOtimizada);

  AEstado.UltimaPosicao :=
    ObterUltimaPosicaoDaSequencia(ASugestao.SequenciaOtimizada);

  DataBase := ObterDataBaseProg(ADataProg);
  HoraInicio := DataBase + ASugestao.HoraPartidaSugerida;
  DuracaoMin := CalcularDuracaoEstimadaMin(ASugestao);

  NovaHoraLivre := HoraInicio + (DuracaoMin / 1440);

  if NovaHoraLivre > AEstado.ProximaHoraLivre then
    AEstado.ProximaHoraLivre := NovaHoraLivre;
end;

function TCriacaoRotas.NormalizarPlataforma(const AValor: string): string;
begin
  Result := Trim(AValor);
  while Pos('  ', Result) > 0 do
    Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;

function TCriacaoRotas.RemoverDuplicadosMantendoOrdem(
  const AItens: TArray<string>; const AOrigem: string = ''): TArray<string>;
var
  Lista: TList<string>;
  Chaves: TDictionary<string, Byte>;
  Item, ItemNorm, OrigemNorm, Chave: string;
begin
  Lista := TList<string>.Create;
  Chaves := TDictionary<string, Byte>.Create;
  try
    OrigemNorm := AnsiUpperCase(NormalizarPlataforma(AOrigem));

    for Item in AItens do
    begin
      ItemNorm := NormalizarPlataforma(Item);
      if ItemNorm = '' then
        Continue;

      if (OrigemNorm <> '') and (AnsiUpperCase(ItemNorm) = OrigemNorm) and (Lista.Count > 0) then
        Continue;

      Chave := AnsiUpperCase(ItemNorm);

      if not Chaves.ContainsKey(Chave) then
      begin
        Lista.Add(ItemNorm);
        Chaves.Add(Chave, 0);
      end;
    end;

    Result := Lista.ToArray;
  finally
    Lista.Free;
    Chaves.Free;
  end;
end;

procedure TCriacaoRotas.ReportarProgresso(AAtual, ATotal: Integer;
  const AMensagem: string);
begin
  if Assigned(FOnProgresso) then
    FOnProgresso(ATotal, AAtual, AMensagem);
end;

function TCriacaoRotas.ValidarOrigemPlataforma(const AOrigem: string;
  const AOrigemPlataformasValidas: string): Boolean;
var
  Origens: TArray<string>;
  I: Integer;
begin
  Result := False;
  Origens := AOrigemPlataformasValidas.Split([';']);
  for I := 0 to High(Origens) do
  begin
    if Trim(Origens[I]) = Trim(AOrigem) then
    begin
      Result := True;
      Break;
    end;
  end;
end;

//===============================================================
//===============================================================
//===============================================================
function TCriacaoRotas.NormalizarSequenciaTexto(const ASequencia: string): string;
var
  SL: TStringList;
  Arr: TArray<string>;
begin
  SL := TStringList.Create;
  try
    SL.StrictDelimiter := True;
    SL.Delimiter := ';';
    SL.DelimitedText := ASequencia;
    Arr := RemoverDuplicadosMantendoOrdem(SL.ToStringArray);
    Result := string.Join(';', Arr);
  finally
    SL.Free;
  end;
end;

function TCriacaoRotas.ObterCapacidadeEmbarcacao(const ANomeEmbarcacao: string): Integer;
var
  Qry: TADOQuery;
begin
  Result := 0;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT CapacidadePAX ' +
      'FROM tblEmbarcacao ' +
      'WHERE NomeEmbarcacao = :Nome';
    Qry.Parameters.ParamByName('Nome').Value := Trim(ANomeEmbarcacao);
    Qry.Open;

    if not Qry.IsEmpty then
      Result := Qry.FieldByName('CapacidadePAX').AsInteger;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.CarregarEstadoEmbarcacoes(const ADataProg: string;
  const AEmbarcacoesDisponiveis: TArray<string>): TArray<TEstadoEmbarcacao>;
var
  I: Integer;
  Qry: TADOQuery;
  Seq, HoraStr: string;
  DataBase, HoraInicio, HoraFim: TDateTime;
  DuracaoMin: Integer;
begin
  SetLength(Result, Length(AEmbarcacoesDisponiveis));
  DataBase := ObterDataBaseProg(ADataProg);

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT idRoteamento, HoraRoteamento, RotaSequencia ' +
      'FROM tblRoteamento ' +
      'WHERE DataRoteamento = :DataProg ' +
      '  AND NomeEmbarcacao = :Embarcacao ' +
      'ORDER BY idRoteamento';

    for I := 0 to High(AEmbarcacoesDisponiveis) do
    begin
      Result[I].Nome := AEmbarcacoesDisponiveis[I];
      Result[I].Capacidade := ObterCapacidadeEmbarcacao(AEmbarcacoesDisponiveis[I]);
      Result[I].QtRotasDia := 0;
      Result[I].DistanciaAcumuladaKm := 0;
      Result[I].UltimaPosicao := '';
      Result[I].ProximaHoraLivre := DataBase;

      Qry.Close;
      Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
      Qry.Parameters.ParamByName('Embarcacao').Value := AEmbarcacoesDisponiveis[I];
      Qry.Open;

      while not Qry.Eof do
      begin
        Inc(Result[I].QtRotasDia);

        Seq := Trim(Qry.FieldByName('RotaSequencia').AsString);
        Result[I].DistanciaAcumuladaKm :=
          Result[I].DistanciaAcumuladaKm + CalcularDistanciaSequenciaTexto(Seq);

        Result[I].UltimaPosicao := ObterUltimaPosicaoDaSequencia(Seq);

        HoraStr := Trim(Qry.FieldByName('HoraRoteamento').AsString);
        HoraInicio := DataBase + StrToTimeDef(HoraStr, 0);
        DuracaoMin := CalcularDuracaoEstimadaMinPorSequencia(Seq);
        HoraFim := HoraInicio + (DuracaoMin / 1440);

        if HoraFim > Result[I].ProximaHoraLivre then
          Result[I].ProximaHoraLivre := HoraFim;

        Qry.Next;
      end;
    end;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.MontarSugestaoParaLote(const AOrigem: string;
  const AExecutantes: TArray<TExecutantePendente>;
  const ADataProg: string): TSugestaoNovaRota;
var
  I: Integer;
  DestinosUnicos: TStringList;
  SeqArray: TArray<string>;
begin
  Result.Origem := AOrigem;
  Result.Executantes := Copy(AExecutantes);
  Result.QtdExecutantes := Length(AExecutantes);

  SetLength(Result.ExecutantesIds, Result.QtdExecutantes);

  DestinosUnicos := TStringList.Create;
  try
    DestinosUnicos.Sorted := False;
    DestinosUnicos.Duplicates := dupIgnore;

    for I := 0 to High(AExecutantes) do
    begin
      Result.ExecutantesIds[I] := AExecutantes[I].IdExecutante;

      if Trim(AExecutantes[I].Destino) <> '' then
        DestinosUnicos.Add(NormalizarPlataforma(AExecutantes[I].Destino));
    end;

    Result.Destinos := RemoverDuplicadosMantendoOrdem(DestinosUnicos.ToStringArray, AOrigem);
    SeqArray := OtimizarSequencia(AOrigem, Result.Destinos);
    Result.SequenciaOtimizada := string.Join(';', SeqArray);
    Result.NomeRotaSugerido := ObterProximoNomeRota(AOrigem, ADataProg);
    Result.HoraPartidaSugerida :=
      ObterProximaHoraPartida(AOrigem, Result.Destinos, ADataProg);
  finally
    DestinosUnicos.Free;
  end;
end;
//===============================================================
//===============================================================
//===============================================================

// Carrega TODAS as coordenadas de uma vez (muito mais rápido!)
function TCriacaoRotas.CarregarTodasAsCoordenadasDeUmaVez: Boolean;
var
  Qry: TADOQuery;
  Coord: TPlataformaCoord;
  HoraStr: string;
  Hora: TTime;
  NomePlataforma: string;
  GrupoFisico: string;
begin
  Result := False;

  FCacheCoord.Clear;
  FCacheHoraSaida.Clear;
  FCachePrioridade.Clear;
  FCacheGrupoFisico.Clear;
  FCacheModal.Clear;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.CommandTimeout := 10;
    Qry.SQL.Text :=
      'SELECT tblPlataforma.Plataforma, ' +
      '       tblPlataforma.NomeSAP, ' +
      '       tblPlataforma.Latitude, ' +
      '       tblPlataforma.Longitude, ' +
      '       tblPlataforma.HoraSaidaOrigem, ' +
      '       tblPlataforma.RT_Modal, ' +
      '       tblPlataforma.PrioridadeDistribuicao ' +
      'FROM tblPlataforma ' +
      'WHERE (tblPlataforma.Latitude IS NOT NULL) ' +
      '  AND (tblPlataforma.Longitude IS NOT NULL) ' +
      '  AND (Trim(tblPlataforma.Plataforma) <> '''')';

    try
      Qry.Open;

      while not Qry.Eof do
      begin
        NomePlataforma := Trim(Qry.FieldByName('Plataforma').AsString);
        GrupoFisico := Trim(Qry.FieldByName('NomeSAP').AsString);

        // Se não houver NomeSAP, usa a própria plataforma como fallback
        if GrupoFisico = '' then
          GrupoFisico := NomePlataforma;

        // Cache: Plataforma -> NomeSAP
        if (NomePlataforma <> '') and (not FCacheGrupoFisico.ContainsKey(NomePlataforma)) then
          FCacheGrupoFisico.Add(NomePlataforma, GrupoFisico)
        else if NomePlataforma <> '' then
          FCacheGrupoFisico.AddOrSetValue(NomePlataforma, GrupoFisico);

        if NomePlataforma <> '' then
          FCacheModal.AddOrSetValue(NomePlataforma, UpperCase(Trim(Qry.FieldByName('RT_Modal').AsString)));

        // Prioridade de distribuição (menor = mais prioritário; 99 = sem prioridade)
        if NomePlataforma <> '' then
        begin
          if Qry.FieldByName('PrioridadeDistribuicao').IsNull then
            FCachePrioridade.AddOrSetValue(NomePlataforma, 99)
          else
            FCachePrioridade.AddOrSetValue(NomePlataforma,
              Qry.FieldByName('PrioridadeDistribuicao').AsInteger);
        end;

        // Coordenadas físicas ficam indexadas pelo NomeSAP
        if (GrupoFisico <> '') and (not FCacheCoord.ContainsKey(GrupoFisico)) then
        begin
          Coord.Nome := GrupoFisico;
          Coord.Latitude := Qry.FieldByName('Latitude').AsFloat;
          Coord.Longitude := Qry.FieldByName('Longitude').AsFloat;
          FCacheCoord.Add(GrupoFisico, Coord);
        end;

        // Hora operacional continua vinculada ao nome operacional (Plataforma)
        HoraStr := Trim(Qry.FieldByName('HoraSaidaOrigem').AsString);
        if HoraStr <> '' then
          Hora := StrToTimeDef(HoraStr, EncodeTime(6, 0, 0, 0))
        else
          Hora := EncodeTime(6, 0, 0, 0); // default 06:00

        if (NomePlataforma <> '') and (not FCacheHoraSaida.ContainsKey(NomePlataforma)) then
          FCacheHoraSaida.Add(NomePlataforma, Hora)
        else if NomePlataforma <> '' then
          FCacheHoraSaida.AddOrSetValue(NomePlataforma, Hora);

        Qry.Next;
      end;

      Result := (FCacheCoord.Count > 0);
    except
      on E: Exception do
        ShowMessage('Erro ao carregar coordenadas, grupos físicos e horários: ' + E.Message);
    end;
  finally
    Qry.Free;
  end;
end;


constructor TCriacaoRotas.Create(AConnColibri, AConnConsulta: TADOConnection);
begin
  FConnColibri := AConnColibri;
  FConnConsulta := AConnConsulta;
  FCacheCoord := TDictionary<string, TPlataformaCoord>.Create;
  FCacheHoraSaida := TDictionary<string, TTime>.Create;
  FCachePrioridade := TDictionary<string, Integer>.Create;
  FCacheGrupoFisico := TDictionary<string, string>.Create;
  FCacheModal := TDictionary<string, string>.Create;
  FPendentesHub := TList<TExecutantePendente>.Create;
  FHubPlataforma := '';
  SetLength(FCachePares, 0);
end;

destructor TCriacaoRotas.Destroy;
begin
  FCacheCoord.Free;
  FCacheHoraSaida.Free;
  FCachePrioridade.Free;
  FCacheGrupoFisico.Free;
  FCacheModal.Free;
  FPendentesHub.Free;
  inherited;
end;

function TCriacaoRotas.ObterPrioridade(const APlataforma: string): Integer;
var
  NormP: string;
begin
  NormP := NormalizarPlataforma(APlataforma);
  if (FCachePrioridade <> nil) and FCachePrioridade.ContainsKey(NormP) then
    Result := FCachePrioridade[NormP]
  else
    Result := 99; // sem prioridade definida
end;

procedure TCriacaoRotas.CarregarParesObrigatorios;
var
  Qry: TADOQuery;
  Par: TParObrigatorio;
begin
  SetLength(FCachePares, 0);
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT PlataformaA, PlataformaB ' +
      'FROM tblParObrigatorio ' +
      'WHERE Ativo = True';
    try
      Qry.Open;
      while not Qry.Eof do
      begin
        Par.PlataformaA := Trim(Qry.FieldByName('PlataformaA').AsString);
        Par.PlataformaB := Trim(Qry.FieldByName('PlataformaB').AsString);
        if (Par.PlataformaA <> '') and (Par.PlataformaB <> '') then
        begin
          SetLength(FCachePares, Length(FCachePares) + 1);
          FCachePares[High(FCachePares)] := Par;
        end;
        Qry.Next;
      end;
    except
      // Silencioso: tblParObrigatorio pode não existir em DB mais antigo
    end;
  finally
    Qry.Free;
  end;
end;

procedure TCriacaoRotas.CarregarHubPlataforma;
var
  Qry: TADOQuery;
begin
  FHubPlataforma := '';
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT TOP 1 Plataforma ' +
      'FROM tblPlataforma ' +
      'WHERE booleanHubPrincipal = True';
    try
      Qry.Open;
      if not Qry.IsEmpty then
        FHubPlataforma := Trim(Qry.FieldByName('Plataforma').AsString);
    except
      // Silencioso: campo pode não existir em DB mais antigo
    end;
  finally
    Qry.Free;
  end;
end;

// Insere a parada hub na posição da sequência que minimiza distância adicional.
// ASeqArray deve incluir a origem como primeiro elemento.
function TCriacaoRotas.InserirParadaHubNaSequencia(
  const ASeqArray: TArray<string>): TArray<string>;
var
  HubCoord, C1, C2: TPlataformaCoord;
  MelhorCusto, CustoSemHub, CustoComHub: Double;
  MelhorPos, I: Integer;
begin
  // Se já está na sequência, não duplica
  for I := 0 to High(ASeqArray) do
    if SameText(NormalizarPlataforma(ASeqArray[I]),
                NormalizarPlataforma(FHubPlataforma)) then
    begin
      Result := ASeqArray;
      Exit;
    end;

  if Length(ASeqArray) = 0 then
  begin
    Result := ASeqArray;
    Exit;
  end;

  HubCoord := ObterCoordenadasPlataforma(NormalizarPlataforma(FHubPlataforma));

  // Caso: sequência com apenas a origem → hub vai logo após
  if Length(ASeqArray) = 1 then
  begin
    SetLength(Result, 2);
    Result[0] := ASeqArray[0];
    Result[1] := NormalizarPlataforma(FHubPlataforma);
    Exit;
  end;

  // Avaliar inserção entre cada par de paradas consecutivas
  MelhorCusto := MaxDouble;
  MelhorPos := 1; // padrão: após a origem

  for I := 0 to High(ASeqArray) - 1 do
  begin
    C1 := ObterCoordenadasPlataforma(NormalizarPlataforma(ASeqArray[I]));
    C2 := ObterCoordenadasPlataforma(NormalizarPlataforma(ASeqArray[I + 1]));

    CustoSemHub  := CalcularDistancia(C1, C2);
    CustoComHub  := CalcularDistancia(C1, HubCoord) +
                    CalcularDistancia(HubCoord, C2);

    if (CustoComHub - CustoSemHub) < MelhorCusto then
    begin
      MelhorCusto := CustoComHub - CustoSemHub;
      MelhorPos   := I + 1; // inserir antes de ASeqArray[I+1]
    end;
  end;

  // Montar nova sequência com hub inserido em MelhorPos
  SetLength(Result, Length(ASeqArray) + 1);
  for I := 0 to MelhorPos - 1 do
    Result[I] := ASeqArray[I];
  Result[MelhorPos] := NormalizarPlataforma(FHubPlataforma);
  for I := MelhorPos to High(ASeqArray) do
    Result[I + 1] := ASeqArray[I];
end;

// Tenta absorver executantes hub cuja Origem = FHubPlataforma e cujo Destino
// já está entre os destinos do lote ASugestao.
// Insere a parada hub na sequência e adiciona os executantes ao lote.
// Retorna o número de executantes absorvidos.
function TCriacaoRotas.AbsorverExecutantesHub(
  var ASugestao: TSugestaoNovaRota): Integer;
var
  DestinosDoLote: TDictionary<string, Boolean>;
  Absorvidos: TList<TExecutantePendente>;
  I, J: Integer;
  DestNorm: string;
  SeqArray: TArray<string>;
begin
  Result := 0;

  if (FHubPlataforma = '') or (FPendentesHub.Count = 0) then
    Exit;

  // Montar set de destinos já no lote
  DestinosDoLote := TDictionary<string, Boolean>.Create;
  Absorvidos := TList<TExecutantePendente>.Create;
  try
    for I := 0 to High(ASugestao.Destinos) do
      DestinosDoLote.AddOrSetValue(
        NormalizarPlataforma(ASugestao.Destinos[I]), True);

    // Selecionar hub executantes que vão a destinos já na rota
    for I := 0 to FPendentesHub.Count - 1 do
    begin
      DestNorm := NormalizarPlataforma(FPendentesHub[I].Destino);
      if DestinosDoLote.ContainsKey(DestNorm) then
        Absorvidos.Add(FPendentesHub[I]);
    end;

    if Absorvidos.Count = 0 then
      Exit;

    // Inserir parada hub na melhor posição da sequência atual
    SeqArray := SplitSequencia(ASugestao.SequenciaOtimizada);
    SeqArray  := InserirParadaHubNaSequencia(SeqArray);
    ASugestao.SequenciaOtimizada := string.Join(';', SeqArray);

    // Adicionar executantes absorvidos ao lote
    J := Length(ASugestao.Executantes);
    SetLength(ASugestao.Executantes, J + Absorvidos.Count);
    SetLength(ASugestao.ExecutantesIds, J + Absorvidos.Count);
    for I := 0 to Absorvidos.Count - 1 do
    begin
      ASugestao.Executantes[J + I]    := Absorvidos[I];
      ASugestao.ExecutantesIds[J + I] := Absorvidos[I].IdExecutante;
    end;
    ASugestao.QtdExecutantes := Length(ASugestao.Executantes);

    // Remover absorvidos da lista de pendentes hub
    for I := 0 to Absorvidos.Count - 1 do
    begin
      for J := FPendentesHub.Count - 1 downto 0 do
        if FPendentesHub[J].IdExecutante = Absorvidos[I].IdExecutante then
        begin
          FPendentesHub.Delete(J);
          Break;
        end;
    end;

    Result := Absorvidos.Count;
  finally
    DestinosDoLote.Free;
    Absorvidos.Free;
  end;
end;

function TCriacaoRotas.DeveConsiderarDistribuicaoMaritima(
  const AOrigem, ADestino: string): Boolean;
var
  ModalOrigem, ModalDestino: string;
begin
  ModalOrigem := ObterModalPlataforma(AOrigem);
  ModalDestino := ObterModalPlataforma(ADestino);

  // Regra 1: se qualquer lado for A, exclui
  if (ModalOrigem = 'A') or (ModalDestino = 'A') then
    Exit(False);

  // Regra 2: T + T exclui
  if (ModalOrigem = 'T') and (ModalDestino = 'T') then
    Exit(False);

  // Regra 3: só aceita M+M, T+M, M+T
  Result :=
    ((ModalOrigem = 'M') and (ModalDestino = 'M')) or
    ((ModalOrigem = 'T') and (ModalDestino = 'M')) or
    ((ModalOrigem = 'M') and (ModalDestino = 'T'));
end;

// =============================================================================
// MÉTODOS AUXILIARES
// =============================================================================

function TCriacaoRotas.ObterGrupoFisico(const APlataforma: string): string;
var
  Qry: TADOQuery;
begin
  Result := Trim(APlataforma);

  if FCacheGrupoFisico.ContainsKey(Result) then
    Exit(Trim(FCacheGrupoFisico[Result]));

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT NomeSAP ' +
      'FROM tblPlataforma ' +
      'WHERE Plataforma = :Plataforma';
    Qry.Parameters.ParamByName('Plataforma').Value := Trim(APlataforma);
    Qry.Open;

    if (not Qry.IsEmpty) and (Trim(Qry.FieldByName('NomeSAP').AsString) <> '') then
    begin
      Result := Trim(Qry.FieldByName('NomeSAP').AsString);
      FCacheGrupoFisico.AddOrSetValue(Trim(APlataforma), Result);
    end;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.ObterCoordenadasPlataforma(const ANome: string): TPlataformaCoord;
var
  GrupoFisico: string;
  CoordFisica: TPlataformaCoord;
begin
  Result.Nome := Trim(ANome);
  Result.Latitude := 0;
  Result.Longitude := 0;

  GrupoFisico := ObterGrupoFisico(ANome);

  if FCacheCoord.TryGetValue(GrupoFisico, CoordFisica) then
  begin
    // preserva o nome operacional, mas usa a coordenada física do grupo
    Result.Nome := Trim(ANome);
    Result.Latitude := CoordFisica.Latitude;
    Result.Longitude := CoordFisica.Longitude;
  end;
end;

function TCriacaoRotas.ObterDataBaseProg(const ADataProg: string): TDateTime;
var
  FS: TFormatSettings;
begin
  FS := TFormatSettings.Create;
  FS.DateSeparator := '/';
  FS.ShortDateFormat := 'dd/mm/yyyy';

  if not TryStrToDate(ADataProg, Result, FS) then
    Result := Date;

  Result := Trunc(Result);
end;

function TCriacaoRotas.CalcularDistancia(const P1, P2: TPlataformaCoord): Double;
var
  Lat1, Lon1, Lat2, Lon2: Double;
  R, dLat, dLon, a, c: Double;
begin
  // Fórmula de Haversine para calcular distância entre coordenadas
  R := 6371; // Raio da Terra em km
  
  Lat1 := DegToRad(P1.Latitude);
  Lon1 := DegToRad(P1.Longitude);
  Lat2 := DegToRad(P2.Latitude);
  Lon2 := DegToRad(P2.Longitude);
  
  dLat := Lat2 - Lat1;
  dLon := Lon2 - Lon1;
  
  a := Sin(dLat/2) * Sin(dLat/2) +
       Cos(Lat1) * Cos(Lat2) *
       Sin(dLon/2) * Sin(dLon/2);
       
  c := 2 * ArcTan2(Sqrt(a), Sqrt(1-a));
  
  Result := R * c;
end;

function TCriacaoRotas.OtimizarSequencia(const AOrigem: string;
  const ADestinos: TArray<string>): TArray<string>;
var
  NaoVisitadas: TList<TPlataformaCoord>;
  GrupoPrio: TList<TPlataformaCoord>;
  Prioridades: TList<Integer>;
  Atual, Proxima: TPlataformaCoord;
  MenorDistancia, Dist: Double;
  I, IndiceProxima, PrioAtual: Integer;
  ResultList: TList<string>;
  DestinosLimpos: TArray<string>;
begin
  NaoVisitadas := TList<TPlataformaCoord>.Create;
  GrupoPrio    := TList<TPlataformaCoord>.Create;
  Prioridades  := TList<Integer>.Create;
  ResultList   := TList<string>.Create;
  try
    Atual := ObterCoordenadasPlataforma(NormalizarPlataforma(AOrigem));
    ResultList.Add(NormalizarPlataforma(AOrigem));

    DestinosLimpos := RemoverDuplicadosMantendoOrdem(ADestinos, AOrigem);

    // Coletar níveis de prioridade únicos e ordenar (menor = mais prioritário)
    for I := 0 to High(DestinosLimpos) do
    begin
      PrioAtual := ObterPrioridade(DestinosLimpos[I]);
      if not Prioridades.Contains(PrioAtual) then
        Prioridades.Add(PrioAtual);
    end;
    Prioridades.Sort;

    // Nearest-neighbor por grupo de prioridade, em ordem crescente
    for PrioAtual in Prioridades do
    begin
      GrupoPrio.Clear;
      for I := 0 to High(DestinosLimpos) do
        if ObterPrioridade(DestinosLimpos[I]) = PrioAtual then
          GrupoPrio.Add(ObterCoordenadasPlataforma(DestinosLimpos[I]));

      NaoVisitadas.Clear;
      NaoVisitadas.AddRange(GrupoPrio.ToArray);

      while NaoVisitadas.Count > 0 do
      begin
        MenorDistancia := MaxDouble;
        IndiceProxima  := -1;

        for I := 0 to NaoVisitadas.Count - 1 do
        begin
          Dist := CalcularDistancia(Atual, NaoVisitadas[I]);
          if Dist < MenorDistancia then
          begin
            MenorDistancia := Dist;
            IndiceProxima  := I;
          end;
        end;

        if IndiceProxima >= 0 then
        begin
          Proxima := NaoVisitadas[IndiceProxima];
          ResultList.Add(NormalizarPlataforma(Proxima.Nome));
          Atual := Proxima;
          NaoVisitadas.Delete(IndiceProxima);
        end
        else
          Break;
      end;
    end;

    Result := RemoverDuplicadosMantendoOrdem(ResultList.ToArray);
  finally
    NaoVisitadas.Free;
    GrupoPrio.Free;
    Prioridades.Free;
    ResultList.Free;
  end;
end;

// Nova função para obter horário do cache:
function TCriacaoRotas.ObterHoraSaidaOrigem(const AOrigem: string): TTime;
begin
  Result := EncodeTime(6, 0, 0, 0); // Default 06:00

  if FCacheHoraSaida.ContainsKey(AOrigem) then
    Result := FCacheHoraSaida[AOrigem];
end;

function TCriacaoRotas.ObterProximoNomeRota(const AOrigem: string; const ADataProg: string): string;
var
  Qry: TADOQuery;
  QtdRotas: Integer;
begin
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text := 'SELECT COUNT(*) AS Total FROM tblRoteamento ' +
                    'WHERE DataRoteamento = :DataProg AND NomeRota LIKE :NomeLike';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Parameters.ParamByName('NomeLike').Value := '% ' + AOrigem + '%';
    Qry.Open;
    
    QtdRotas := Qry.FieldByName('Total').AsInteger;
    
    // Formato: "1ª PCM-09", "2ª PCM-09", etc.
    Result := IntToStr(QtdRotas + 1) + 'ª ' + AOrigem;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.ObterProximaHoraPartida(const AOrigem: string;
  const ADestinos: TArray<string>; const ADataProg: string): TTime;
var
  Qry: TADOQuery;
  HoraBase: TTime;
  QtdRotas: Integer;
  GrupoHorario: string;
  I: Integer;
begin
  GrupoHorario := GrupoHorarioDoLote(ADestinos);

  // Quando houver destino com turno, fila por NomeSAP|Turno
  if (GrupoHorario <> '') and (GrupoHorario <> '__CONFLITO__') then
  begin
    HoraBase := ObterHoraSaidaOrigem(AOrigem);

    // Usa a hora operacional do primeiro destino com turno
    for I := 0 to High(ADestinos) do
    begin
      if ObterGrupoHorario(ADestinos[I]) <> '' then
      begin
        HoraBase := ObterHoraSaidaOrigem(ADestinos[I]);
        Break;
      end;
    end;

    QtdRotas := ContarRotasMesmoGrupoHorario(ADataProg, GrupoHorario);
    Result := IncMinute(HoraBase, QtdRotas * 15);
    Exit;
  end;

  // Regra antiga por origem
  HoraBase := ObterHoraSaidaOrigem(AOrigem);

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT COUNT(*) AS Total ' +
      'FROM tblRoteamento ' +
      'WHERE DataRoteamento = :DataProg ' +
      '  AND NomeRota LIKE :NomeLike';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Parameters.ParamByName('NomeLike').Value := '% ' + AOrigem + '%';
    Qry.Open;

    QtdRotas := Qry.FieldByName('Total').AsInteger;
    Result := IncMinute(HoraBase, QtdRotas * 15);
  finally
    Qry.Free;
  end;
end;

// =============================================================================
// 1. ANÁLISE E AGRUPAMENTO
// =============================================================================

function TCriacaoRotas.AnalisarExecutantesPendentes(const ADataProg: string;
    const AOrigemPlataformasValidas: string): TArray<TSugestaoNovaRota>;
var
  QryExec: TADOQuery;
  DictOrigens: TObjectDictionary<string, TList<TExecutantePendente>>;
  ListaOrigens: TList<string>;
  OrigemAtual: string;
  Exec: TExecutantePendente;
  I, J: Integer;
  ListaSugestoes: TList<TSugestaoNovaRota>;
  Sugestao: TSugestaoNovaRota;
  DestinosUnicos: TList<string>;
  SeqArray: TArray<string>;
begin
  // ADICIONE ESTA LINHA NO INÍCIO:
  if not CarregarTodasAsCoordenadasDeUmaVez then
  begin
    ShowMessage('Erro: Não foi possível carregar as coordenadas das plataformas.');
    SetLength(Result, 0);
    Exit;
  end;

  DictOrigens := TObjectDictionary<string, TList<TExecutantePendente>>.Create([doOwnsValues]);
  ListaOrigens := TList<string>.Create;
  ListaSugestoes := TList<TSugestaoNovaRota>.Create;
  DestinosUnicos := TList<string>.Create;
  QryExec := TADOQuery.Create(nil);
  try
    // 1. Buscar todos os executantes pendentes
    QryExec.Connection := FConnColibri;
    QryExec.SQL.Text := 
      'SELECT pe.idProgramacaoExecutante, pe.Origem, pd.txtDestino AS Destino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      'AND (pe.InseridoProgramacaoTransporte = False OR pe.InseridoProgramacaoTransporte IS NULL) ' +
      'AND NOT EXISTS (' +
      '  SELECT 1 ' +
      '  FROM tblAux_Rota_Distribuicao ard ' +
      '  WHERE ard.CodigoProgramacaoExecutante = pe.idProgramacaoExecutante' +
      ')';
    QryExec.Parameters.ParamByName('DataProg').Value := ADataProg;
    QryExec.Open;
    
    // 2. Agrupar por Origem
    while not QryExec.Eof do
    begin
      Exec.IdExecutante := QryExec.FieldByName('idProgramacaoExecutante').AsInteger;
      Exec.Origem := Trim(QryExec.FieldByName('Origem').AsString);
      Exec.Destino := Trim(QryExec.FieldByName('Destino').AsString);

      if not DeveConsiderarDistribuicaoMaritima(Exec.Origem, Exec.Destino) then
      begin
        QryExec.Next;
        Continue;
      end;

      if Exec.Origem = Exec.Destino then
      begin
        QryExec.Next;
        Continue;
      end;

      // Executantes do hub são tratados separadamente (absorção em rotas existentes)
      if (FHubPlataforma <> '') and
         SameText(NormalizarPlataforma(Exec.Origem),
                  NormalizarPlataforma(FHubPlataforma)) then
      begin
        FPendentesHub.Add(Exec);
        QryExec.Next;
        Continue;
      end;

      // VALIDAÇÃO: Pula executantes cuja origem não está marcada como Origem válida
      if not ValidarOrigemPlataforma(Exec.Origem, AOrigemPlataformasValidas) then
      begin
        QryExec.Next;
        Continue;
      end;

      if Exec.Origem <> '' then
      begin
        if not DictOrigens.ContainsKey(Exec.Origem) then
        begin
          DictOrigens.Add(Exec.Origem, TList<TExecutantePendente>.Create);
          ListaOrigens.Add(Exec.Origem);
        end;

        DictOrigens[Exec.Origem].Add(Exec);
      end;
      
      QryExec.Next;
    end;
    
    // 3. Criar sugestões por Origem
    for I := 0 to ListaOrigens.Count - 1 do
    begin
      OrigemAtual := ListaOrigens[I];
      DestinosUnicos.Clear;
      
      Sugestao.Origem := OrigemAtual;
      Sugestao.QtdExecutantes := DictOrigens[OrigemAtual].Count;
      SetLength(Sugestao.ExecutantesIds, Sugestao.QtdExecutantes);

      SetLength(Sugestao.Executantes, DictOrigens[OrigemAtual].Count);
      SetLength(Sugestao.ExecutantesIds, DictOrigens[OrigemAtual].Count);

      for J := 0 to DictOrigens[OrigemAtual].Count - 1 do
      begin
        Exec := DictOrigens[OrigemAtual][J];

        Sugestao.Executantes[J] := Exec;
        Sugestao.ExecutantesIds[J] := Exec.IdExecutante;

        if (Exec.Destino <> '') and not DestinosUnicos.Contains(Exec.Destino) then
          DestinosUnicos.Add(Exec.Destino);
      end;

      Sugestao.QtdExecutantes := Length(Sugestao.Executantes);
      Sugestao.Destinos := DestinosUnicos.ToArray;
      
      // 4. Otimizar Sequência (Origem -> Destinos por distância)
      SeqArray := OtimizarSequencia(OrigemAtual, Sugestao.Destinos);
      Sugestao.SequenciaOtimizada := string.Join(';', SeqArray);
      
      // 5. Definir Nome e Horário
      Sugestao.NomeRotaSugerido := ObterProximoNomeRota(OrigemAtual, ADataProg);
      Sugestao.HoraPartidaSugerida :=
        ObterProximaHoraPartida(OrigemAtual, Sugestao.Destinos, ADataProg);

      ListaSugestoes.Add(Sugestao);
    end;
    
    Result := ListaSugestoes.ToArray;
  finally
    QryExec.Free;
    DictOrigens.Free;
    ListaOrigens.Free;
    ListaSugestoes.Free;
    DestinosUnicos.Free;
  end;
end;

// =============================================================================
// 2. CRIAÇÃO DE ROTA
// =============================================================================
function TCriacaoRotas.CriarRota(const ASugestao: TSugestaoNovaRota;
  const ADataProg: string; const ANomeEmbarcacao: string;
  out AMensagemErro: string): Integer;
var
  QryInsert, QryID: TADOQuery;
  Capacidade: Integer;
begin
  Result := 0;
  AMensagemErro := '';

  QryInsert := TADOQuery.Create(nil);
  QryID := TADOQuery.Create(nil);
  try
    Capacidade := ObterCapacidadeEmbarcacao(ANomeEmbarcacao);

    QryInsert.Connection := FConnColibri;
    QryInsert.SQL.Text :=
      'INSERT INTO tblRoteamento (' +
      '  NomeRota, NomeEmbarcacao, DataRoteamento, HoraRoteamento, RotaSequencia, CapacidadePAX' +
      ') VALUES (' +
      '  :NomeRota, :Embarcacao, :DataProg, :Hora, :Sequencia, :Capacidade' +
      ')';

    QryInsert.Parameters.ParamByName('NomeRota').Value := ASugestao.NomeRotaSugerido;
    QryInsert.Parameters.ParamByName('Embarcacao').Value := ANomeEmbarcacao;
    QryInsert.Parameters.ParamByName('DataProg').Value := ADataProg;
    QryInsert.Parameters.ParamByName('Hora').Value := FormatDateTime('hh:nn', ASugestao.HoraPartidaSugerida);
    QryInsert.Parameters.ParamByName('Sequencia').Value :=
      NormalizarSequenciaTexto(ASugestao.SequenciaOtimizada);
    QryInsert.Parameters.ParamByName('Capacidade').Value := Capacidade;
    try
      QryInsert.ExecSQL;

      QryID.Connection := FConnColibri;
      QryID.SQL.Text := 'SELECT @@IDENTITY AS NovoID';
      QryID.Open;

      if not QryID.IsEmpty then
        Result := QryID.FieldByName('NovoID').AsInteger;
    except
      on E: Exception do
        AMensagemErro := E.Message;
    end;
  finally
    QryInsert.Free;
    QryID.Free;
  end;
end;

// =============================================================================
// 3. PROCESSO COMPLETO (BATCH)
// =============================================================================
function TCriacaoRotas.CriarEAlocarLote(const ADataProg: string;
  const AEmbarcacoesDisponiveis: TArray<string>;
  const AOrigemPlataformasValidas: string; out AMensagem: string): Boolean;
var
  Sugestoes: TArray<TSugestaoNovaRota>;
  Estados: TArray<TEstadoEmbarcacao>;
  DistLogistica: TDistribuicaoLogistica;
  RotasExclusivas: TDictionary<Integer, Byte>;
  PendentesOrigem: TList<TExecutantePendente>;
  ExecutantesOrdenados: TArray<TExecutantePendente>;
  SugestaoLote: TSugestaoNovaRota;
  MsgErro, MsgAlocacaoExistentes, MsgBridge: string;
  I, J, K, IdxEmb, IdNovaRota, QtdLote: Integer;
  Sucessos, Falhas, RotasCriadas: Integer;
  TotalPendentes, TotalEtapas, EtapaAtual: Integer;
  LoteBridge: TArray<TExecutantePendente>;
  NomeEmbarcacaoRota: string;
  CapBridge: Integer;

  procedure AvancarProgresso(const AMsg: string);
  begin
    if EtapaAtual < TotalEtapas then
      Inc(EtapaAtual);
    ReportarProgresso(EtapaAtual, TotalEtapas, AMsg);
  end;

begin
  Result := False;
  TotalPendentes := 0;
  TotalEtapas := 1;
  EtapaAtual := 0;
  AMensagem := '';
  MsgAlocacaoExistentes := '';
  MsgBridge := '';
  Sucessos := 0;
  Falhas := 0;
  RotasCriadas := 0;
  DistLogistica := nil;
  RotasExclusivas := nil;

  ReportarProgresso(0, 1, 'Iniciando distribuição automática...');

  if Length(AEmbarcacoesDisponiveis) = 0 then
  begin
    AMensagem := 'Nenhuma embarcação disponível para distribuição.';
    ReportarProgresso(1, 1, AMensagem);
    Exit;
  end;

  DistLogistica := TDistribuicaoLogistica.Create(FConnColibri, FConnConsulta);
  RotasExclusivas := CarregarRotasExclusivasDoDia(ADataProg);

  if FCacheGrupoFisico.Count = 0 then
    CarregarTodasAsCoordenadasDeUmaVez;

  CarregarParesObrigatorios;

  // Hub: identifica plataforma hub e zera lista de pendentes hub
  CarregarHubPlataforma;
  FPendentesHub.Clear;

  try
    ProcessarRotasBridgeDoDia(
      ADataProg,
      DistLogistica,
      RotasExclusivas,
      MsgBridge,
      Sucessos,
      Falhas,
      RotasCriadas
    );

    if RotasExclusivas.Count > 0 then
      AMensagem := Format(
        '%d rota(s) fixa(s) já vinculada(s) foram preservadas e não entraram na distribuição automática.',
        [RotasExclusivas.Count]
      ) + sLineBreak + AMensagem;

    if Trim(MsgBridge) <> '' then
      AMensagem := AMensagem + MsgBridge + sLineBreak;

    // NÃO chamar aqui se você já limpou tudo antes
    // DistLogistica.AlocarLoteAutomaticamente(ADataProg, MsgAlocacaoExistentes);

    Sugestoes := AnalisarExecutantesPendentes(ADataProg, AOrigemPlataformasValidas);

    if Length(Sugestoes) = 0 then
    begin
      AMensagem :=
        Format('Processo concluído: %d rotas criadas. %d executantes alocados, %d falhas.',
          [RotasCriadas, Sucessos, Falhas]);

      ReportarProgresso(1, 1, 'Processo concluído.');
      Result := True;
      Exit;
    end;

    // Conta pendentes para dimensionar a barra
    for I := 0 to High(Sugestoes) do
      Inc(TotalPendentes, Length(Sugestoes[I].Executantes));

    // Folga suficiente para:
    // - preparação
    // - carregamento de estado
    // - criação de rotas
    // - vinculação dos executantes
    TotalEtapas := 5 + (TotalPendentes * 2);
    EtapaAtual := 0;

    ReportarProgresso(EtapaAtual, TotalEtapas, 'Carregando estado da frota...');

    // Etapa 3: carregar estado atual da frota no dia
    Estados := CarregarEstadoEmbarcacoes(ADataProg, AEmbarcacoesDisponiveis);
    AvancarProgresso('Estado da frota carregado.');

    // Etapa 4: criar quantas rotas forem necessárias por origem
    for I := 0 to High(Sugestoes) do
    begin
      PendentesOrigem := TList<TExecutantePendente>.Create;
      try
        ExecutantesOrdenados := Copy(Sugestoes[I].Executantes);
        OrdenarExecutantesPorSequencia(ExecutantesOrdenados, Sugestoes[I].Origem);

        for J := 0 to High(ExecutantesOrdenados) do
          PendentesOrigem.Add(ExecutantesOrdenados[J]);

        while PendentesOrigem.Count > 0 do
        begin
          NomeEmbarcacaoRota := '';
          IdxEmb := -1;
          SetLength(LoteBridge, 0);

          // 1) Tenta primeiro usar a própria SOV/bridge,
          // sem concorrer com embarcação externa
          if OrigemUsaBridge(Sugestoes[I].Origem) then
          begin
            NomeEmbarcacaoRota := ObterNomeEmbarcacaoBridge(Sugestoes[I].Origem);
            CapBridge := ObterCapacidadeEmbarcacao(NomeEmbarcacaoRota);

            if NomeEmbarcacaoRota = '' then
              CapBridge := 0;

            LoteBridge := MontarLoteBridgeMesmoGrupo(
              Sugestoes[I].Origem,
              PendentesOrigem,
              CapBridge
            );

            if Length(LoteBridge) > 0 then
              SugestaoLote := MontarSugestaoParaLote(
                Sugestoes[I].Origem,
                LoteBridge,
                ADataProg
              );
          end;

          // 2) Se não montou lote bridge, usa fluxo normal de embarcação externa
          if NomeEmbarcacaoRota = '' then
          begin
            IdxEmb := EscolherMelhorEmbarcacaoParaPendentes(
              Estados,
              PendentesOrigem,
              Sugestoes[I].Origem,
              ADataProg,
              SugestaoLote
            );

            if IdxEmb < 0 then
            begin
              Inc(Falhas, PendentesOrigem.Count);
              AMensagem := AMensagem +
                Format(
                  'Nenhuma embarcação elegível para continuar a origem %s. Restaram %d executantes sem alocação.',
                  [Sugestoes[I].Origem, PendentesOrigem.Count]
                ) + sLineBreak;
              Break;
            end;

            NomeEmbarcacaoRota := Estados[IdxEmb].Nome;
          end;

          // Tenta absorver executantes hub na rota antes de criá-la
          AbsorverExecutantesHub(SugestaoLote);

          QtdLote := Length(SugestaoLote.Executantes);

          if QtdLote <= 0 then
          begin
            Inc(Falhas, PendentesOrigem.Count);
            AMensagem := AMensagem +
              Format(
                'Não foi possível montar lote para a origem %s.',
                [Sugestoes[I].Origem]
              ) + sLineBreak;
            Break;
          end;

          AvancarProgresso(
            Format(
              'Criando rota para origem %s com embarcação %s...',
              [SugestaoLote.Origem, NomeEmbarcacaoRota]
            )
          );

          IdNovaRota := CriarRota(
            SugestaoLote,
            ADataProg,
            NomeEmbarcacaoRota,
            MsgErro
          );

          if IdNovaRota > 0 then
          begin
            Inc(RotasCriadas);

            for K := 0 to High(SugestaoLote.Executantes) do
            begin
              if DistLogistica.VincularExecutanteARota(
                   SugestaoLote.Executantes[K].IdExecutante,
                   IdNovaRota,
                   MsgErro
                 ) then
              begin
                FConnColibri.Execute(
                  Format(
                    'UPDATE tblProgramacaoExecutante ' +
                    'SET InseridoProgramacaoTransporte = True ' +
                    'WHERE idProgramacaoExecutante = %d',
                    [SugestaoLote.Executantes[K].IdExecutante]
                  )
                );
                Inc(Sucessos);
              end
              else
              begin
                Inc(Falhas);
                AMensagem := AMensagem +
                  Format(
                    'Falha ao vincular executante %d à rota %d: %s',
                    [SugestaoLote.Executantes[K].IdExecutante, IdNovaRota, MsgErro]
                  ) + sLineBreak;
              end;

              AvancarProgresso(
                Format(
                  'Alocando executantes - rota %d (%s)',
                  [IdNovaRota, SugestaoLote.Origem]
                )
              );
            end;

            // Remove exatamente os executantes do lote usado
            RemoverLoteDaLista(PendentesOrigem, SugestaoLote.Executantes);

            // Só atualiza estado quando usou embarcação externa
            if IdxEmb >= 0 then
              AtualizarEstadoEmbarcacaoAposRota(Estados[IdxEmb], SugestaoLote, ADataProg);
          end
          else
          begin
            Inc(Falhas, QtdLote);
            AMensagem := AMensagem +
              Format(
                'Erro ao criar rota para origem %s com embarcação %s: %s',
                [Sugestoes[I].Origem, NomeEmbarcacaoRota, MsgErro]
              ) + sLineBreak;

            // Remove o lote para não entrar em loop infinito
            RemoverLoteDaLista(PendentesOrigem, SugestaoLote.Executantes);

            AvancarProgresso(
              Format(
                'Falha ao criar rota para origem %s.',
                [Sugestoes[I].Origem]
              )
            );
          end;
        end;
      finally
        PendentesOrigem.Free;
      end;
    end;

    // Fallback hub: executantes não absorvidos recebem rota própria
    if (FHubPlataforma <> '') and (FPendentesHub.Count > 0) then
    begin
      AvancarProgresso(
        Format('Criando rota de fallback para %d executante(s) do hub %s...',
          [FPendentesHub.Count, FHubPlataforma])
      );
      SugestaoLote := MontarSugestaoParaLote(
        FHubPlataforma,
        FPendentesHub.ToArray,
        ADataProg
      );
      IdxEmb := EscolherMelhorEmbarcacaoParaPendentes(
        Estados,
        FPendentesHub,
        FHubPlataforma,
        ADataProg,
        SugestaoLote
      );
      if IdxEmb >= 0 then
      begin
        NomeEmbarcacaoRota := Estados[IdxEmb].Nome;
        IdNovaRota := CriarRota(SugestaoLote, ADataProg, NomeEmbarcacaoRota, MsgErro);
        if IdNovaRota > 0 then
        begin
          Inc(RotasCriadas);
          for K := 0 to High(SugestaoLote.Executantes) do
          begin
            if DistLogistica.VincularExecutanteARota(
                 SugestaoLote.Executantes[K].IdExecutante,
                 IdNovaRota, MsgErro) then
            begin
              FConnColibri.Execute(
                Format('UPDATE tblProgramacaoExecutante ' +
                       'SET InseridoProgramacaoTransporte = True ' +
                       'WHERE idProgramacaoExecutante = %d',
                  [SugestaoLote.Executantes[K].IdExecutante]));
              Inc(Sucessos);
            end
            else
              Inc(Falhas);
          end;
          AtualizarEstadoEmbarcacaoAposRota(Estados[IdxEmb], SugestaoLote, ADataProg);
          FPendentesHub.Clear;
        end
        else
        begin
          Inc(Falhas, FPendentesHub.Count);
          AMensagem := AMensagem +
            Format('Erro ao criar rota de fallback para hub %s: %s',
              [FHubPlataforma, MsgErro]) + sLineBreak;
        end;
      end
      else
      begin
        Inc(Falhas, FPendentesHub.Count);
        AMensagem := AMensagem +
          Format('%d executante(s) do hub %s não foram absorvidos e não há embarcação disponível.',
            [FPendentesHub.Count, FHubPlataforma]) + sLineBreak;
      end;
    end;

    if Trim(MsgAlocacaoExistentes) <> '' then
      AMensagem := MsgAlocacaoExistentes + sLineBreak + sLineBreak + AMensagem;

    if Trim(AMensagem) <> '' then
      AMensagem := AMensagem + sLineBreak + sLineBreak;

    AMensagem := AMensagem +
      Format('Processo concluído: %d rotas criadas. %d executantes alocados, %d falhas.',
        [RotasCriadas, Sucessos, Falhas]);

    ReportarProgresso(TotalEtapas, TotalEtapas, 'Processo concluído.');
    Result := True;
  finally
    RotasExclusivas.Free;
    DistLogistica.Free;
  end;
end;

procedure TCriacaoRotas.RemoverLoteDaLista(
  APendentes: TList<TExecutantePendente>;
  const ALote: TArray<TExecutantePendente>);
var
  Ids: TDictionary<Integer, Byte>;
  I: Integer;
begin
  Ids := TDictionary<Integer, Byte>.Create;
  try
    for I := 0 to High(ALote) do
      Ids.AddOrSetValue(ALote[I].IdExecutante, 0);

    for I := APendentes.Count - 1 downto 0 do
      if Ids.ContainsKey(APendentes[I].IdExecutante) then
        APendentes.Delete(I);
  finally
    Ids.Free;
  end;
end;

function TCriacaoRotas.ObterIndiceDestinoNaSequencia(const ADestino: string;
  const ASequencia: TArray<string>): Integer;
var
  I: Integer;
  DestinoNorm: string;
begin
  Result := MaxInt;
  DestinoNorm := AnsiUpperCase(NormalizarPlataforma(ADestino));

  for I := 0 to High(ASequencia) do
  begin
    if AnsiUpperCase(NormalizarPlataforma(ASequencia[I])) = DestinoNorm then
      Exit(I);
  end;
end;

function TCriacaoRotas.ObterModalPlataforma(const APlataforma: string): string;
var
  Qry: TADOQuery;
begin
  Result := '';

  if FCacheModal.ContainsKey(Trim(APlataforma)) then
    Exit(UpperCase(Trim(FCacheModal[Trim(APlataforma)])));

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT RT_Modal ' +
      'FROM tblPlataforma ' +
      'WHERE Plataforma = :Plataforma';
    Qry.Parameters.ParamByName('Plataforma').Value := Trim(APlataforma);
    Qry.Open;

    if not Qry.IsEmpty then
    begin
      Result := UpperCase(Trim(Qry.FieldByName('RT_Modal').AsString));
      FCacheModal.AddOrSetValue(Trim(APlataforma), Result);
    end;
  finally
    Qry.Free;
  end;
end;

procedure TCriacaoRotas.OrdenarExecutantesPorSequencia(
  var AExecutantes: TArray<TExecutantePendente>; const AOrigem: string);
var
  DestinosUnicos: TStringList;
  SequenciaBase: TArray<string>;
  I, J: Integer;
  Tmp: TExecutantePendente;
  IdxI, IdxJ: Integer;
begin
  if Length(AExecutantes) <= 1 then
    Exit;

  DestinosUnicos := TStringList.Create;
  try
    DestinosUnicos.Sorted := False;
    DestinosUnicos.Duplicates := dupIgnore;

    for I := 0 to High(AExecutantes) do
    begin
      if Trim(AExecutantes[I].Destino) <> '' then
        DestinosUnicos.Add(NormalizarPlataforma(AExecutantes[I].Destino));
    end;

    SequenciaBase := OtimizarSequencia(
      AOrigem,
      RemoverDuplicadosMantendoOrdem(DestinosUnicos.ToStringArray, AOrigem)
    );

    // bubble sort simples e suficiente para esse cenário
    for I := Low(AExecutantes) to High(AExecutantes) - 1 do
    begin
      for J := I + 1 to High(AExecutantes) do
      begin
        IdxI := ObterIndiceDestinoNaSequencia(AExecutantes[I].Destino, SequenciaBase);
        IdxJ := ObterIndiceDestinoNaSequencia(AExecutantes[J].Destino, SequenciaBase);

        if IdxJ < IdxI then
        begin
          Tmp := AExecutantes[I];
          AExecutantes[I] := AExecutantes[J];
          AExecutantes[J] := Tmp;
        end
        else if (IdxJ = IdxI) and
                (AnsiUpperCase(NormalizarPlataforma(AExecutantes[J].Destino)) <
                 AnsiUpperCase(NormalizarPlataforma(AExecutantes[I].Destino))) then
        begin
          Tmp := AExecutantes[I];
          AExecutantes[I] := AExecutantes[J];
          AExecutantes[J] := Tmp;
        end;
      end;
    end;
  finally
    DestinosUnicos.Free;
  end;
end;


function TCriacaoRotas.ExtrairTurnoPlataforma(const APlataforma: string): string;
var
  S: string;
begin
  S := UpperCase(Trim(APlataforma));

  if Pos('(D)', S) > 0 then
    Exit('D');

  if Pos('(N)', S) > 0 then
    Exit('N');

  Result := '';
end;

function TCriacaoRotas.ObterGrupoHorario(const APlataforma: string): string;
var
  GrupoFisico, Turno: string;
begin
  GrupoFisico := ObterGrupoFisico(APlataforma);
  Turno := ExtrairTurnoPlataforma(APlataforma);

  if (GrupoFisico <> '') and (Turno <> '') then
    Result := GrupoFisico + '|' + Turno
  else
    Result := '';
end;

function TCriacaoRotas.GrupoHorarioDoLote(const ADestinos: TArray<string>): string;
var
  I: Integer;
  GrupoAtual, GrupoTrava: string;
begin
  GrupoTrava := '';

  for I := 0 to High(ADestinos) do
  begin
    GrupoAtual := ObterGrupoHorario(ADestinos[I]);

    if GrupoAtual = '' then
      Continue;

    if GrupoTrava = '' then
      GrupoTrava := GrupoAtual
    else if not SameText(GrupoTrava, GrupoAtual) then
      Exit('__CONFLITO__');
  end;

  Result := GrupoTrava;
end;

function TCriacaoRotas.ExtrairDestinosDaSequencia(const ASequencia: string): TArray<string>;
var
  Arr: TArray<string>;
  I, J: Integer;
begin
  Arr := SplitSequencia(ASequencia);

  if Length(Arr) <= 1 then
    Exit(nil);

  SetLength(Result, Length(Arr) - 1);
  J := 0;

  for I := 1 to High(Arr) do
  begin
    Result[J] := Trim(Arr[I]);
    Inc(J);
  end;

  Result := RemoverDuplicadosMantendoOrdem(Result);
end;

function TCriacaoRotas.ContarRotasMesmoGrupoHorario(
  const ADataProg, AGrupoHorario: string): Integer;
var
  Qry: TADOQuery;
  DestinosRota: TArray<string>;
  GrupoRota: string;
begin
  Result := 0;

  if Trim(AGrupoHorario) = '' then
    Exit;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT RotaSequencia ' +
      'FROM tblRoteamento ' +
      'WHERE DataRoteamento = :DataProg';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Open;

    while not Qry.Eof do
    begin
      DestinosRota := ExtrairDestinosDaSequencia(
        Trim(Qry.FieldByName('RotaSequencia').AsString)
      );

      GrupoRota := GrupoHorarioDoLote(DestinosRota);

      if SameText(GrupoRota, AGrupoHorario) then
        Inc(Result);

      Qry.Next;
    end;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.MontarLoteRespeitandoGrupoHorario(
  const APendentes: TList<TExecutantePendente>;
  ACapacidade: Integer
): TArray<TExecutantePendente>;
var
  I, J, K, Count: Integer;
  GrupoTrava, GrupoAtual: string;
  DestinosNoLote: TDictionary<string, Boolean>;
  PareiroDestino: string;
  JaNoLote: Boolean;
begin
  SetLength(Result, 0);

  if APendentes.Count = 0 then
    Exit;

  if ACapacidade <= 0 then
    ACapacidade := APendentes.Count;

  GrupoTrava := '';
  Count := 0;

  for I := 0 to APendentes.Count - 1 do
  begin
    GrupoAtual := ObterGrupoHorario(APendentes[I].Destino);

    // trava no primeiro destino com turno encontrado
    if (GrupoAtual <> '') and (GrupoTrava = '') then
      GrupoTrava := GrupoAtual;

    // se já travou um grupo horário, não deixa misturar outro
    if (GrupoTrava <> '') and (GrupoAtual <> '') and
       (not SameText(GrupoTrava, GrupoAtual)) then
      Continue;

    SetLength(Result, Length(Result) + 1);
    Result[High(Result)] := APendentes[I];
    Inc(Count);

    if Count >= ACapacidade then
      Break;
  end;

  // Expansão por pares obrigatórios:
  // se destino A está no lote, incluir também executantes com destino B do mesmo par
  if (Length(FCachePares) = 0) or (Length(Result) = 0) then
    Exit;

  DestinosNoLote := TDictionary<string, Boolean>.Create;
  try
    for I := 0 to High(Result) do
      DestinosNoLote.AddOrSetValue(NormalizarPlataforma(Result[I].Destino), True);

    for K := 0 to High(FCachePares) do
    begin
      if DestinosNoLote.ContainsKey(NormalizarPlataforma(FCachePares[K].PlataformaA)) then
        PareiroDestino := NormalizarPlataforma(FCachePares[K].PlataformaB)
      else if DestinosNoLote.ContainsKey(NormalizarPlataforma(FCachePares[K].PlataformaB)) then
        PareiroDestino := NormalizarPlataforma(FCachePares[K].PlataformaA)
      else
        Continue;

      // Já está incluído — nada a fazer
      if DestinosNoLote.ContainsKey(PareiroDestino) then
        Continue;

      // Adicionar pendentes que vão ao destino pareiro (se ainda cabem)
      for J := 0 to APendentes.Count - 1 do
      begin
        if Count >= ACapacidade then Break;
        if NormalizarPlataforma(APendentes[J].Destino) <> PareiroDestino then Continue;

        JaNoLote := False;
        for I := 0 to High(Result) do
          if Result[I].IdExecutante = APendentes[J].IdExecutante then
          begin
            JaNoLote := True;
            Break;
          end;

        if not JaNoLote then
        begin
          SetLength(Result, Length(Result) + 1);
          Result[High(Result)] := APendentes[J];
          Inc(Count);
          DestinosNoLote.AddOrSetValue(PareiroDestino, True);
        end;
      end;
    end;
  finally
    DestinosNoLote.Free;
  end;
end;

function TCriacaoRotas.OrigemUsaBridge(const AOrigem: string): Boolean;
var
  Qry: TADOQuery;
begin
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT TOP 1 NomeEmbarcacao ' +
      'FROM tblEmbarcacao ' +
      'WHERE UsaBridgeMesmoGrupo = True ' +
      '  AND (NomeEmbarcacao = :Origem OR NomeEmbarcacao = :OrigemBridge)';
    Qry.Parameters.ParamByName('Origem').Value := Trim(AOrigem);
    Qry.Parameters.ParamByName('OrigemBridge').Value := Trim(AOrigem) + ' BRIDGE';
    Qry.Open;

    Result := not Qry.IsEmpty;
  finally
    Qry.Free;
  end;
end;

function TCriacaoRotas.MontarLoteBridgeMesmoGrupo(
  const AOrigem: string;
  const APendentes: TList<TExecutantePendente>;
  ACapacidade: Integer): TArray<TExecutantePendente>;
var
  I, Count: Integer;
  GrupoOrigem, GrupoAtual: string;
  GrupoHorarioTrava, GrupoHorarioAtual: string;
begin
  SetLength(Result, 0);

  if APendentes.Count = 0 then
    Exit;

  if ACapacidade <= 0 then
    ACapacidade := APendentes.Count;

  GrupoOrigem := ObterGrupoFisico(AOrigem);
  GrupoHorarioTrava := '';
  Count := 0;

  for I := 0 to APendentes.Count - 1 do
  begin
    GrupoAtual := ObterGrupoFisico(APendentes[I].Destino);
    if not SameText(GrupoOrigem, GrupoAtual) then
      Continue;

    GrupoHorarioAtual := ObterGrupoHorario(APendentes[I].Destino);

    if (GrupoHorarioAtual <> '') and (GrupoHorarioTrava = '') then
      GrupoHorarioTrava := GrupoHorarioAtual;

    if (GrupoHorarioTrava <> '') and (GrupoHorarioAtual <> '') and
       (not SameText(GrupoHorarioTrava, GrupoHorarioAtual)) then
      Continue;

    SetLength(Result, Length(Result) + 1);
    Result[High(Result)] := APendentes[I];
    Inc(Count);

    if Count >= ACapacidade then
      Break;
  end;
end;

end.
