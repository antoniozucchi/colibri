unit uPreparacaoSolver;

{
  Responsabilidade: preparar os dados de entrada para o solver.py.

  Fluxo:
    1. FiltrarExecutantes  — lê a programação do dia, aplica regras de exclusão
                             e retorna elegíveis + excluídos com motivo.
    2. AgregarDemanda      — agrupa os elegíveis por plataforma de destino,
                             produzindo o vetor (plataforma, PAX_TMIB, PAX_M9,
                             prioridade) que o solver consome.
    3. CarregarListaEmbarcacoes — lista embarcações marcadas para distribuição
                             automática (campo Distribuicao = True).
    4. PrepararInput       — une os passos 1-3 em uma estrutura TDadosSolverInput.
    5. GerarArquivoXlsx    — grava o solver_input.xlsx no formato exigido pelo
                             solver.py (via Excel OLE Automation).

  Nota: HoraPartida e RotaFixa dos barcos ficam em branco na preparação e devem
  ser preenchidos pela interface operacional antes de chamar GerarArquivoXlsx.
}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.RegularExpressions, System.Math,
  Data.DB, Data.Win.ADODB,
  ComObj, System.Variants, System.StrUtils;

const
  // Motivos de exclusão da distribuição automática
  MOTIVO_MODAL_AEREO          = 'MODAL_AEREO';
  MOTIVO_BRIDGE_SOV           = 'BRIDGE_SOV';
  MOTIVO_SEM_CODIGO_SOLVER    = 'SEM_CODIGO_SOLVER';
  MOTIVO_EXCLUIDO_MANUAL      = 'EXCLUIDO_MANUAL';
  MOTIVO_DESTINO_VAZIO        = 'DESTINO_VAZIO';
  MOTIVO_ORIGEM_NAO_SUPORTADA = 'ORIGEM_NAO_SUPORTADA';

  // Valores do campo tblPlataforma.TipoNoLogistico
  // Classificam o papel estrutural do nó na rede de distribuição offshore.
  TIPO_NO_TERMINAL  = 'TERMINAL';  // terminal marítimo em terra (ex: TMIB) — RT_Modal = T
  TIPO_NO_HUB       = 'HUB';       // plataforma de concentração/distribuição de PAX (ex: PCM-09)

type
  // ---------------------------------------------------------------------------
  // Estruturas de dados
  // ---------------------------------------------------------------------------

  TExecutanteExcluido = record
    IdProgramacaoExecutante: Integer;
    Origem: string;
    Destino: string;
    MotivoExclusao: string;   // constante MOTIVO_*
    DescricaoMotivo: string;  // texto legível para exibição e gravação no banco
  end;

  TExecutanteElegivel = record
    IdProgramacaoExecutante: Integer;
    Origem: string;
    OrigemSolver: string;          // 'TMIB' ou 'M9'
    Destino: string;               // nome no formato Colibri: 'PCM-01 (D)'
    CodigoNormSolver: string;      // código que o solver entende: 'PCM-01'
    GrupoHorario: string;          // 'D', 'N' ou '' (sem restrição de turno)
    PrioridadeDistribuicao: Integer;
  end;

  TDemandaSolver = record
    CodigoPlataforma: string;  // CodigoNormSolver do destino
    PaxTMIB: Integer;
    PaxM9: Integer;
    Prioridade: Integer;       // menor valor entre os executantes do grupo (1 > 2 > 3 > 99)
  end;

  TBarcoSolver = record
    NomeEmbarcacao: string;
    Disponivel: Boolean;
    HoraPartida: string;   // 'HH:MM' — preenchido pela interface operacional
    RotaFixa: string;      // sequência de rota fixa — preenchido pela interface
  end;

  TDadosSolverInput = record
    DataOperacao: TDateTime;
    TrocaTurma: Boolean;
    RendidosM9: Integer;
    GrupoHorarioFiltro: string;  // '' = todos, 'D' ou 'N'
    Barcos: TArray<TBarcoSolver>;
    Demandas: TArray<TDemandaSolver>;
    Elegiveis: TArray<TExecutanteElegivel>;
    Excluidos: TArray<TExecutanteExcluido>;
  end;

  // ---------------------------------------------------------------------------
  // Classe principal
  // ---------------------------------------------------------------------------

  TPreparacaoSolver = class
  private
    FConnColibri: TADOConnection;
    FConnConsulta: TADOConnection;

    // Caches por nome de plataforma (chave em maiúsculas)
    FCacheModal: TDictionary<string, string>;
    FCacheGrupoFisico: TDictionary<string, string>;
    FCacheCodigoNorm: TDictionary<string, string>;
    FCachePrioridade: TDictionary<string, Integer>;
    FCacheExcluirManual: TDictionary<string, Boolean>;
    FCacheOrigemBridge: TDictionary<string, Boolean>;
    FCacheTipoNoLogistico: TDictionary<string, string>;

    procedure LimparCaches;

    // Consultas ao banco (com cache)
    function ObterModalPlataforma(const APlataforma: string): string;
    function ObterGrupoFisicoPlataforma(const APlataforma: string): string;
    function ObterExcluirManualPlataforma(const APlataforma: string): Boolean;
    function ObterPrioridadePlataforma(const APlataforma: string): Integer;
    function EhOrigemBridge(const AOrigem: string): Boolean;
    // Retorna TIPO_NO_TERMINAL, TIPO_NO_HUB ou '' (plataforma offshore regular).
    // Fonte primária: tblPlataforma.TipoNoLogistico.
    // Fallback estrutural quando o campo estiver vazio ou ausente no banco.
    function ObterTipoNoLogistico(const APlataforma: string): string;

    // Helpers sem banco
    function ExtrairTurnoPlataforma(const APlataforma: string): string;
    function DerivarCodigoNorm(const ANomePlataforma: string): string;
    function TryObterCodigoNormSolver(const APlataforma: string;
      out ACodigo: string): Boolean;
    function EhExecutanteBridge(const AOrigem, ADestino: string): Boolean;
    function DeveConsiderarDistribuicaoMaritima(
      const AOrigem, ADestino: string): Boolean;
    // Resolve o papel logístico da origem via TipoNoLogistico do cadastro.
    // Retorna: 'TMIB' (terminal terrestre), 'M9' (hub de distribuição) ou 'OUTRO'.
    function ResolverOrigemSolver(const AOrigem: string): string;

  public
    constructor Create(AConnColibri, AConnConsulta: TADOConnection);
    destructor Destroy; override;

    // Filtra executantes do dia aplicando todas as regras de exclusão.
    // AGrupoHorarioFiltro: '' = sem filtro | 'D' = somente diurno | 'N' = somente noturno
    procedure FiltrarExecutantes(
      const ADataProg: string;
      const AGrupoHorarioFiltro: string;
      out AElegiveis: TArray<TExecutanteElegivel>;
      out AExcluidos: TArray<TExecutanteExcluido>);

    // Agrega demanda dos elegíveis por plataforma de destino.
    function AgregarDemanda(
      const AElegiveis: TArray<TExecutanteElegivel>): TArray<TDemandaSolver>;

    // Detecta quais grupos de turno existem na programação do dia.
    // Retorna subconjunto de {'', 'D', 'N'} de acordo com o que encontrar.
    function DetectarGruposHorario(const ADataProg: string): TArray<string>;

    // Carrega embarcações com Distribuicao = True e Ativo = True.
    // HoraPartida e RotaFixa ficam em branco para preenchimento pela interface.
    function CarregarListaEmbarcacoes: TArray<TBarcoSolver>;

    // Prepara a estrutura completa de entrada para o solver.
    function PrepararInput(
      const ADataProg: string;
      ATrocaTurma: Boolean;
      ARendidosM9: Integer;
      const AGrupoHorarioFiltro: string): TDadosSolverInput;

    // Gera o arquivo solver_input.xlsx no formato exigido pelo solver.py.
    // Requer Microsoft Excel instalado (OLE Automation).
    procedure GerarArquivoXlsx(
      const ADados: TDadosSolverInput;
      const ACaminhoArquivo: string);
  end;

implementation

// Posições fixas do layout lido por read_solver_input() em solver.py
const
  SOLVER_ROW_TROCA_TURMA  = 4;
  SOLVER_ROW_RENDIDOS_M9  = 5;
  SOLVER_FIRST_BOAT_ROW   = 9;
  SOLVER_COL_BOAT_NOME    = 2;
  SOLVER_COL_BOAT_DISP    = 3;
  SOLVER_COL_BOAT_HORA    = 4;
  SOLVER_COL_BOAT_ROTA    = 5;
  SOLVER_COL_DEM_PLAT     = 2;
  SOLVER_COL_DEM_M9       = 3;
  SOLVER_COL_DEM_TMIB     = 4;
  SOLVER_COL_DEM_PRIO     = 5;

{ TPreparacaoSolver }

constructor TPreparacaoSolver.Create(AConnColibri,
  AConnConsulta: TADOConnection);
begin
  FConnColibri  := AConnColibri;
  FConnConsulta := AConnConsulta;

  FCacheModal            := TDictionary<string, string>.Create;
  FCacheGrupoFisico      := TDictionary<string, string>.Create;
  FCacheCodigoNorm       := TDictionary<string, string>.Create;
  FCachePrioridade       := TDictionary<string, Integer>.Create;
  FCacheExcluirManual    := TDictionary<string, Boolean>.Create;
  FCacheOrigemBridge     := TDictionary<string, Boolean>.Create;
  FCacheTipoNoLogistico  := TDictionary<string, string>.Create;
end;

destructor TPreparacaoSolver.Destroy;
begin
  FCacheModal.Free;
  FCacheGrupoFisico.Free;
  FCacheCodigoNorm.Free;
  FCachePrioridade.Free;
  FCacheExcluirManual.Free;
  FCacheOrigemBridge.Free;
  FCacheTipoNoLogistico.Free;
  inherited;
end;

procedure TPreparacaoSolver.LimparCaches;
begin
  FCacheModal.Clear;
  FCacheGrupoFisico.Clear;
  FCacheCodigoNorm.Clear;
  FCachePrioridade.Clear;
  FCacheExcluirManual.Clear;
  FCacheOrigemBridge.Clear;
  FCacheTipoNoLogistico.Clear;
end;

// =============================================================================
// CONSULTAS AO BANCO (COM CACHE)
// =============================================================================

function TPreparacaoSolver.ObterModalPlataforma(
  const APlataforma: string): string;
var
  Chave: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCacheModal.TryGetValue(Chave, Result) then
    Exit;

  Result := '';
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT RT_Modal FROM tblPlataforma WHERE Plataforma = :Plat';
    Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
    Qry.Open;
    if not Qry.IsEmpty then
      Result := UpperCase(Trim(Qry.FieldByName('RT_Modal').AsString));
  finally
    Qry.Free;
  end;

  FCacheModal.AddOrSetValue(Chave, Result);
end;

function TPreparacaoSolver.ObterGrupoFisicoPlataforma(
  const APlataforma: string): string;
var
  Chave, GrupoFisico, NomeSAP: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCacheGrupoFisico.TryGetValue(Chave, Result) then
    Exit;

  Result := '';
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    // GrupoFisico é o campo novo adicionado na migração 1.7.0.4.
    // Fallback para NomeSAP que era o campo usado anteriormente na mesma função.
    Qry.SQL.Text :=
      'SELECT GrupoFisico, NomeSAP FROM tblPlataforma WHERE Plataforma = :Plat';
    Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
    Qry.Open;
    if not Qry.IsEmpty then
    begin
      GrupoFisico := Trim(Qry.FieldByName('GrupoFisico').AsString);
      NomeSAP     := Trim(Qry.FieldByName('NomeSAP').AsString);
      Result := IfThen(GrupoFisico <> '', GrupoFisico, NomeSAP);
    end;
  except
    // GrupoFisico pode não existir em banco não migrado — usa NomeSAP sozinho
    try
      Qry.Close;
      Qry.SQL.Text := 'SELECT NomeSAP FROM tblPlataforma WHERE Plataforma = :Plat';
      Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
      Qry.Open;
      if not Qry.IsEmpty then
        Result := Trim(Qry.FieldByName('NomeSAP').AsString);
    except
    end;
  end;

  FCacheGrupoFisico.AddOrSetValue(Chave, Result);
end;

function TPreparacaoSolver.ObterExcluirManualPlataforma(
  const APlataforma: string): Boolean;
var
  Chave: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCacheExcluirManual.TryGetValue(Chave, Result) then
    Exit;

  Result := False;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT ExcluirDistribuicaoAuto FROM tblPlataforma ' +
      'WHERE Plataforma = :Plat';
    Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
    Qry.Open;
    if not Qry.IsEmpty then
      Result := Qry.FieldByName('ExcluirDistribuicaoAuto').AsBoolean;
  except
    // Campo ainda não existe em banco não migrado — ignora
  end;
  Qry.Free;

  FCacheExcluirManual.AddOrSetValue(Chave, Result);
end;

function TPreparacaoSolver.ObterTipoNoLogistico(
  const APlataforma: string): string;
var
  Chave, TipoBanco, Modal: string;
  EhOrigem: Boolean;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCacheTipoNoLogistico.TryGetValue(Chave, Result) then
    Exit;

  TipoBanco := '';
  Modal     := '';
  EhOrigem  := False;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    // Lê TipoNoLogistico (campo novo), RT_Modal e booleanOrigem em uma consulta só.
    // TipoNoLogistico pode não existir em banco não migrado — trata no except.
    try
      Qry.SQL.Text :=
        'SELECT TipoNoLogistico, RT_Modal, booleanOrigem ' +
        'FROM tblPlataforma WHERE Plataforma = :Plat';
      Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
      Qry.Open;
      if not Qry.IsEmpty then
      begin
        TipoBanco := UpperCase(Trim(Qry.FieldByName('TipoNoLogistico').AsString));
        Modal     := UpperCase(Trim(Qry.FieldByName('RT_Modal').AsString));
        EhOrigem  := Qry.FieldByName('booleanOrigem').AsBoolean;
      end;
    except
      // TipoNoLogistico ainda não existe — lê apenas os campos que já existem
      try
        Qry.Close;
        Qry.SQL.Text :=
          'SELECT RT_Modal, booleanOrigem FROM tblPlataforma WHERE Plataforma = :Plat';
        Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
        Qry.Open;
        if not Qry.IsEmpty then
        begin
          Modal    := UpperCase(Trim(Qry.FieldByName('RT_Modal').AsString));
          EhOrigem := Qry.FieldByName('booleanOrigem').AsBoolean;
        end;
      except
      end;
    end;
  finally
    Qry.Free;
  end;

  if TipoBanco <> '' then
    // Campo preenchido no cadastro — usa sem fallback
    Result := TipoBanco
  else
  begin
    // Campo vazio ou ausente: inferência estrutural
    //   RT_Modal = T → terminal terrestre (TMIB e similares)
    //   booleanOrigem = True e não é terminal → hub de distribuição
    // Nota: este fallback é uma aproximação. Preencher TipoNoLogistico
    //       no cadastro elimina a ambiguidade.
    if Modal = 'T' then
      Result := TIPO_NO_TERMINAL
    else if EhOrigem then
      Result := TIPO_NO_HUB
    else
      Result := '';
  end;

  FCacheTipoNoLogistico.AddOrSetValue(Chave, Result);
end;

function TPreparacaoSolver.ObterPrioridadePlataforma(
  const APlataforma: string): Integer;
var
  Chave: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCachePrioridade.TryGetValue(Chave, Result) then
    Exit;

  Result := 99;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT PrioridadeDistribuicao FROM tblPlataforma ' +
      'WHERE Plataforma = :Plat';
    Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
    Qry.Open;
    if not Qry.IsEmpty then
      if not Qry.FieldByName('PrioridadeDistribuicao').IsNull then
        Result := Qry.FieldByName('PrioridadeDistribuicao').AsInteger;
  except
    // Campo ainda não existe em banco não migrado — mantém 99
  end;
  Qry.Free;

  FCachePrioridade.AddOrSetValue(Chave, Result);
end;

function TPreparacaoSolver.EhOrigemBridge(const AOrigem: string): Boolean;
var
  Chave: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(AOrigem));
  if FCacheOrigemBridge.TryGetValue(Chave, Result) then
    Exit;

  Result := False;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT TOP 1 NomeEmbarcacao FROM tblEmbarcacao ' +
      'WHERE UsaBridgeMesmoGrupo = True ' +
      '  AND (NomeEmbarcacao = :Orig OR NomeEmbarcacao = :OrigBridge)';
    Qry.Parameters.ParamByName('Orig').Value       := Trim(AOrigem);
    Qry.Parameters.ParamByName('OrigBridge').Value := Trim(AOrigem) + ' BRIDGE';
    Qry.Open;
    Result := not Qry.IsEmpty;
  finally
    Qry.Free;
  end;

  FCacheOrigemBridge.AddOrSetValue(Chave, Result);
end;

// =============================================================================
// HELPERS SEM BANCO
// =============================================================================

function TPreparacaoSolver.ExtrairTurnoPlataforma(
  const APlataforma: string): string;
var
  S: string;
begin
  S := UpperCase(Trim(APlataforma));
  if Pos('(D)', S) > 0 then Exit('D');
  if Pos('(N)', S) > 0 then Exit('N');
  Result := '';
end;

function TPreparacaoSolver.DerivarCodigoNorm(
  const ANomePlataforma: string): string;
var
  S: string;
  Num: Integer;
begin
  // Remove sufixo de turno, normaliza e tenta derivar o código do solver
  S := UpperCase(Trim(ANomePlataforma));
  S := StringReplace(S, ' (D)', '', [rfReplaceAll]);
  S := StringReplace(S, ' (N)', '', [rfReplaceAll]);
  S := Trim(S);

  // Já no formato aceito pelo solver
  if TRegEx.IsMatch(S, '^(PCM|PCB|PGA|PRB|PDO)-\d{2}$') then
    Exit(S);

  if S = 'TMIB' then
    Exit('TMIB');

  // MX → PCM-0X  (ex: M1 → PCM-01, M11 → PCM-11)
  if TRegEx.IsMatch(S, '^M\d+$') then
    if TryStrToInt(Copy(S, 2, MaxInt), Num) then
      Exit(Format('PCM-%2.2d', [Num]));

  // BX → PCB-0X
  if TRegEx.IsMatch(S, '^B\d+$') then
    if TryStrToInt(Copy(S, 2, MaxInt), Num) then
      Exit(Format('PCB-%2.2d', [Num]));

  // PGAX → PGA-0X
  if TRegEx.IsMatch(S, '^PGA\d+$') then
    if TryStrToInt(Copy(S, 4, MaxInt), Num) then
      Exit(Format('PGA-%2.2d', [Num]));

  // PDOX → PDO-0X
  if TRegEx.IsMatch(S, '^PDO\d+$') then
    if TryStrToInt(Copy(S, 4, MaxInt), Num) then
      Exit(Format('PDO-%2.2d', [Num]));

  // PRBX → PRB-0X
  if TRegEx.IsMatch(S, '^PRB\d+$') then
    if TryStrToInt(Copy(S, 4, MaxInt), Num) then
      Exit(Format('PRB-%2.2d', [Num]));

  Result := '';
end;

function TPreparacaoSolver.TryObterCodigoNormSolver(const APlataforma: string;
  out ACodigo: string): Boolean;
var
  Chave, CodBanco: string;
  Qry: TADOQuery;
begin
  Chave := UpperCase(Trim(APlataforma));
  if FCacheCodigoNorm.TryGetValue(Chave, ACodigo) then
    Exit(ACodigo <> '');

  CodBanco := '';
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT CodigoNormSolver FROM tblPlataforma WHERE Plataforma = :Plat';
    Qry.Parameters.ParamByName('Plat').Value := Trim(APlataforma);
    Qry.Open;
    if not Qry.IsEmpty then
      CodBanco := Trim(Qry.FieldByName('CodigoNormSolver').AsString);
  except
    // Campo ainda não existe em banco não migrado — vai para derivação
  end;
  Qry.Free;

  if CodBanco <> '' then
    ACodigo := CodBanco
  else
    ACodigo := DerivarCodigoNorm(APlataforma);

  FCacheCodigoNorm.AddOrSetValue(Chave, ACodigo);
  Result := ACodigo <> '';
end;

function TPreparacaoSolver.EhExecutanteBridge(
  const AOrigem, ADestino: string): Boolean;
var
  GrupoOrigem, GrupoDestino, TurnoDestino: string;
begin
  Result := False;
  if not EhOrigemBridge(AOrigem) then
    Exit;

  GrupoOrigem  := ObterGrupoFisicoPlataforma(AOrigem);
  GrupoDestino := ObterGrupoFisicoPlataforma(ADestino);

  // Somente é bridge se pertence ao mesmo grupo físico
  if not SameText(GrupoOrigem, GrupoDestino) or (GrupoOrigem = '') then
    Exit;

  TurnoDestino := ExtrairTurnoPlataforma(ADestino);
  Result := (TurnoDestino = 'D') or (TurnoDestino = 'N');
end;

function TPreparacaoSolver.DeveConsiderarDistribuicaoMaritima(
  const AOrigem, ADestino: string): Boolean;
var
  MO, MD: string;
begin
  MO := ObterModalPlataforma(AOrigem);
  MD := ObterModalPlataforma(ADestino);

  // Qualquer lado aéreo → não é marítima
  if (MO = 'A') or (MD = 'A') then
    Exit(False);

  // T+T → transbordo puro, não entra na distribuição marítima
  if (MO = 'T') and (MD = 'T') then
    Exit(False);

  // Aceita M+M, T+M, M+T
  Result :=
    ((MO = 'M') and (MD = 'M')) or
    ((MO = 'T') and (MD = 'M')) or
    ((MO = 'M') and (MD = 'T'));
end;

function TPreparacaoSolver.ResolverOrigemSolver(const AOrigem: string): string;
var
  TipoNo: string;
begin
  // Identificação baseada no papel logístico do cadastro, não no nome literal.
  // TERMINAL → nó terrestre de apoio (ex: TMIB) → OrigemSolver 'TMIB'
  // HUB      → plataforma de concentração/distribuição (ex: PCM-09) → 'M9'
  // qualquer outro → 'OUTRO' (não suportado como origem pelo solver)
  TipoNo := ObterTipoNoLogistico(AOrigem);
  if TipoNo = TIPO_NO_TERMINAL then
    Result := 'TMIB'
  else if TipoNo = TIPO_NO_HUB then
    Result := 'M9'
  else
    Result := 'OUTRO';
end;

// =============================================================================
// MÉTODOS PÚBLICOS
// =============================================================================

procedure TPreparacaoSolver.FiltrarExecutantes(
  const ADataProg: string;
  const AGrupoHorarioFiltro: string;
  out AElegiveis: TArray<TExecutanteElegivel>;
  out AExcluidos: TArray<TExecutanteExcluido>);
var
  Qry: TADOQuery;
  ListaEleg: TList<TExecutanteElegivel>;
  ListaExcl: TList<TExecutanteExcluido>;
  Eleg: TExecutanteElegivel;
  Excl: TExecutanteExcluido;
  IdExec: Integer;
  Origem, Destino, OrigemSolver, CodigoNorm, GrupoHorario: string;
begin
  LimparCaches;

  ListaEleg := TList<TExecutanteElegivel>.Create;
  ListaExcl := TList<TExecutanteExcluido>.Create;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT pe.idProgramacaoExecutante, pe.Origem, pd.txtDestino AS Destino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ' +
      '  ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      '  AND (pe.InseridoProgramacaoTransporte = False ' +
      '       OR pe.InseridoProgramacaoTransporte IS NULL) ' +
      'ORDER BY pe.idProgramacaoExecutante';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Open;

    while not Qry.Eof do
    begin
      IdExec  := Qry.FieldByName('idProgramacaoExecutante').AsInteger;
      Origem  := Trim(Qry.FieldByName('Origem').AsString);
      Destino := Trim(Qry.FieldByName('Destino').AsString);

      // ----- Destino vazio -----
      if Destino = '' then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_DESTINO_VAZIO;
        Excl.DescricaoMotivo := 'Destino não informado na programação.';
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- ExcluirDistribuicaoAuto -----
      if ObterExcluirManualPlataforma(Destino) then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_EXCLUIDO_MANUAL;
        Excl.DescricaoMotivo := Format(
          'Plataforma %s marcada para exclusão manual da distribuição automática.',
          [Destino]);
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- Modal marítimo -----
      if not DeveConsiderarDistribuicaoMaritima(Origem, Destino) then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_MODAL_AEREO;
        Excl.DescricaoMotivo := Format(
          'Combinação %s → %s não é distribuição marítima (modal aéreo ou T+T).',
          [Origem, Destino]);
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- Bridge / SOV -----
      if EhExecutanteBridge(Origem, Destino) then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_BRIDGE_SOV;
        Excl.DescricaoMotivo := Format(
          'Rota BRIDGE: %s → %s. Será processada pelo fluxo bridge.',
          [Origem, Destino]);
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- Origem não suportada -----
      OrigemSolver := ResolverOrigemSolver(Origem);
      if OrigemSolver = 'OUTRO' then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_ORIGEM_NAO_SUPORTADA;
        Excl.DescricaoMotivo := Format(
          'Origem "%s" não suportada pelo solver. Apenas TMIB e M9.',
          [Origem]);
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- CodigoNormSolver ausente -----
      if not TryObterCodigoNormSolver(Destino, CodigoNorm) then
      begin
        Excl := Default(TExecutanteExcluido);
        Excl.IdProgramacaoExecutante := IdExec;
        Excl.Origem          := Origem;
        Excl.Destino         := Destino;
        Excl.MotivoExclusao  := MOTIVO_SEM_CODIGO_SOLVER;
        Excl.DescricaoMotivo := Format(
          'Plataforma "%s" sem código normalizado (campo CodigoNormSolver vazio e nome não derivável).',
          [Destino]);
        ListaExcl.Add(Excl);
        Qry.Next;
        Continue;
      end;

      // ----- Filtro de turno (se solicitado) -----
      GrupoHorario := ExtrairTurnoPlataforma(Destino);
      if (AGrupoHorarioFiltro <> '') and (GrupoHorario <> '') and
         (not SameText(GrupoHorario, AGrupoHorarioFiltro)) then
      begin
        // Turno diferente do filtro: não inclui e não registra como excluído
        // (aparecerá em outro grupo)
        Qry.Next;
        Continue;
      end;

      // ----- Elegível -----
      Eleg := Default(TExecutanteElegivel);
      Eleg.IdProgramacaoExecutante := IdExec;
      Eleg.Origem                  := Origem;
      Eleg.OrigemSolver            := OrigemSolver;
      Eleg.Destino                 := Destino;
      Eleg.CodigoNormSolver        := CodigoNorm;
      Eleg.GrupoHorario            := GrupoHorario;
      Eleg.PrioridadeDistribuicao  := ObterPrioridadePlataforma(Destino);
      ListaEleg.Add(Eleg);

      Qry.Next;
    end;

    AElegiveis := ListaEleg.ToArray;
    AExcluidos := ListaExcl.ToArray;
  finally
    ListaEleg.Free;
    ListaExcl.Free;
    Qry.Free;
  end;
end;

function TPreparacaoSolver.AgregarDemanda(
  const AElegiveis: TArray<TExecutanteElegivel>): TArray<TDemandaSolver>;
var
  // Chave: CodigoNormSolver → índice em Lista
  Mapa: TDictionary<string, Integer>;
  Lista: TList<TDemandaSolver>;
  Eleg: TExecutanteElegivel;
  Dem: TDemandaSolver;
  Idx: Integer;
begin
  Mapa  := TDictionary<string, Integer>.Create;
  Lista := TList<TDemandaSolver>.Create;
  try
    for Eleg in AElegiveis do
    begin
      if Mapa.TryGetValue(Eleg.CodigoNormSolver, Idx) then
      begin
        Dem := Lista[Idx];
        if Eleg.OrigemSolver = 'TMIB' then
          Inc(Dem.PaxTMIB)
        else
          Inc(Dem.PaxM9);
        // Prioridade: guarda o menor número (maior urgência)
        if Eleg.PrioridadeDistribuicao < Dem.Prioridade then
          Dem.Prioridade := Eleg.PrioridadeDistribuicao;
        Lista[Idx] := Dem;
      end
      else
      begin
        Dem := Default(TDemandaSolver);
        Dem.CodigoPlataforma := Eleg.CodigoNormSolver;
        Dem.Prioridade       := Eleg.PrioridadeDistribuicao;
        if Eleg.OrigemSolver = 'TMIB' then
          Dem.PaxTMIB := 1
        else
          Dem.PaxM9 := 1;
        Mapa.Add(Eleg.CodigoNormSolver, Lista.Count);
        Lista.Add(Dem);
      end;
    end;
    Result := Lista.ToArray;
  finally
    Mapa.Free;
    Lista.Free;
  end;
end;

function TPreparacaoSolver.DetectarGruposHorario(
  const ADataProg: string): TArray<string>;
var
  Qry: TADOQuery;
  Grupos: TList<string>;
  Turno: string;
  TemSemTurno, TemD, TemN: Boolean;
begin
  TemSemTurno := False;
  TemD        := False;
  TemN        := False;

  Qry    := TADOQuery.Create(nil);
  Grupos := TList<string>.Create;
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT DISTINCT pd.txtDestino ' +
      'FROM tblProgramacaoExecutante pe ' +
      'INNER JOIN tblProgramacaoDiaria pd ' +
      '  ON pe.CodigoProgramacaoDiaria = pd.idProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      '  AND pd.txtDestino IS NOT NULL ' +
      '  AND (pe.InseridoProgramacaoTransporte = False ' +
      '       OR pe.InseridoProgramacaoTransporte IS NULL)';
    Qry.Parameters.ParamByName('DataProg').Value := ADataProg;
    Qry.Open;

    while not Qry.Eof do
    begin
      Turno := ExtrairTurnoPlataforma(Trim(Qry.FieldByName('txtDestino').AsString));
      if    Turno = 'D' then TemD        := True
      else if Turno = 'N' then TemN      := True
      else                    TemSemTurno := True;
      Qry.Next;
    end;

    if TemSemTurno or (not TemD and not TemN) then Grupos.Add('');
    if TemD  then Grupos.Add('D');
    if TemN  then Grupos.Add('N');

    Result := Grupos.ToArray;
  finally
    Qry.Free;
    Grupos.Free;
  end;
end;

function TPreparacaoSolver.CarregarListaEmbarcacoes: TArray<TBarcoSolver>;
var
  Qry: TADOQuery;
  Lista: TList<TBarcoSolver>;
  Barco: TBarcoSolver;
begin
  Lista := TList<TBarcoSolver>.Create;
  Qry   := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    // Distribuicao = campo adicionado na migração 1.7.0.4
    // Fallback: se o campo ainda não existir, carrega todas as ativas
    try
      Qry.SQL.Text :=
        'SELECT NomeEmbarcacao FROM tblEmbarcacao ' +
        'WHERE Distribuicao = True AND Ativo = True ' +
        'ORDER BY NomeEmbarcacao';
      Qry.Open;
    except
      Qry.Close;
      Qry.SQL.Text :=
        'SELECT NomeEmbarcacao FROM tblEmbarcacao ' +
        'WHERE Ativo = True ' +
        'ORDER BY NomeEmbarcacao';
      Qry.Open;
    end;

    while not Qry.Eof do
    begin
      Barco := Default(TBarcoSolver);
      Barco.NomeEmbarcacao := Trim(Qry.FieldByName('NomeEmbarcacao').AsString);
      Barco.Disponivel     := True;   // padrão: disponível
      Barco.HoraPartida    := '';     // preenchido pela interface operacional
      Barco.RotaFixa       := '';     // preenchido pela interface operacional
      Lista.Add(Barco);
      Qry.Next;
    end;

    Result := Lista.ToArray;
  finally
    Qry.Free;
    Lista.Free;
  end;
end;

function TPreparacaoSolver.PrepararInput(
  const ADataProg: string;
  ATrocaTurma: Boolean;
  ARendidosM9: Integer;
  const AGrupoHorarioFiltro: string): TDadosSolverInput;
var
  Elegiveis: TArray<TExecutanteElegivel>;
  Excluidos: TArray<TExecutanteExcluido>;
begin
  FiltrarExecutantes(ADataProg, AGrupoHorarioFiltro, Elegiveis, Excluidos);

  Result := Default(TDadosSolverInput);
  Result.DataOperacao       := StrToDateDef(ADataProg, Now);
  Result.TrocaTurma         := ATrocaTurma;
  Result.RendidosM9         := ARendidosM9;
  Result.GrupoHorarioFiltro := AGrupoHorarioFiltro;
  Result.Elegiveis          := Elegiveis;
  Result.Excluidos          := Excluidos;
  Result.Demandas           := AgregarDemanda(Elegiveis);
  Result.Barcos             := CarregarListaEmbarcacoes;
end;

procedure TPreparacaoSolver.GerarArquivoXlsx(
  const ADados: TDadosSolverInput;
  const ACaminhoArquivo: string);
var
  ExcelApp, WB, WS: OleVariant;
  R, I: Integer;
  Barco: TBarcoSolver;
  Dem: TDemandaSolver;
begin
  if Length(ADados.Barcos) = 0 then
    raise Exception.Create(
      'Nenhuma embarcação configurada para distribuição. ' +
      'Verifique o cadastro (campo Distribuicao).');

  if Length(ADados.Demandas) = 0 then
    raise Exception.Create(
      'Nenhuma demanda elegível para distribuição na data informada.');

  ExcelApp := CreateOleObject('Excel.Application');
  try
    ExcelApp.Visible        := False;
    ExcelApp.DisplayAlerts  := False;
    ExcelApp.ScreenUpdating := False;

    WB := ExcelApp.Workbooks.Add;
    WS := WB.Worksheets[1];
    WS.Name := 'solver_input';

    // -----------------------------------------------------------------------
    // Cabeçalho informativo (linhas 1-3 — não lidas pelo solver)
    // -----------------------------------------------------------------------
    WS.Cells[1, 2].Value := 'DISTRIBUIÇÃO LOGÍSTICA — SOLVER INPUT';
    WS.Cells[2, 2].Value := 'Data: ' + DateToStr(ADados.DataOperacao);
    WS.Cells[3, 2].Value := 'Gerado: ' + DateTimeToStr(Now);

    // -----------------------------------------------------------------------
    // Configuração (linhas 4-5 — lidas pelo solver nas colunas exatas)
    // -----------------------------------------------------------------------
    WS.Cells[SOLVER_ROW_TROCA_TURMA, 2].Value := 'Troca de turma:';
    WS.Cells[SOLVER_ROW_TROCA_TURMA, SOLVER_COL_BOAT_DISP].Value :=
      IfThen(ADados.TrocaTurma, 'SIM', 'NAO');

    WS.Cells[SOLVER_ROW_RENDIDOS_M9, 2].Value := 'Rendidos em M9:';
    WS.Cells[SOLVER_ROW_RENDIDOS_M9, SOLVER_COL_BOAT_DISP].Value :=
      ADados.RendidosM9;

    // -----------------------------------------------------------------------
    // Cabeçalho da seção de embarcações (linhas 7-8 — não lidas pelo solver)
    // -----------------------------------------------------------------------
    WS.Cells[7, 2].Value := 'EMBARCAÇÕES';
    WS.Cells[8, SOLVER_COL_BOAT_NOME].Value := 'Nome';
    WS.Cells[8, SOLVER_COL_BOAT_DISP].Value := 'Disponível';
    WS.Cells[8, SOLVER_COL_BOAT_HORA].Value := 'Hora Saída';
    WS.Cells[8, SOLVER_COL_BOAT_ROTA].Value := 'Rota Fixa';

    // -----------------------------------------------------------------------
    // Embarcações (a partir da linha SOLVER_FIRST_BOAT_ROW = 9)
    // O solver lê col B até encontrar célula vazia.
    // -----------------------------------------------------------------------
    R := SOLVER_FIRST_BOAT_ROW;
    for I := 0 to High(ADados.Barcos) do
    begin
      Barco := ADados.Barcos[I];
      WS.Cells[R, SOLVER_COL_BOAT_NOME].Value :=
        Barco.NomeEmbarcacao;
      WS.Cells[R, SOLVER_COL_BOAT_DISP].Value :=
        IfThen(Barco.Disponivel, 'SIM', 'NAO');
      WS.Cells[R, SOLVER_COL_BOAT_HORA].Value :=
        Barco.HoraPartida;
      WS.Cells[R, SOLVER_COL_BOAT_ROTA].Value :=
        Barco.RotaFixa;
      Inc(R);
    end;
    // R aponta para a primeira linha em branco após os barcos.
    // O solver para o loop de barcos aqui (célula B[R] vazia).

    // -----------------------------------------------------------------------
    // Seção de demanda
    // O solver calcula:
    //   demand_header_row = R + 1
    //   demand_col_row    = R + 2
    //   demand_start      = R + 3
    // -----------------------------------------------------------------------
    WS.Cells[R + 1, 2].Value := 'DEMANDA';
    WS.Cells[R + 2, SOLVER_COL_DEM_PLAT].Value := 'Plataforma';
    WS.Cells[R + 2, SOLVER_COL_DEM_M9  ].Value := 'PAX M9';
    WS.Cells[R + 2, SOLVER_COL_DEM_TMIB].Value := 'PAX TMIB';
    WS.Cells[R + 2, SOLVER_COL_DEM_PRIO].Value := 'Prioridade';

    R := R + 3;  // demand_start
    for I := 0 to High(ADados.Demandas) do
    begin
      Dem := ADados.Demandas[I];
      WS.Cells[R, SOLVER_COL_DEM_PLAT].Value := Dem.CodigoPlataforma;
      WS.Cells[R, SOLVER_COL_DEM_M9  ].Value := Dem.PaxM9;
      WS.Cells[R, SOLVER_COL_DEM_TMIB].Value := Dem.PaxTMIB;
      WS.Cells[R, SOLVER_COL_DEM_PRIO].Value := Dem.Prioridade;
      Inc(R);
    end;

    // Salva como .xlsx (51 = xlOpenXMLWorkbook)
    WB.SaveAs(ACaminhoArquivo, 51);
    WB.Close(False);
  finally
    try
      ExcelApp.Quit;
    except
    end;
    ExcelApp := Unassigned;
  end;
end;

end.
