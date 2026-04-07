unit uExecucaoSolver;

{
  Responsabilidade: executar o solver.py e persistir o resultado no banco.

  Fluxo:
    1. LerConfig             — lê PYTHON_EXE_PATH e SOLVER_WORK_PATH de tblDistribuicaoConfig.
    2. ExecutarPython         — chama python.exe solver.py via CreateProcess com timeout de 120 s.
    3. ParsearArquivoOutput  — lê distribuicao.txt e constrói TArray<TRotaSolverParsed>.
    4. MapearExecutantes      — distribui os TExecutanteElegivel pelas paradas de cada rota.
    5. PersistirResultado     — grava tblOperacaoDistribuicao, tblDistribuicaoRota,
                                tblDistribuicaoPaxAlocado, tblDistribuicaoExclusao,
                                tblRoteamento e tblAux_Rota_Distribuicao.

  Formato de distribuicao.txt (linha de dados por embarcacao):
    "<NomeEmbarcacao>  <HH:MM>  <RotaString>"
  RotaString e "/"-separada:
    "TMIB +N/M1 -X/M9 -Y +Z/M2 -W (-V)"
    - +N  -> N pax embarcando nesta parada
    - -N  -> N pax de TMIB desembarcando aqui
    - (-N) -> N pax de M9 desembarcando aqui
}

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections,
  System.RegularExpressions, System.StrUtils, System.DateUtils,
  Data.DB, Data.Win.ADODB,
  Winapi.Windows,
  uPreparacaoSolver;

type
  // Representa uma parada individual dentro de uma rota do solver
  TParadaSolver = record
    NomeCurto: string;           // Nome retornado pelo solver: 'M1', 'M9', 'TMIB'
    CodigoNorm: string;          // Nome normalizado Colibri: 'PCM-01', 'PCM-09', 'TMIB'
    EhHubM9: Boolean;
    EhTerminal: Boolean;
    PaxEmbarca: Integer;         // +N
    PaxDesembarcaTMIB: Integer;  // -N
    PaxDesembarcaM9: Integer;    // (-N)
  end;

  // Rota completa parseada para uma embarcacao
  TRotaSolverParsed = record
    NomeEmbarcacao: string;
    HoraPartida: string;
    RotaStringOriginal: string;
    Paradas: TArray<TParadaSolver>;
    UsaHubM9: Boolean;
    TotalPaxTMIB: Integer;
    TotalPaxM9: Integer;
  end;

  // Vinculo executante -> rota (resultado do mapeamento)
  TExecAloc = record
    IdProgramacaoExecutante: Integer;
    NomeEmbarcacao: string;
    OrigemSolver: string;        // 'TMIB' ou 'M9'
    CodigoNorm: string;
    PosicaoNaSequencia: Integer;
  end;

  TExecucaoSolver = class
  private
    FConnColibri: TADOConnection;
    FConnConsulta: TADOConnection;

    function LerConfig(const AChave: string): string;

    function ExecutarPython(const APythonExe, AWorkPath: string;
      out AMensagem: string): Boolean;

    function NormalizarCodigoPlataforma(const ANomeCurto: string): string;
    function ParsearParada(const AToken: string): TParadaSolver;
    function ParsearRotaString(const ARotaString: string): TArray<TParadaSolver>;
    function ParsearArquivoOutput(const AArquivo: string): TArray<TRotaSolverParsed>;

    procedure MapearExecutantes(
      const AElegiveis: TArray<TExecutanteElegivel>;
      const ARotas: TArray<TRotaSolverParsed>;
      out AAlocs: TArray<TExecAloc>);

    function ObterCapacidadeEmbarcacao(const ANomeEmbarcacao: string): Integer;
    function SequenciaColibri(const AParadas: TArray<TParadaSolver>): string;
    function GerarNomeSolverRota(const ANomeEmbarcacao: string;
      const ADataOp: TDateTime): string;

    function InserirOperacao(
      const ADadosInput: TDadosSolverInput;
      const APerfil, AUsuario: string;
      const AArqInput, AArqOutput: string;
      const AStatus, AMsgStatus: string): Integer;

    procedure InserirExclusoes(AIdOperacao: Integer;
      const AExcluidos: TArray<TExecutanteExcluido>);

    function InserirRotaCompatibilidade(
      const ANomeRota, ANomeEmbarcacao: string;
      const ADataOp: TDateTime;
      const AHoraPartida, ASequencia: string;
      ACapacidade, AIdOperacao: Integer): Integer;

    procedure AlocarExecutanteCompatibilidade(AIdExecutante, AIdRota: Integer);

    function InserirDistribuicaoRota(
      AIdOperacao, AIdRoteamento: Integer;
      const ARota: TRotaSolverParsed): Integer;

    procedure InserirPaxAlocado(
      AIdDistribuicaoRota, AIdExecutante: Integer;
      const AOrigemSolver: string;
      APosicao: Integer);

    procedure PersistirResultado(
      const ADadosInput: TDadosSolverInput;
      const ARotas: TArray<TRotaSolverParsed>;
      const AAlocs: TArray<TExecAloc>;
      AIdOperacao: Integer);

    procedure AtualizarStatusOperacao(AIdOperacao: Integer;
      const AStatus, AMsgStatus: string);

  public
    constructor Create(AConnColibri, AConnConsulta: TADOConnection);
    destructor Destroy; override;

    function ExecutarDistribuicao(
      const ADadosInput: TDadosSolverInput;
      const APerfil, AUsuario: string;
      out AIdOperacao: Integer;
      out AMensagem: string): Boolean;
  end;

implementation

// =============================================================================
// Constructor / Destructor
// =============================================================================

constructor TExecucaoSolver.Create(AConnColibri, AConnConsulta: TADOConnection);
begin
  inherited Create;
  FConnColibri  := AConnColibri;
  FConnConsulta := AConnConsulta;
end;

destructor TExecucaoSolver.Destroy;
begin
  inherited;
end;

// =============================================================================
// Configuracao
// =============================================================================

function TExecucaoSolver.LerConfig(const AChave: string): string;
var
  Qry: TADOQuery;
begin
  Result := '';
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT ValorConfig FROM tblDistribuicaoConfig ' +
      'WHERE ChaveConfig = :Chave';
    Qry.Parameters.ParamByName('Chave').Value := AChave;
    try
      Qry.Open;
      if not Qry.IsEmpty then
        Result := Qry.FieldByName('ValorConfig').AsString;
    except
      // Tabela ausente em banco nao-migrado
    end;
  finally
    Qry.Free;
  end;
end;

// =============================================================================
// Execucao do processo Python
// =============================================================================

function TExecucaoSolver.ExecutarPython(const APythonExe, AWorkPath: string;
  out AMensagem: string): Boolean;
const
  TIMEOUT_MS = 120000;
var
  SI: TStartupInfo;
  PI: TProcessInformation;
  CmdLine: string;
  WaitResult: DWORD;
  ExitCode: DWORD;
begin
  Result := False;
  AMensagem := '';

  ZeroMemory(@SI, SizeOf(SI));
  SI.cb := SizeOf(SI);
  SI.dwFlags := STARTF_USESHOWWINDOW;
  SI.wShowWindow := SW_HIDE;
  ZeroMemory(@PI, SizeOf(PI));

  if Pos(' ', APythonExe) > 0 then
    CmdLine := '"' + APythonExe + '" solver.py'
  else
    CmdLine := APythonExe + ' solver.py';

  if not CreateProcess(
      nil, PChar(CmdLine), nil, nil, False,
      CREATE_NO_WINDOW, nil,
      PChar(AWorkPath), SI, PI) then
  begin
    AMensagem := 'Falha ao iniciar o processo Python: ' +
      SysErrorMessage(GetLastError);
    Exit;
  end;

  try
    WaitResult := WaitForSingleObject(PI.hProcess, TIMEOUT_MS);
    case WaitResult of
      WAIT_TIMEOUT:
      begin
        TerminateProcess(PI.hProcess, 1);
        AMensagem := Format('Timeout: solver.py excedeu %d segundos',
          [TIMEOUT_MS div 1000]);
      end;
      WAIT_OBJECT_0:
      begin
        GetExitCodeProcess(PI.hProcess, ExitCode);
        if ExitCode = 0 then
          Result := True
        else
          AMensagem := Format('solver.py retornou codigo de saida %d', [ExitCode]);
      end;
    else
      AMensagem := 'WaitForSingleObject retornou valor inesperado';
    end;
  finally
    CloseHandle(PI.hProcess);
    CloseHandle(PI.hThread);
  end;
end;

// =============================================================================
// Parsing do distribuicao.txt
// =============================================================================

function TExecucaoSolver.NormalizarCodigoPlataforma(
  const ANomeCurto: string): string;
var
  M: TMatch;
  Prefixo: string;
  Num: Integer;
begin
  Result := ANomeCurto;
  if (ANomeCurto = '') or (ANomeCurto = 'TMIB') then Exit;

  // Padrao: letras + numero -> PREFIX-NN
  // Ex: M1->PCM-01, M9->PCM-09, B1->PCB-01, PGA2->PGA-02
  M := TRegEx.Match(ANomeCurto, '^([A-Za-z]+)(\d+)$');
  if not M.Success then Exit;

  Prefixo := UpperCase(M.Groups[1].Value);
  Num := StrToIntDef(M.Groups[2].Value, 0);

  if      Prefixo = 'M'   then Result := Format('PCM-%2.2d', [Num])
  else if Prefixo = 'B'   then Result := Format('PCB-%2.2d', [Num])
  else if Prefixo = 'PGA' then Result := Format('PGA-%2.2d', [Num])
  else if Prefixo = 'PDO' then Result := Format('PDO-%2.2d', [Num])
  else if Prefixo = 'PRB' then Result := Format('PRB-%2.2d', [Num]);
end;

function TExecucaoSolver.ParsearParada(const AToken: string): TParadaSolver;
var
  SpacePos: Integer;
  OpsParte, OpsResto: string;
  M: TMatch;
begin
  Result := Default(TParadaSolver);

  SpacePos := Pos(' ', Trim(AToken));
  if SpacePos = 0 then
  begin
    Result.NomeCurto := Trim(AToken);
    OpsParte := '';
  end
  else
  begin
    Result.NomeCurto := Trim(Copy(AToken, 1, SpacePos - 1));
    OpsParte := Trim(Copy(AToken, SpacePos + 1, MaxInt));
  end;

  Result.CodigoNorm := NormalizarCodigoPlataforma(Result.NomeCurto);
  Result.EhTerminal := (Result.CodigoNorm = 'TMIB');
  Result.EhHubM9    := (Result.CodigoNorm = 'PCM-09');

  if OpsParte = '' then Exit;

  // (-N) -> pax de M9 desembarcando (com parenteses)
  M := TRegEx.Match(OpsParte, '\(-\s*(\d+)\s*\)');
  if M.Success then
  begin
    Result.PaxDesembarcaM9 := StrToIntDef(M.Groups[1].Value, 0);
    // Remove a ocorrencia (-N) para nao interferir no proximo match
    OpsResto := StringReplace(OpsParte, M.Value, '', []);
  end
  else
    OpsResto := OpsParte;

  // -N (sem parenteses) -> pax de TMIB desembarcando
  M := TRegEx.Match(OpsResto, '-\s*(\d+)');
  if M.Success then
    Result.PaxDesembarcaTMIB := StrToIntDef(M.Groups[1].Value, 0);

  // +N -> pax embarcando nesta parada
  M := TRegEx.Match(OpsParte, '\+\s*(\d+)');
  if M.Success then
    Result.PaxEmbarca := StrToIntDef(M.Groups[1].Value, 0);
end;

function TExecucaoSolver.ParsearRotaString(
  const ARotaString: string): TArray<TParadaSolver>;
var
  Tokens: TArray<string>;
  I: Integer;
  Paradas: TArray<TParadaSolver>;
begin
  Tokens := ARotaString.Split(['/']);
  SetLength(Paradas, Length(Tokens));
  for I := 0 to High(Tokens) do
    Paradas[I] := ParsearParada(Tokens[I]);
  Result := Paradas;
end;

function TExecucaoSolver.ParsearArquivoOutput(
  const AArquivo: string): TArray<TRotaSolverParsed>;
var
  Lines: TStringList;
  Line, NomeEmb, HoraP, RotaStr: string;
  Parts: TArray<string>;
  Rota: TRotaSolverParsed;
  Parada: TParadaSolver;
  Rotas: TList<TRotaSolverParsed>;
  I: Integer;
begin
  Result := [];
  if not FileExists(AArquivo) then Exit;

  Rotas := TList<TRotaSolverParsed>.Create;
  Lines := TStringList.Create;
  try
    Lines.LoadFromFile(AArquivo);

    for I := 0 to Lines.Count - 1 do
    begin
      Line := Trim(Lines[I]);

      // Pula cabecalho e linhas vazias
      if (Line = '') or
         StartsText('DISTRIBUICAO', Line) or
         StartsStr('=', Line) then
        Continue;

      // Formato: "NomeEmbarcacao  HH:MM  RotaString"
      // Dois espacos consecutivos como delimitador de campo
      Parts := Line.Split(['  '], 3);
      if Length(Parts) < 3 then Continue;

      NomeEmb := Trim(Parts[0]);
      HoraP   := Trim(Parts[1]);
      RotaStr := Trim(Parts[2]);

      if NomeEmb = '' then Continue;

      Rota := Default(TRotaSolverParsed);
      Rota.NomeEmbarcacao     := NomeEmb;
      Rota.HoraPartida        := HoraP;
      Rota.RotaStringOriginal := RotaStr;
      Rota.Paradas            := ParsearRotaString(RotaStr);

      for Parada in Rota.Paradas do
      begin
        Inc(Rota.TotalPaxTMIB, Parada.PaxDesembarcaTMIB);
        Inc(Rota.TotalPaxM9,   Parada.PaxDesembarcaM9);
        if Parada.EhHubM9 then
          Rota.UsaHubM9 := True;
      end;

      Rotas.Add(Rota);
    end;

    Result := Rotas.ToArray;
  finally
    Lines.Free;
    Rotas.Free;
  end;
end;

// =============================================================================
// Mapeamento executante -> parada
// =============================================================================

procedure TExecucaoSolver.MapearExecutantes(
  const AElegiveis: TArray<TExecutanteElegivel>;
  const ARotas: TArray<TRotaSolverParsed>;
  out AAlocs: TArray<TExecAloc>);
var
  // Chave: "CodigoNorm|OrigemSolver" -> fila FIFO de executantes elegiveis
  Filas: TDictionary<string, TQueue<TExecutanteElegivel>>;
  Elegivel: TExecutanteElegivel;
  Chave: string;
  Fila: TQueue<TExecutanteElegivel>;
  AllocList: TList<TExecAloc>;
  RotaIdx, ParadaIdx, PaxIdx: Integer;
  Rota: TRotaSolverParsed;
  Parada: TParadaSolver;
  Aloc: TExecAloc;
begin
  AAlocs := [];

  Filas := TDictionary<string, TQueue<TExecutanteElegivel>>.Create;
  try
    // Monta filas na ordem do array AElegiveis (preserva prioridade original)
    for Elegivel in AElegiveis do
    begin
      Chave := Elegivel.CodigoNormSolver + '|' + Elegivel.OrigemSolver;
      if not Filas.TryGetValue(Chave, Fila) then
      begin
        Fila := TQueue<TExecutanteElegivel>.Create;
        Filas.Add(Chave, Fila);
      end;
      Fila.Enqueue(Elegivel);
    end;

    AllocList := TList<TExecAloc>.Create;
    try
      for RotaIdx := 0 to High(ARotas) do
      begin
        Rota := ARotas[RotaIdx];
        for ParadaIdx := 0 to High(Rota.Paradas) do
        begin
          Parada := Rota.Paradas[ParadaIdx];

          // TMIB e M9 sao pontos de embarque, nao destinos de PAX final
          if Parada.EhTerminal or Parada.EhHubM9 then Continue;

          // Pax de TMIB desembarcando nesta plataforma
          if Parada.PaxDesembarcaTMIB > 0 then
          begin
            Chave := Parada.CodigoNorm + '|TMIB';
            if Filas.TryGetValue(Chave, Fila) then
            begin
              for PaxIdx := 1 to Parada.PaxDesembarcaTMIB do
              begin
                if Fila.Count = 0 then Break;
                Elegivel := Fila.Dequeue;

                Aloc := Default(TExecAloc);
                Aloc.IdProgramacaoExecutante := Elegivel.IdProgramacaoExecutante;
                Aloc.NomeEmbarcacao          := Rota.NomeEmbarcacao;
                Aloc.OrigemSolver            := 'TMIB';
                Aloc.CodigoNorm              := Parada.CodigoNorm;
                Aloc.PosicaoNaSequencia      := ParadaIdx;
                AllocList.Add(Aloc);
              end;
            end;
          end;

          // Pax de M9 desembarcando nesta plataforma
          if Parada.PaxDesembarcaM9 > 0 then
          begin
            Chave := Parada.CodigoNorm + '|M9';
            if Filas.TryGetValue(Chave, Fila) then
            begin
              for PaxIdx := 1 to Parada.PaxDesembarcaM9 do
              begin
                if Fila.Count = 0 then Break;
                Elegivel := Fila.Dequeue;

                Aloc := Default(TExecAloc);
                Aloc.IdProgramacaoExecutante := Elegivel.IdProgramacaoExecutante;
                Aloc.NomeEmbarcacao          := Rota.NomeEmbarcacao;
                Aloc.OrigemSolver            := 'M9';
                Aloc.CodigoNorm              := Parada.CodigoNorm;
                Aloc.PosicaoNaSequencia      := ParadaIdx;
                AllocList.Add(Aloc);
              end;
            end;
          end;
        end;
      end;

      AAlocs := AllocList.ToArray;
    finally
      AllocList.Free;
    end;
  finally
    for Fila in Filas.Values do
      Fila.Free;
    Filas.Free;
  end;
end;

// =============================================================================
// Persistencia - helpers
// =============================================================================

function TExecucaoSolver.ObterCapacidadeEmbarcacao(
  const ANomeEmbarcacao: string): Integer;
var
  Qry: TADOQuery;
begin
  Result := 0;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnConsulta;
    Qry.SQL.Text :=
      'SELECT CapacidadePAX FROM tblEmbarcacao WHERE NomeEmbarcacao = :Nome';
    Qry.Parameters.ParamByName('Nome').Value := ANomeEmbarcacao;
    try
      Qry.Open;
      if not Qry.IsEmpty then
        Result := Qry.FieldByName('CapacidadePAX').AsInteger;
    except
    end;
  finally
    Qry.Free;
  end;
end;

function TExecucaoSolver.SequenciaColibri(
  const AParadas: TArray<TParadaSolver>): string;
var
  Parts: TStringList;
  Parada: TParadaSolver;
begin
  // Formato Colibri: "TMIB;PCM-01;PCM-09;PCM-02" (separado por ';')
  Parts := TStringList.Create;
  try
    Parts.StrictDelimiter := True;
    Parts.Delimiter := ';';
    for Parada in AParadas do
      if Parada.CodigoNorm <> '' then
        Parts.Add(Parada.CodigoNorm);
    Result := Parts.DelimitedText;
  finally
    Parts.Free;
  end;
end;

function TExecucaoSolver.GerarNomeSolverRota(const ANomeEmbarcacao: string;
  const ADataOp: TDateTime): string;
var
  Qry: TADOQuery;
  QtdSolver: Integer;
begin
  // Formato: "SLV Nº {NomeEmbarcacao}" evita colisao com rotas manuais
  QtdSolver := 0;
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'SELECT COUNT(*) AS Total FROM tblRoteamento ' +
      'WHERE DataRoteamento = :DataOp AND NomeRota LIKE ''SLV%''';
    Qry.Parameters.ParamByName('DataOp').Value := DateToStr(ADataOp);
    try
      Qry.Open;
      if not Qry.IsEmpty then
        QtdSolver := Qry.FieldByName('Total').AsInteger;
    except
    end;
  finally
    Qry.Free;
  end;

  Result := Format('SLV %d° %s', [QtdSolver + 1, ANomeEmbarcacao]);
end;

procedure TExecucaoSolver.AtualizarStatusOperacao(AIdOperacao: Integer;
  const AStatus, AMsgStatus: string);
begin
  try
    FConnColibri.Execute(
      Format(
        'UPDATE tblOperacaoDistribuicao ' +
        'SET StatusExecucao = ''%s'', MensagemStatus = ''%s'' ' +
        'WHERE idOperacaoDistribuicao = %d',
        [AStatus,
         StringReplace(AMsgStatus, '''', '''''', [rfReplaceAll]),
         AIdOperacao]
      )
    );
  except
  end;
end;

// =============================================================================
// Persistencia - INSERTs individuais
// =============================================================================

function TExecucaoSolver.InserirOperacao(
  const ADadosInput: TDadosSolverInput;
  const APerfil, AUsuario: string;
  const AArqInput, AArqOutput: string;
  const AStatus, AMsgStatus: string): Integer;
var
  Qry, QryID: TADOQuery;
begin
  Result := 0;
  Qry   := TADOQuery.Create(nil);
  QryID := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblOperacaoDistribuicao ' +
      '(DataOperacao, Perfil, Versao, DataHoraExecucao, UsuarioExecucao, ' +
      ' StatusExecucao, MensagemStatus, TrocaTurma, RendidosM9, ' +
      ' ArquivoInput, ArquivoOutput) ' +
      'VALUES ' +
      '(:DataOp, :Perfil, :Versao, :DataHoraExec, :Usuario, ' +
      ' :Status, :MsgStatus, :TrocaTurma, :RendidosM9, ' +
      ' :ArqInput, :ArqOutput)';

    Qry.Parameters.ParamByName('DataOp').Value      := ADadosInput.DataOperacao;
    Qry.Parameters.ParamByName('Perfil').Value      := APerfil;
    Qry.Parameters.ParamByName('Versao').Value      := 1;
    Qry.Parameters.ParamByName('DataHoraExec').Value := Now;
    Qry.Parameters.ParamByName('Usuario').Value     := AUsuario;
    Qry.Parameters.ParamByName('Status').Value      := AStatus;
    Qry.Parameters.ParamByName('MsgStatus').Value   := AMsgStatus;
    Qry.Parameters.ParamByName('TrocaTurma').Value  := ADadosInput.TrocaTurma;
    Qry.Parameters.ParamByName('RendidosM9').Value  := ADadosInput.RendidosM9;
    Qry.Parameters.ParamByName('ArqInput').Value    := AArqInput;
    Qry.Parameters.ParamByName('ArqOutput').Value   := AArqOutput;

    Qry.ExecSQL;

    QryID.Connection := FConnColibri;
    QryID.SQL.Text := 'SELECT @@IDENTITY AS NovoID';
    QryID.Open;
    if not QryID.IsEmpty then
      Result := QryID.FieldByName('NovoID').AsInteger;
  finally
    Qry.Free;
    QryID.Free;
  end;
end;

procedure TExecucaoSolver.InserirExclusoes(AIdOperacao: Integer;
  const AExcluidos: TArray<TExecutanteExcluido>);
var
  Qry: TADOQuery;
  Excluido: TExecutanteExcluido;
begin
  if Length(AExcluidos) = 0 then Exit;

  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblDistribuicaoExclusao ' +
      '(idOperacaoDistribuicao, idProgramacaoExecutante, ' +
      ' MotivoExclusao, DescricaoMotivo) ' +
      'VALUES (:IdOp, :IdExec, :Motivo, :Descricao)';

    for Excluido in AExcluidos do
    begin
      Qry.Parameters.ParamByName('IdOp').Value      := AIdOperacao;
      Qry.Parameters.ParamByName('IdExec').Value    := Excluido.IdProgramacaoExecutante;
      Qry.Parameters.ParamByName('Motivo').Value    := Excluido.MotivoExclusao;
      Qry.Parameters.ParamByName('Descricao').Value := Excluido.DescricaoMotivo;
      try
        Qry.ExecSQL;
      except
      end;
    end;
  finally
    Qry.Free;
  end;
end;

function TExecucaoSolver.InserirRotaCompatibilidade(
  const ANomeRota, ANomeEmbarcacao: string;
  const ADataOp: TDateTime;
  const AHoraPartida, ASequencia: string;
  ACapacidade, AIdOperacao: Integer): Integer;
var
  Qry, QryID: TADOQuery;
begin
  Result := 0;
  Qry   := TADOQuery.Create(nil);
  QryID := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblRoteamento ' +
      '(NomeRota, NomeEmbarcacao, DataRoteamento, ' +
      ' HoraRoteamento, RotaSequencia, CapacidadePAX) ' +
      'VALUES ' +
      '(:NomeRota, :Embarcacao, :DataOp, :Hora, :Sequencia, :Capacidade)';

    Qry.Parameters.ParamByName('NomeRota').Value   := ANomeRota;
    Qry.Parameters.ParamByName('Embarcacao').Value := ANomeEmbarcacao;
    Qry.Parameters.ParamByName('DataOp').Value     := DateToStr(ADataOp);
    Qry.Parameters.ParamByName('Hora').Value       := AHoraPartida;
    Qry.Parameters.ParamByName('Sequencia').Value  := ASequencia;
    Qry.Parameters.ParamByName('Capacidade').Value := ACapacidade;

    Qry.ExecSQL;

    QryID.Connection := FConnColibri;
    QryID.SQL.Text := 'SELECT @@IDENTITY AS NovoID';
    QryID.Open;
    if not QryID.IsEmpty then
      Result := QryID.FieldByName('NovoID').AsInteger;

    // Atualiza campos extras adicionados na migracao 1.7.0.4 (tolerante a falha)
    if Result > 0 then
    begin
      try
        FConnColibri.Execute(
          Format(
            'UPDATE tblRoteamento ' +
            'SET OrigemDistribuicao = ''SOLVER'', GeradaPorSolver = True, ' +
            '    IdOperacaoDistribuicao = %d ' +
            'WHERE idRoteamento = %d',
            [AIdOperacao, Result]
          )
        );
      except
      end;
    end;
  finally
    Qry.Free;
    QryID.Free;
  end;
end;

procedure TExecucaoSolver.AlocarExecutanteCompatibilidade(
  AIdExecutante, AIdRota: Integer);
var
  Qry: TADOQuery;
begin
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblAux_Rota_Distribuicao ' +
      '(CodigoProgramacaoExecutante, CodigoRota) ' +
      'VALUES (:IdExecutante, :IdRota)';
    Qry.Parameters.ParamByName('IdExecutante').Value := AIdExecutante;
    Qry.Parameters.ParamByName('IdRota').Value       := AIdRota;
    try
      Qry.ExecSQL;
    except
    end;
  finally
    Qry.Free;
  end;
end;

function TExecucaoSolver.InserirDistribuicaoRota(
  AIdOperacao, AIdRoteamento: Integer;
  const ARota: TRotaSolverParsed): Integer;
var
  Qry, QryID: TADOQuery;
begin
  Result := 0;
  Qry   := TADOQuery.Create(nil);
  QryID := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblDistribuicaoRota ' +
      '(idOperacaoDistribuicao, idRoteamento, NomeEmbarcacao, HoraPartida, ' +
      ' SequenciaSolver, DistanciaNM, UsaHubM9, PaxTMIB, PaxM9, RotaFixa) ' +
      'VALUES ' +
      '(:IdOp, :IdRot, :NomeEmb, :Hora, :Seq, :DistNM, :UsaM9, :PaxT, :PaxM, :RotaFixa)';

    Qry.Parameters.ParamByName('IdOp').Value     := AIdOperacao;
    Qry.Parameters.ParamByName('IdRot').Value    := AIdRoteamento;
    Qry.Parameters.ParamByName('NomeEmb').Value  := ARota.NomeEmbarcacao;
    Qry.Parameters.ParamByName('Hora').Value     := ARota.HoraPartida;
    Qry.Parameters.ParamByName('Seq').Value      := ARota.RotaStringOriginal;
    Qry.Parameters.ParamByName('DistNM').Value   := 0.0;
    Qry.Parameters.ParamByName('UsaM9').Value    := ARota.UsaHubM9;
    Qry.Parameters.ParamByName('PaxT').Value     := ARota.TotalPaxTMIB;
    Qry.Parameters.ParamByName('PaxM').Value     := ARota.TotalPaxM9;
    Qry.Parameters.ParamByName('RotaFixa').Value := False;

    Qry.ExecSQL;

    QryID.Connection := FConnColibri;
    QryID.SQL.Text := 'SELECT @@IDENTITY AS NovoID';
    QryID.Open;
    if not QryID.IsEmpty then
      Result := QryID.FieldByName('NovoID').AsInteger;
  finally
    Qry.Free;
    QryID.Free;
  end;
end;

procedure TExecucaoSolver.InserirPaxAlocado(
  AIdDistribuicaoRota, AIdExecutante: Integer;
  const AOrigemSolver: string;
  APosicao: Integer);
var
  Qry: TADOQuery;
begin
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FConnColibri;
    Qry.SQL.Text :=
      'INSERT INTO tblDistribuicaoPaxAlocado ' +
      '(idDistribuicaoRota, idProgramacaoExecutante, OrigemSolver, ' +
      ' PosicaoNaSequencia, DataHoraChegadaEstimada) ' +
      'VALUES (:IdRota, :IdExec, :Origem, :Posicao, NULL)';
    Qry.Parameters.ParamByName('IdRota').Value  := AIdDistribuicaoRota;
    Qry.Parameters.ParamByName('IdExec').Value  := AIdExecutante;
    Qry.Parameters.ParamByName('Origem').Value  := AOrigemSolver;
    Qry.Parameters.ParamByName('Posicao').Value := APosicao;
    try
      Qry.ExecSQL;
    except
    end;
  finally
    Qry.Free;
  end;
end;

// =============================================================================
// Persistencia - processo completo
// =============================================================================

procedure TExecucaoSolver.PersistirResultado(
  const ADadosInput: TDadosSolverInput;
  const ARotas: TArray<TRotaSolverParsed>;
  const AAlocs: TArray<TExecAloc>;
  AIdOperacao: Integer);
var
  RotaIdx: Integer;
  Rota: TRotaSolverParsed;
  Aloc: TExecAloc;
  IdRoteamento, IdDistribuicaoRota: Integer;
  NomeRota, Sequencia: string;
  Capacidade: Integer;
  // Mapa NomeEmbarcacao -> IdRoteamento (para AlocarExecutanteCompatibilidade)
  RoteamentoPorEmb: TDictionary<string, Integer>;
  // Mapa NomeEmbarcacao -> IdDistribuicaoRota (para InserirPaxAlocado)
  DistribuicaoPorEmb: TDictionary<string, Integer>;
begin
  RoteamentoPorEmb   := TDictionary<string, Integer>.Create;
  DistribuicaoPorEmb := TDictionary<string, Integer>.Create;
  try
    // Passo 1: cria as rotas (tblRoteamento + tblDistribuicaoRota)
    for RotaIdx := 0 to High(ARotas) do
    begin
      Rota := ARotas[RotaIdx];

      Capacidade := ObterCapacidadeEmbarcacao(Rota.NomeEmbarcacao);
      NomeRota   := GerarNomeSolverRota(Rota.NomeEmbarcacao, ADadosInput.DataOperacao);
      Sequencia  := SequenciaColibri(Rota.Paradas);

      // Rota de compatibilidade (tabela existente)
      IdRoteamento := InserirRotaCompatibilidade(
        NomeRota,
        Rota.NomeEmbarcacao,
        ADadosInput.DataOperacao,
        Rota.HoraPartida,
        Sequencia,
        Capacidade,
        AIdOperacao
      );

      // Rota nova (tabela do solver)
      IdDistribuicaoRota := InserirDistribuicaoRota(AIdOperacao, IdRoteamento, Rota);

      RoteamentoPorEmb.AddOrSetValue(Rota.NomeEmbarcacao, IdRoteamento);
      DistribuicaoPorEmb.AddOrSetValue(Rota.NomeEmbarcacao, IdDistribuicaoRota);
    end;

    // Passo 2: associa executantes (tblDistribuicaoPaxAlocado + tblAux_Rota_Distribuicao)
    for Aloc in AAlocs do
    begin
      if DistribuicaoPorEmb.TryGetValue(Aloc.NomeEmbarcacao, IdDistribuicaoRota) then
        InserirPaxAlocado(
          IdDistribuicaoRota,
          Aloc.IdProgramacaoExecutante,
          Aloc.OrigemSolver,
          Aloc.PosicaoNaSequencia
        );

      if RoteamentoPorEmb.TryGetValue(Aloc.NomeEmbarcacao, IdRoteamento) then
        AlocarExecutanteCompatibilidade(Aloc.IdProgramacaoExecutante, IdRoteamento);
    end;
  finally
    RoteamentoPorEmb.Free;
    DistribuicaoPorEmb.Free;
  end;
end;

// =============================================================================
// Ponto de entrada principal
// =============================================================================

function TExecucaoSolver.ExecutarDistribuicao(
  const ADadosInput: TDadosSolverInput;
  const APerfil, AUsuario: string;
  out AIdOperacao: Integer;
  out AMensagem: string): Boolean;
var
  PythonExe, WorkPath: string;
  ArqInput, ArqOutput: string;
  Rotas: TArray<TRotaSolverParsed>;
  Alocs: TArray<TExecAloc>;
  MsgPython: string;
begin
  Result      := False;
  AIdOperacao := 0;
  AMensagem   := '';

  // 1. Configuracao
  PythonExe := LerConfig('PYTHON_EXE_PATH');
  WorkPath  := LerConfig('SOLVER_WORK_PATH');

  if PythonExe = '' then PythonExe := 'python.exe';

  if WorkPath = '' then
  begin
    AMensagem :=
      'SOLVER_WORK_PATH nao configurado. ' +
      'Configure a pasta de trabalho do solver em Configuracoes.';
    Exit;
  end;

  if not DirectoryExists(WorkPath) then
  begin
    AMensagem := 'Pasta de trabalho do solver nao encontrada: ' + WorkPath;
    Exit;
  end;

  ArqInput  := IncludeTrailingPathDelimiter(WorkPath) + 'solver_input.xlsx';
  ArqOutput := IncludeTrailingPathDelimiter(WorkPath) + 'distribuicao.txt';

  // 2. Verifica se o arquivo de entrada foi gerado
  if not FileExists(ArqInput) then
  begin
    AMensagem := 'solver_input.xlsx nao encontrado em: ' + WorkPath +
      '. Execute a preparacao antes de distribuir.';
    Exit;
  end;

  // 3. Registro inicial da operacao (status EXECUTANDO)
  try
    AIdOperacao := InserirOperacao(
      ADadosInput, APerfil, AUsuario,
      ArqInput, ArqOutput,
      'EXECUTANDO', '');
  except
    on E: Exception do
    begin
      AMensagem := 'Erro ao registrar operacao no banco: ' + E.Message;
      Exit;
    end;
  end;

  // 4. Persiste exclusoes da distribuicao
  InserirExclusoes(AIdOperacao, ADadosInput.Excluidos);

  // 5. Executa o solver Python
  if not ExecutarPython(PythonExe, WorkPath, MsgPython) then
  begin
    AtualizarStatusOperacao(AIdOperacao, 'ERRO', MsgPython);
    AMensagem := MsgPython;
    Exit;
  end;

  // 6. Verifica arquivo de saida
  if not FileExists(ArqOutput) then
  begin
    AMensagem := 'solver.py concluiu mas distribuicao.txt nao foi gerado em: ' + WorkPath;
    AtualizarStatusOperacao(AIdOperacao, 'ERRO', AMensagem);
    Exit;
  end;

  // 7. Parseia o resultado
  try
    Rotas := ParsearArquivoOutput(ArqOutput);
  except
    on E: Exception do
    begin
      AMensagem := 'Erro ao ler distribuicao.txt: ' + E.Message;
      AtualizarStatusOperacao(AIdOperacao, 'ERRO', AMensagem);
      Exit;
    end;
  end;

  if Length(Rotas) = 0 then
  begin
    AMensagem := 'distribuicao.txt nao contem rotas validas.';
    AtualizarStatusOperacao(AIdOperacao, 'SEM_ROTAS', AMensagem);
    Exit;
  end;

  // 8. Mapeia executantes para as paradas de cada rota
  MapearExecutantes(ADadosInput.Elegiveis, Rotas, Alocs);

  // 9. Persiste no banco
  try
    PersistirResultado(ADadosInput, Rotas, Alocs, AIdOperacao);
  except
    on E: Exception do
    begin
      AMensagem := 'Erro ao persistir resultado: ' + E.Message;
      AtualizarStatusOperacao(AIdOperacao, 'ERRO_PERSIST', AMensagem);
      Exit;
    end;
  end;

  // 10. Atualiza status final
  AMensagem := Format('%d rotas geradas, %d PAX alocados.',
    [Length(Rotas), Length(Alocs)]);
  AtualizarStatusOperacao(AIdOperacao, 'CONCLUIDO', AMensagem);
  Result := True;
end;

end.
