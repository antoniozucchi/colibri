unit uFuncaoExecutanteUtils;

interface

uses
  System.SysUtils, Data.Win.ADODB, System.Generics.Collections, Data.DB;

procedure GarantirTabelaDeParaFuncaoExecutante(const AConnection: TADOConnection);
procedure CarregarMapaDeParaFuncaoExecutante(const AConnection: TADOConnection);
procedure LimparMapaDeParaFuncaoExecutante;
procedure RegistrarDeParaFuncaoExecutante(const AFuncaoOrigem, AFuncaoPadrao: string);
function ChaveFuncaoExecutante(const AFuncao: string): string;
function NormalizarFuncaoExecutante(const AFuncao: string): string;

implementation

uses
  uAccessDBUtils;

var
  GMapaDeParaFuncaoExecutante: TDictionary<string, string>;

function RemoverAcentosMaiusculo(const ATexto: string): string;
begin
  Result := UpperCase(Trim(ATexto));
  Result := StringReplace(Result, '‰', 'E', [rfReplaceAll]);
  Result := StringReplace(Result, '‚', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, '‡', 'C', [rfReplaceAll]);
  Result := StringReplace(Result, '', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Á', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, 'À', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, 'Â', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ã', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ä', 'A', [rfReplaceAll]);
  Result := StringReplace(Result, 'É', 'E', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ê', 'E', [rfReplaceAll]);
  Result := StringReplace(Result, 'È', 'E', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ë', 'E', [rfReplaceAll]);
  Result := StringReplace(Result, 'Í', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ì', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Î', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ï', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ó', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ò', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ô', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Õ', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ö', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ú', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ù', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Û', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ü', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ç', 'C', [rfReplaceAll]);
end;

function NormalizarEspacos(const ATexto: string): string;
begin
  Result := Trim(ATexto);
  while Pos('  ', Result) > 0 do
    Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;

function RemoverQualificadoresSenioridade(const ATexto: string): string;
begin
  Result := ' ' + NormalizarEspacos(ATexto) + ' ';
  Result := StringReplace(Result, ' PLENO ', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' SENIOR ', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' MASTER ', ' ', [rfReplaceAll]);
  Result := NormalizarEspacos(Result);
end;

function LimparFuncaoBase(const AFuncao: string): string;
begin
  Result := RemoverAcentosMaiusculo(AFuncao);
  Result := StringReplace(Result, '(A)', '', [rfReplaceAll]);
  Result := StringReplace(Result, '(O)', '', [rfReplaceAll]);
  Result := StringReplace(Result, '/', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, '-', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, '.', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ',', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ';', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ':', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, '(', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ')', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' TEC ', ' TECNICO ', [rfReplaceAll]);
  Result := StringReplace(Result, ' TECNICA ', ' TECNICO ', [rfReplaceAll]);
  Result := StringReplace(Result, ' TECNICA DE ', ' TECNICO DE ', [rfReplaceAll]);
  Result := NormalizarEspacos(Result);

  if Pos('TEC ', Result) = 1 then
    Result := 'TECNICO ' + Copy(Result, 5, MaxInt);

  if Pos('TECNICA ', Result) = 1 then
    Result := 'TECNICO ' + Copy(Result, 9, MaxInt);

  if Pos('TECNICA DE ', Result) = 1 then
    Result := 'TECNICO DE ' + Copy(Result, 12, MaxInt);

  Result := RemoverQualificadoresSenioridade(Result);
  Result := NormalizarEspacos(Result);
end;

procedure LimparMapaDeParaFuncaoExecutante;
begin
  if GMapaDeParaFuncaoExecutante <> nil then
    GMapaDeParaFuncaoExecutante.Clear;
end;

procedure RegistrarDeParaFuncaoExecutante(const AFuncaoOrigem,
  AFuncaoPadrao: string);
var
  ChaveOrigem, FuncaoPadrao: string;
begin
  if GMapaDeParaFuncaoExecutante = nil then
    GMapaDeParaFuncaoExecutante := TDictionary<string, string>.Create;

  ChaveOrigem := LimparFuncaoBase(AFuncaoOrigem);
  FuncaoPadrao := LimparFuncaoBase(AFuncaoPadrao);

  if (ChaveOrigem = '') or (FuncaoPadrao = '') then
    Exit;

  GMapaDeParaFuncaoExecutante.AddOrSetValue(ChaveOrigem, FuncaoPadrao);
end;

function ChaveFuncaoExecutante(const AFuncao: string): string;
begin
  Result := LimparFuncaoBase(AFuncao);
end;

procedure GarantirTabelaDeParaFuncaoExecutante(const AConnection: TADOConnection);
var
  Q: TADOQuery;

  procedure InserirAliasSeAusente(const AFuncaoOrigem, AFuncaoPadrao: string);
  begin
    Q.Close;
    Q.SQL.Text :=
      'SELECT TOP 1 idDeParaFuncaoExecutante ' +
      'FROM tblDeParaFuncaoExecutante ' +
      'WHERE ChaveOrigem = :ChaveOrigem';
    Q.Parameters.ParamByName('ChaveOrigem').Value := ChaveFuncaoExecutante(AFuncaoOrigem);
    Q.Open;
    if not Q.IsEmpty then
      Exit;

    Q.Close;
    Q.SQL.Text :=
      'INSERT INTO tblDeParaFuncaoExecutante (' +
      '  FuncaoOrigem, FuncaoPadrao, ChaveOrigem, ChavePadrao, Ativa, Observacao ' +
      ') VALUES (' +
      '  :FuncaoOrigem, :FuncaoPadrao, :ChaveOrigem, :ChavePadrao, :Ativa, :Observacao' +
      ')';
    Q.Parameters.ParamByName('FuncaoOrigem').Value := Trim(AFuncaoOrigem);
    Q.Parameters.ParamByName('FuncaoPadrao').Value := Trim(AFuncaoPadrao);
    Q.Parameters.ParamByName('ChaveOrigem').Value := ChaveFuncaoExecutante(AFuncaoOrigem);
    Q.Parameters.ParamByName('ChavePadrao').Value := ChaveFuncaoExecutante(AFuncaoPadrao);
    Q.Parameters.ParamByName('Ativa').Value := True;
    Q.Parameters.ParamByName('Observacao').Value := 'Carga inicial de normalizacao';
    Q.ExecSQL;
  end;
begin
  if AConnection = nil then
    Exit;

  CriarTableDB(
    'tblDeParaFuncaoExecutante',
    'idDeParaFuncaoExecutante',
    '[FuncaoOrigem] VARCHAR(150), ' +
    '[FuncaoPadrao] VARCHAR(150), ' +
    '[ChaveOrigem] VARCHAR(150), ' +
    '[ChavePadrao] VARCHAR(150), ' +
    '[Ativa] YESNO, ' +
    '[Observacao] VARCHAR(255)',
    AConnection
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := AConnection;
    Q.ParamCheck := True;

    InserirAliasSeAusente('TÉCNICA DE SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TÉCNICO(A) DE SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TÉC DE SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TÉC SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TEC. SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TÉC EM SEGURANÇA DO TRABALHO', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TECNICO DE SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TECNICO DE SEGURANCA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TECNICO DE SEGURANÇA DO TRABALHO', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TECNICO SEGURANÇA', 'TÉCNICO DE SEGURANÇA');
    InserirAliasSeAusente('TÉCNICA DE SEGURANÇA', 'TÉCNICO DE SEGURANÇA');

    InserirAliasSeAusente('TÉCNICA DE MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TÉC MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TEC MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TECNICO DE MANUTENCAO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TÉCNICO DE MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TECNICO MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TÉCNICO MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TÉC EM MANUTENÇÃO', 'TÉCNICO DE MANUTENÇÃO');
    InserirAliasSeAusente('TECNICO EM MANUTENÇÃO N3', 'TÉCNICO EM MANUTENÇÃO N3');
    InserirAliasSeAusente('TÉCNICO EM MANUTENÇÃO N3', 'TÉCNICO EM MANUTENÇÃO N3');

    InserirAliasSeAusente('TÉC DE PLANEJAMENTO', 'TÉCNICO DE PLANEJAMENTO');
    InserirAliasSeAusente('TECNICO DE PLANEJAMENTO', 'TÉCNICO DE PLANEJAMENTO');
    InserirAliasSeAusente('TECNICO PLANEJAMENTO', 'TÉCNICO DE PLANEJAMENTO');
    InserirAliasSeAusente('TÉC PLANEJAMENTO', 'TÉCNICO DE PLANEJAMENTO');
    InserirAliasSeAusente('TA‰CNICO DE PLANEJAMENTO', 'TÉCNICO DE PLANEJAMENTO');

    InserirAliasSeAusente('TÉC TELECOMUNICAÇÕES', 'TÉCNICO DE TELECOMUNICAÇÕES');
    InserirAliasSeAusente('TECNICO TELECOMUNICAÇÕES', 'TÉCNICO DE TELECOMUNICAÇÕES');
    InserirAliasSeAusente('TÉC DE TELECOMUNICAÇÕES', 'TÉCNICO DE TELECOMUNICAÇÕES');

    InserirAliasSeAusente('TÉC ELÉTRICA', 'TÉCNICO ELÉTRICA');
    InserirAliasSeAusente('TECNICO ELÉTRICA', 'TÉCNICO ELÉTRICA');
    InserirAliasSeAusente('TÉC ELETRÔNICA', 'TÉCNICO ELETRÔNICA');
    InserirAliasSeAusente('TECNICO ELETRÔNICA', 'TÉCNICO ELETRÔNICA');
    InserirAliasSeAusente('TÉC EM ELETRÔNICA', 'TÉCNICO ELETRÔNICA');
    InserirAliasSeAusente('TÉC MECÂNICA', 'TÉCNICO MECÂNICA');
    InserirAliasSeAusente('TECNICO MECÂNICA', 'TÉCNICO MECÂNICA');
    InserirAliasSeAusente('TÉC MECÂNICA II', 'TÉCNICO MECÂNICA II');
    InserirAliasSeAusente('TECNICO EM MECANICA', 'TÉCNICO MECÂNICA');
    InserirAliasSeAusente('TECNICO EM MECANICA II', 'TÉCNICO MECÂNICA II');
    InserirAliasSeAusente('TA‰CNICO MECA‚NICA', 'TÉCNICO MECÂNICA');
    InserirAliasSeAusente('TÉC QUÍMICO', 'TÉCNICO QUÍMICO');
    InserirAliasSeAusente('TECNICO QUIMICO DE PETROLEO SENIOR', 'TÉCNICO QUÍMICO');
    InserirAliasSeAusente('TÉC REFRIGERAÇÃO', 'TÉCNICO DE REFRIGERAÇÃO');
    InserirAliasSeAusente('TECNICO REFRIGERAÇÃO', 'TÉCNICO DE REFRIGERAÇÃO');

    InserirAliasSeAusente('AUX DE ANDAIME', 'AUXILIAR DE ANDAIME');
    InserirAliasSeAusente('AUXILIAR DE ANDIME', 'AUXILIAR DE ANDAIME');
    InserirAliasSeAusente('AUX ADMINISTRATIVO', 'AUXILIAR ADMINISTRATIVO');
    InserirAliasSeAusente('AUX CALDEIRARIA', 'AUXILIAR DE CALDEIRARIA');
    InserirAliasSeAusente('AUX DE CALDEIRARIA', 'AUXILIAR DE CALDEIRARIA');
    InserirAliasSeAusente('AUX DE LOGISTICA', 'AUXILIAR DE LOGISTICA');
    InserirAliasSeAusente('AUX MANUTENÇÃO', 'AUXILIAR DE MANUTENÇÃO');
    InserirAliasSeAusente('AUX DE MANUTENÇÃO', 'AUXILIAR DE MANUTENÇÃO');
    InserirAliasSeAusente('AUX DE PLATAFORMA', 'AUXILIAR DE PLATAFORMA');
    InserirAliasSeAusente('AUX DE TOPÓGRAFO', 'AUXILIAR DE TOPOGRAFIA');
    InserirAliasSeAusente('AUX FAROLEIRO', 'AUXILIAR FAROLEIRO');
    InserirAliasSeAusente('AUX MOVIMENTAÇÃO DE CARGA', 'AUXILIAR DE MOVIMENTAÇÃO');
    InserirAliasSeAusente('AUXILIAR DE MOVIMENTAÇÃO', 'AUXILIAR DE MOVIMENTAÇÃO');
    InserirAliasSeAusente('AUX PLATAFORMISTA', 'AUXILIAR DE PLATAFORMA');
    InserirAliasSeAusente('AUXILIAR DE MOVIMENTAÇÃO', 'AUXILIAR DE MOVIMENTAÇÃO');
    InserirAliasSeAusente('AUXILIAR DE TOPOGRAFIA', 'AUXILIAR DE TOPOGRAFIA');
    InserirAliasSeAusente('AUXILIAR DE OPERAÇÃO', 'AUXILIAR DE OPERAÇÃO');
    InserirAliasSeAusente('AUXILIAR DE COZINHA', 'AUXILIAR DE COZINHA');
    InserirAliasSeAusente('AUXILIAR DE CALDEIRARIA', 'AUXILIAR DE CALDEIRARIA');
    InserirAliasSeAusente('AUXILIAR DE ANDAIME', 'AUXILIAR DE ANDAIME');
    InserirAliasSeAusente('AUXILIAR', 'AUXILIAR');

    InserirAliasSeAusente('EDUCADOR FISICO', 'EDUCADOR FÍSICO');
    InserirAliasSeAusente('EDUCADORA FISICA', 'EDUCADOR FÍSICO');

    InserirAliasSeAusente('OPERADOR DE SERVIÇOS', 'OPERADOR DE SERVICOS');
    InserirAliasSeAusente('OPERADOR SERVICOS', 'OPERADOR DE SERVICOS');
    InserirAliasSeAusente('OPERADOR DER SERVICOS', 'OPERADOR DE SERVICOS');
    InserirAliasSeAusente('OPERADOR DE SERVICOS', 'OPERADOR DE SERVICOS');
    InserirAliasSeAusente('OPERADOR DE SLICK LINE', 'OPERADOR DE SLICKLINE');
    InserirAliasSeAusente('OPERADOR SLICKLINE', 'OPERADOR DE SLICKLINE');
    InserirAliasSeAusente('OPERADOR SLK', 'OPERADOR DE SLICKLINE');
    InserirAliasSeAusente('OP DE EQUIPAMENTOS SLICKLINE', 'OPERADOR DE SLICKLINE');
    InserirAliasSeAusente('OPERACAO COM ARAME', 'OPERADOR DE SLICKLINE');
    InserirAliasSeAusente('SPECIALIST SLICKLINE', 'ESPECIALISTA SLICKLINE');
    InserirAliasSeAusente('SLICKLINE', 'ESPECIALISTA SLICKLINE');

    InserirAliasSeAusente('MECANICO', 'MECÂNICO');
    InserirAliasSeAusente('MECANICO DE AVIAÇÃO', 'MECÂNICO DE AVIAÇÃO');
    InserirAliasSeAusente('TECNICO MECANICO', 'TÉCNICO MECÂNICO');
    InserirAliasSeAusente('TECNICO MECÂNICO', 'TÉCNICO MECÂNICO');
    InserirAliasSeAusente('TA‰CNICO MECA‚NICO', 'TÉCNICO MECÂNICO');
    InserirAliasSeAusente('ENG DE EQUIPAMENTOS MASTER', 'ENGENHEIRO DE EQUIPAMENTOS MASTER');
    InserirAliasSeAusente('ENG DE SEGURANÇA', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENG MECÂNCO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENG MECÂNICO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENG PESCARIA', 'ENGENHEIRO DE PESCARIA');
    InserirAliasSeAusente('ENG SEGURANÇA DE PROCESSO', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGª CIVIL', 'ENGENHEIRO CIVIL');
    InserirAliasSeAusente('ENGª DE PROCESSOS', 'ENGENHEIRO DE PROCESSOS');
    InserirAliasSeAusente('ENGº CIVIL', 'ENGENHEIRO CIVIL');
    InserirAliasSeAusente('ENGª DE SEGURANÇA', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGª MECANICA', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENGENHEIRA', 'ENGENHEIRO');
    InserirAliasSeAusente('ENGENHEIRA DE ALIMENTOS', 'ENGENHEIRO DE ALIMENTOS');
    InserirAliasSeAusente('ENGENHEIRA DE PROCESSO', 'ENGENHEIRO DE PROCESSOS');
    InserirAliasSeAusente('ENGENHEIRA QUÍMICA', 'ENGENHEIRO QUÍMICO');
    InserirAliasSeAusente('ENGENHEIRO', 'ENGENHEIRO');
    InserirAliasSeAusente('ENGENHEIRO CIVIL', 'ENGENHEIRO CIVIL');
    InserirAliasSeAusente('ENGENHEIRO DE SEGURANCA', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGENHEIRO SEGURANÇA', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGENHEIRO DE EQUIPAMENTOS ELETRICA SENIOR', 'ENGENHEIRO DE EQUIPAMENTOS');
    InserirAliasSeAusente('ENGENHEIRO DE EQUIPAMENTOS MASTER', 'ENGENHEIRO DE EQUIPAMENTOS');
    InserirAliasSeAusente('ENGENHEIRO DE EQUIPAMENTOS MECANICA', 'ENGENHEIRO DE EQUIPAMENTOS');
    InserirAliasSeAusente('ENGENHEIRO DE EQUIPAMENTOS PLENO', 'ENGENHEIRO DE EQUIPAMENTOS');
    InserirAliasSeAusente('ENGENHEIRO DE EQUIPAMENTOS SENIOR', 'ENGENHEIRO DE EQUIPAMENTOS');
    InserirAliasSeAusente('ENGENHEIRO DE PETROLEO', 'ENGENHEIRO DE PETRÓLEO');
    InserirAliasSeAusente('ENGENHEIRO DE PETROLEO PLENO', 'ENGENHEIRO DE PETRÓLEO');
    InserirAliasSeAusente('ENGENHEIRO DE PROCESSAMENTO PLENO', 'ENGENHEIRO DE PROCESSOS');
    InserirAliasSeAusente('ENGENHEIRO ELETRICISTA', 'ENGENHEIRO ELÉTRICO');
    InserirAliasSeAusente('ENGENHEIRO MECANICO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENGENHEIRO MÊCANICO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENGº AMBIENTAL', 'ENGENHEIRO AMBIENTAL');
    InserirAliasSeAusente('ENGº DE PRODUÇÃO', 'ENGENHEIRO DE PRODUÇÃO');
    InserirAliasSeAusente('ENGENHEIRO MECANICO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENGº DE SEGURANÇA', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGº SEGURANÇA DO TRABALHO', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGº DE TELECOM', 'ENGENHEIRO DE TELECOM');
    InserirAliasSeAusente('ENGº ELETRICISTA', 'ENGENHEIRO ELÉTRICO');
    InserirAliasSeAusente('ENGº MECÂNICO', 'ENGENHEIRO MECÂNICO');
    InserirAliasSeAusente('ENGº PETRÓLEO', 'ENGENHEIRO DE PETRÓLEO');
    InserirAliasSeAusente('ENGª DE SEGURANÇA', 'ENGENHEIRO DE SEGURANÇA');

    InserirAliasSeAusente('INSPETORA DE EQUIPEMENTOS', 'INSPETOR DE EQUIPAMENTOS');
    InserirAliasSeAusente('INSPETRO DE EQUIPAMENTOS', 'INSPETOR DE EQUIPAMENTOS');
    InserirAliasSeAusente('ISPETOR', 'INSPETOR');
    InserirAliasSeAusente('INSPETORA NAVAL', 'INSPETOR NAVAL');
    InserirAliasSeAusente('NUTRICIONISTA OFFSHORE', 'NUTRICIONISTA');
    InserirAliasSeAusente('NUTRICIONISTA PLENO', 'NUTRICIONISTA');
    InserirAliasSeAusente('PROGRAMADOR DE LOGISTICA', 'PROGRAMADOR LOGISTICO');
    InserirAliasSeAusente('PROGRAMADORA LOGISTICA', 'PROGRAMADOR LOGISTICO');
    InserirAliasSeAusente('PSICOLOGA', 'PSICOLOGO');
    InserirAliasSeAusente('TAIFERO', 'TAIFEIRO');
    InserirAliasSeAusente('AUDITORA', 'AUDITOR');
    InserirAliasSeAusente('AUDITORA INTERNA', 'AUDITOR INTERNO');
    InserirAliasSeAusente('AUDITOR LIDER', 'AUDITOR LIDER');
    InserirAliasSeAusente('COMISSARIA', 'COMISSARIO');
    InserirAliasSeAusente('ENFERMEIRA', 'ENFERMEIRO');
    InserirAliasSeAusente('ENFERMEIRA DO TRABALHO', 'ENFERMEIRO DO TRABALHO');
    InserirAliasSeAusente('COORDENADORA ADM', 'COORDENADOR ADM');
    InserirAliasSeAusente('COORDENADORA DE CONTRATO', 'COORDENADOR DE CONTRATO');
    InserirAliasSeAusente('COORDENADORA DE MARKETING', 'COORDENADOR DE MARKETING');
    InserirAliasSeAusente('COORDENADORA DE QSMS', 'COORDENADOR DE QSMS');
    InserirAliasSeAusente('COORDENADORA DE QUALIDADE', 'COORDENADOR DE QUALIDADE');
    InserirAliasSeAusente('SUPERVISORA', 'SUPERVISOR');
    InserirAliasSeAusente('SUPERVISORA PIOP', 'SUPERVISOR PIOP');
    InserirAliasSeAusente('SUPERVISORA SMS', 'SUPERVISOR SMS');
    InserirAliasSeAusente('SUPERVISOR DE OPERACOES', 'SUPERVISOR DE OPERACAO');
    InserirAliasSeAusente('ASSISTENTE DE OPERACOES', 'ASSISTENTE DE OPERACAO');
    InserirAliasSeAusente('ASSISTENTE OPERACOES', 'ASSISTENTE DE OPERACAO');
    InserirAliasSeAusente('OPERADOR DE PROCESSO', 'OPERADOR DE PROCESSOS');
    InserirAliasSeAusente('FISCAL DE POCOS', 'FISCAL DE POCO');
    InserirAliasSeAusente('FISCAL MANUTENCAO MECANICA', 'FISCAL MECANICA');
    InserirAliasSeAusente('PROF DE NIVEL SUPERIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF NIVEL TECNICO', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROF PETROBRAS', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL SUPERIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL SUPERIOR JUNIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL SUPERIOR MASTER', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL SUPERIOR PLENO', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL SUPERIOR SENIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS NIVEL SUPERIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL TECNICO', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL TECNICO MASTER', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL TECNICO PLENO', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROF PETROBRAS DE NIVEL TECNICO SENIOR', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFESSOR PETROBRAS NS', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFESSOR PETROBRAS NT', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFESSORA PETROBRAS NT', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFIS BR NIVEL TECNICO JR', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL SUPERIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL SUPERIOR JUNIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL SUPERIOR MASTER', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL SUPERIOR PLENO', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL SUPERIOR SENIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS NIVEL SUPERIOR', 'PROF PETROBRAS DE NIVEL SUPERIOR');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL TECNICO', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL TECNICO JUNIOR', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL TECNICO MASTER', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL TECNICO PLENO', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('PROFISSIONAL PETROBRAS DE NIVEL TECNICO SENIOR', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('TECNICO DE EQUIPAMENTO PETROBRAS', 'PROF PETROBRAS DE NIVEL TECNICO');
    InserirAliasSeAusente('CALDEREIRO', 'CALDEIREIRO');
    InserirAliasSeAusente('DIRETOR PRESITENTE', 'DIRETOR PRESIDENTE');
    InserirAliasSeAusente('PINTOR ESCALDOR N1', 'PINTOR ESCALADOR N1');
    InserirAliasSeAusente('EDUCADOR FASICO', 'EDUCADOR FISICO');
    InserirAliasSeAusente('ENGENHEIRO DE SEGURANA‡A', 'ENGENHEIRO DE SEGURANÇA');
    InserirAliasSeAusente('ENGENHEIRO MECA‚NICO', 'ENGENHEIRO MECANICO');
    InserirAliasSeAusente('MECA‚NICO', 'MECANICO');
    InserirAliasSeAusente('TA‰CNICO MECA‚NICO', 'TECNICO MECANICO');
  finally
    Q.Free;
  end;
end;

procedure CarregarMapaDeParaFuncaoExecutante(const AConnection: TADOConnection);
var
  Q: TADOQuery;
begin
  LimparMapaDeParaFuncaoExecutante;

  if AConnection = nil then
    Exit;

  GarantirTabelaDeParaFuncaoExecutante(AConnection);

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := AConnection;
    Q.ParamCheck := False;
    Q.SQL.Text :=
      'SELECT FuncaoOrigem, FuncaoPadrao ' +
      'FROM tblDeParaFuncaoExecutante ' +
      'WHERE Ativa = True OR Ativa IS NULL';
    Q.Open;

    while not Q.Eof do
    begin
      RegistrarDeParaFuncaoExecutante(
        Q.FieldByName('FuncaoOrigem').AsString,
        Q.FieldByName('FuncaoPadrao').AsString
      );
      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

function NormalizarFuncaoExecutante(const AFuncao: string): string;
var
  Chave: string;
  FuncaoPadrao: string;
begin
  Chave := LimparFuncaoBase(AFuncao);

  if Chave = '' then
    Exit('');

  if (GMapaDeParaFuncaoExecutante <> nil) and
     GMapaDeParaFuncaoExecutante.TryGetValue(Chave, FuncaoPadrao) then
    Exit(FuncaoPadrao);

  if Pos('ENGENHEIRA ', Chave) = 1 then
    Chave := 'ENGENHEIRO ' + Copy(Chave, 11, MaxInt)
  else if Chave = 'ENGENHEIRA' then
    Chave := 'ENGENHEIRO';

  if Pos('ENGA ', Chave) = 1 then
    Chave := 'ENGENHEIRO ' + Copy(Chave, 6, MaxInt)
  else if Chave = 'ENGA' then
    Chave := 'ENGENHEIRO';

  if Pos('ENGO ', Chave) = 1 then
    Chave := 'ENGENHEIRO ' + Copy(Chave, 6, MaxInt)
  else if Chave = 'ENGO' then
    Chave := 'ENGENHEIRO';

  if Pos('ENG ', Chave) = 1 then
    Chave := 'ENGENHEIRO ' + Copy(Chave, 5, MaxInt)
  else if Chave = 'ENG' then
    Chave := 'ENGENHEIRO';

  if Chave = 'ENGENHEIRO SEGURANCA' then
    Exit('ENGENHEIRO DE SEGURANCA');

  if Chave = 'ENGENHEIRO SEGURANCA DO TRABALHO' then
    Exit('ENGENHEIRO DE SEGURANCA');

  if Chave = 'ENGENHEIRO DE SEGURANCA DE PROCESSO' then
    Exit('ENGENHEIRO DE SEGURANCA');

  if Chave = 'ENGENHEIRO ELETRICISTA' then
    Exit('ENGENHEIRO ELETRICO');

  if Chave = 'ENGENHEIRO MECANCO' then
    Exit('ENGENHEIRO MECANICO');

  if Chave = 'ENGENHEIRO MENCANICO' then
    Exit('ENGENHEIRO MECANICO');

  if (Chave = 'VAZIA') or (Chave = 'NULO') then
    Exit('');

  if Pos('SUPERVISORA ', Chave) = 1 then
    Chave := 'SUPERVISOR ' + Copy(Chave, 12, MaxInt)
  else if Chave = 'SUPERVISORA' then
    Chave := 'SUPERVISOR';

  if Pos('COORDENADORA ', Chave) = 1 then
    Chave := 'COORDENADOR ' + Copy(Chave, 13, MaxInt)
  else if Chave = 'COORDENADORA' then
    Chave := 'COORDENADOR';

  if Pos('AUDITORA ', Chave) = 1 then
    Chave := 'AUDITOR ' + Copy(Chave, 10, MaxInt)
  else if Chave = 'AUDITORA' then
    Chave := 'AUDITOR';

  if Chave = 'COMISSARIA' then
    Chave := 'COMISSARIO';

  if Pos('TECNICA ', Chave) = 1 then
    Chave := 'TECNICO ' + Copy(Chave, 9, MaxInt)
  else if Chave = 'TECNICA' then
    Chave := 'TECNICO';

  if Pos('TECNICO EM ', Chave) = 1 then
    Chave := 'TECNICO DE ' + Copy(Chave, 12, MaxInt);

  if Chave = 'TECNICO AUTOMACAO' then
    Exit('TECNICO DE AUTOMACAO');

  if Chave = 'TECNICO DE SEGURANCA DO TRABALHO' then
    Exit('TECNICO DE SEGURANCA');

  if Chave = 'TECNICO EM SEGURANCA DO TRABALHO' then
    Exit('TECNICO DE SEGURANCA');

  if Chave = 'TECNICO SEGURANCA DAY TRIP' then
    Exit('TECNICO DE SEGURANCA');

  if Chave = 'TECNICO DE MANUTENCAO PLENO' then
    Exit('TECNICO DE MANUTENCAO');

  if Chave = 'TECNICO DE OPERACAO PLENO' then
    Exit('TECNICO DE OPERACAO');

  if Chave = 'TECNICO DE OPERACAO DE CAMPO' then
    Exit('TECNICO DE OPERACAO');

  if Chave = 'TECNICO OPERACOES' then
    Exit('TECNICO DE OPERACAO');

  if Chave = 'TECNICO DE INSPECAO DE EQUIPAMENTOS' then
    Exit('TECNICO DE INSPECAO');

  if Chave = 'TECNICO DE INSPECAO DE EQUIPAMENTOS E INSTALACOES' then
    Exit('TECNICO DE INSPECAO');

  if Chave = 'TECNICO ENFERMAGEM' then
    Exit('TECNICO DE ENFERMAGEM');

  if Chave = 'TECNICO DE ENFERMAGEM DO TRABALHO PLENO' then
    Exit('TECNICO DE ENFERMAGEM');

  if Chave = 'TECNICO DE LOGISTICA DE TRANSPORTE' then
    Exit('TECNICO DE LOGISTICA');

  if Chave = 'TECNICO MANUTENCAO ELETRICA' then
    Exit('TECNICO ELETRICA');

  if Chave = 'TECNICO TELECOMUNICACOES' then
    Exit('TECNICO DE TELECOMUNICACOES');

  if Chave = 'ASSISTENTE DE OPERACOES' then
    Exit('ASSISTENTE DE OPERACAO');

  if Chave = 'ASSISTENTE OPERACOES' then
    Exit('ASSISTENTE DE OPERACAO');

  if Chave = 'OPERADOR DE PROCESSO' then
    Exit('OPERADOR DE PROCESSOS');

  if Chave = 'SUPERVISOR DE OPERACOES' then
    Exit('SUPERVISOR DE OPERACAO');

  if Chave = 'FISCAL DE POCOS' then
    Exit('FISCAL DE POCO');

  if (Chave = 'TECNICO DE SEGURANCA') or
     (Chave = 'TECNICO SEGURANCA') or
     (Chave = 'TECNICA DE SEGURANCA') or
     (Chave = 'TECNICO A DE SEGURANCA') then
    Exit('TECNICO DE SEGURANCA');

  if (Chave = 'TECNICO DE MANUTENCAO') or
     (Chave = 'TECNICO MANUTENCAO') or
     (Chave = 'TECNICA DE MANUTENCAO') then
    Exit('TECNICO DE MANUTENCAO');

  if (Chave = 'AUX DE MANUTENCAO') or
     (Chave = 'AUX MANUTENCAO') then
    Exit('AUXILIAR DE MANUTENCAO');

  if Chave = 'AUX DE TOPOGRAFO' then
    Exit('AUXILIAR DE TOPOGRAFIA');

  if (Chave = 'COORDENADOR MANUTENCAO') or
     (Chave = 'COORDENADOR DE MANUTENCAO') then
    Exit('COORDENADOR DE MANUTENCAO');

  if (Chave = 'OPERADOR PROCESSOS') or
     (Chave = 'OPERADOR DE PROCESSOS') then
    Exit('OPERADOR DE PROCESSOS');

  Result := Chave;
end;

initialization
  GMapaDeParaFuncaoExecutante := TDictionary<string, string>.Create;

finalization
  GMapaDeParaFuncaoExecutante.Free;

end.
