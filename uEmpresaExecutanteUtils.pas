unit uEmpresaExecutanteUtils;

interface

uses
  System.SysUtils, Data.Win.ADODB, System.Generics.Collections, Data.DB;

procedure GarantirTabelaDeParaEmpresaExecutante(const AConnection: TADOConnection);
procedure CarregarMapaDeParaEmpresaExecutante(const AConnection: TADOConnection);
procedure LimparMapaDeParaEmpresaExecutante;
procedure RegistrarDeParaEmpresaExecutante(const AEmpresaOrigem, AEmpresaPadrao: string);
function ChaveEmpresaExecutante(const AEmpresa: string): string;
function NormalizarEmpresaExecutante(const AEmpresa: string): string;

implementation

uses
  uAccessDBUtils;

var
  GMapaDeParaEmpresaExecutante: TDictionary<string, string>;

function NormalizarEspacosEmpresa(const ATexto: string): string;
begin
  Result := Trim(ATexto);
  while Pos('  ', Result) > 0 do
    Result := StringReplace(Result, '  ', ' ', [rfReplaceAll]);
end;

function LimparEmpresaBase(const AEmpresa: string): string;
begin
  Result := UpperCase(Trim(AEmpresa));
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
  Result := StringReplace(Result, 'Í', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ì', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Î', 'I', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ó', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ò', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ô', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Õ', 'O', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ú', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ù', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Û', 'U', [rfReplaceAll]);
  Result := StringReplace(Result, 'Ç', 'C', [rfReplaceAll]);
  Result := StringReplace(Result, '.', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ',', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ';', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ':', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' LTDA', '', [rfReplaceAll]);
  Result := StringReplace(Result, ' LTDA ', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' EIRELI', '', [rfReplaceAll]);
  Result := StringReplace(Result, ' EIRELI ', ' ', [rfReplaceAll]);
  Result := StringReplace(Result, ' S A', ' SA', [rfReplaceAll]);
  Result := StringReplace(Result, ' SERVICOS TECNICOS', ' SERVICOS TECNICOS', [rfReplaceAll]);
  Result := NormalizarEspacosEmpresa(Result);
end;

function ChaveEmpresaExecutante(const AEmpresa: string): string;
begin
  Result := LimparEmpresaBase(AEmpresa);
end;

procedure LimparMapaDeParaEmpresaExecutante;
begin
  if GMapaDeParaEmpresaExecutante <> nil then
    GMapaDeParaEmpresaExecutante.Clear;
end;

procedure RegistrarDeParaEmpresaExecutante(const AEmpresaOrigem,
  AEmpresaPadrao: string);
var
  ChaveOrigem, EmpresaPadrao: string;
begin
  if GMapaDeParaEmpresaExecutante = nil then
    GMapaDeParaEmpresaExecutante := TDictionary<string, string>.Create;

  ChaveOrigem := ChaveEmpresaExecutante(AEmpresaOrigem);
  EmpresaPadrao := ChaveEmpresaExecutante(AEmpresaPadrao);

  if (ChaveOrigem = '') or (EmpresaPadrao = '') then
    Exit;

  GMapaDeParaEmpresaExecutante.AddOrSetValue(ChaveOrigem, EmpresaPadrao);
end;

procedure GarantirTabelaDeParaEmpresaExecutante(const AConnection: TADOConnection);
var
  Q: TADOQuery;

  procedure InserirAliasSeAusente(const AEmpresaOrigem, AEmpresaPadrao: string);
  begin
    Q.Close;
    Q.SQL.Text :=
      'SELECT TOP 1 idDeParaEmpresaExecutante ' +
      'FROM tblDeParaEmpresaExecutante ' +
      'WHERE ChaveOrigem = :ChaveOrigem';
    Q.Parameters.ParamByName('ChaveOrigem').Value := ChaveEmpresaExecutante(AEmpresaOrigem);
    Q.Open;
    if not Q.IsEmpty then
      Exit;

    Q.Close;
    Q.SQL.Text :=
      'INSERT INTO tblDeParaEmpresaExecutante (' +
      '  EmpresaOrigem, EmpresaPadrao, ChaveOrigem, ChavePadrao, Ativa, Observacao ' +
      ') VALUES (' +
      '  :EmpresaOrigem, :EmpresaPadrao, :ChaveOrigem, :ChavePadrao, :Ativa, :Observacao' +
      ')';
    Q.Parameters.ParamByName('EmpresaOrigem').Value := Trim(AEmpresaOrigem);
    Q.Parameters.ParamByName('EmpresaPadrao').Value := Trim(AEmpresaPadrao);
    Q.Parameters.ParamByName('ChaveOrigem').Value := ChaveEmpresaExecutante(AEmpresaOrigem);
    Q.Parameters.ParamByName('ChavePadrao').Value := ChaveEmpresaExecutante(AEmpresaPadrao);
    Q.Parameters.ParamByName('Ativa').Value := True;
    Q.Parameters.ParamByName('Observacao').Value := 'Carga inicial de normalizacao';
    Q.ExecSQL;
  end;
begin
  if AConnection = nil then
    Exit;

  CriarTableDB(
    'tblDeParaEmpresaExecutante',
    'idDeParaEmpresaExecutante',
    '[EmpresaOrigem] VARCHAR(150), ' +
    '[EmpresaPadrao] VARCHAR(150), ' +
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

    InserirAliasSeAusente('ATLAS INSPEÇÕES', 'ATLAS INSPECOES');
    InserirAliasSeAusente('BV', 'BUREAU VERITAS');
    InserirAliasSeAusente('BVQI', 'BUREAU VERITAS');
    InserirAliasSeAusente('CMM OFFSHORE BRASIL', 'CMM');
    InserirAliasSeAusente('ELASA ELO FORNECIMENTO DE ALIMENTAÇÃO DE MACAE LTDA', 'ELASA ELO FORNECIMENTO DE ALIMENTACAO DE MACAE');
    InserirAliasSeAusente('FCC SOLUÇÕES', 'FCC SOLUCOES');
    InserirAliasSeAusente('INERCO CONSULTORIA BRASIL', 'INERCO');
    InserirAliasSeAusente('LIDER AVIAÇÃO', 'LIDER AVIACAO');
    InserirAliasSeAusente('MODELLE', 'MODELE');
    InserirAliasSeAusente('MSL SERVIÇOS', 'MSL SERVICOS');
    InserirAliasSeAusente('NORDESTE EMERGÊNCIAS', 'NORDESTE EMERGENCIAS');
    InserirAliasSeAusente('NUCLIMA REFRIGERAÇÃO', 'NUCLIMA REFRIGERACAO');
    InserirAliasSeAusente('OCEÂNICA', 'OCEANICA');
    InserirAliasSeAusente('PLANLINK SOLUÇÕES', 'PLANLINK');
    InserirAliasSeAusente('RINA BRASIL SERVIÇOS TÉCNICOS', 'RINA BRASIL');
    InserirAliasSeAusente('SERBRAS SERVIÇOS', 'SERBRAS');
    InserirAliasSeAusente('SG MANUTENÇÃO', 'SG MANUTENCAO');
    InserirAliasSeAusente('T&D SERVIÇOS', 'T&D SERVICOS');
    InserirAliasSeAusente('TELEMÁTICA SISTEMAS', 'TELEMATICA SISTEMAS');
    InserirAliasSeAusente('W3HAUS COMUNICAÇÃO', 'W3HAUS COMUNICACAO');
    InserirAliasSeAusente('WN', 'NORWIND OFFSHORE');
    InserirAliasSeAusente('WWC', 'WILD WELL CONTROL');
  finally
    Q.Free;
  end;
end;

procedure CarregarMapaDeParaEmpresaExecutante(const AConnection: TADOConnection);
var
  Q: TADOQuery;
begin
  LimparMapaDeParaEmpresaExecutante;

  if AConnection = nil then
    Exit;

  GarantirTabelaDeParaEmpresaExecutante(AConnection);

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := AConnection;
    Q.ParamCheck := False;
    Q.SQL.Text :=
      'SELECT EmpresaOrigem, EmpresaPadrao ' +
      'FROM tblDeParaEmpresaExecutante ' +
      'WHERE Ativa = True OR Ativa IS NULL';
    Q.Open;

    while not Q.Eof do
    begin
      RegistrarDeParaEmpresaExecutante(
        Q.FieldByName('EmpresaOrigem').AsString,
        Q.FieldByName('EmpresaPadrao').AsString
      );
      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

function NormalizarEmpresaExecutante(const AEmpresa: string): string;
var
  Chave, EmpresaPadrao: string;
begin
  Chave := ChaveEmpresaExecutante(AEmpresa);
  if Chave = '' then
    Exit('');

  if (Chave = 'VAZIA') or (Chave = 'NULO') then
    Exit('');

  if (GMapaDeParaEmpresaExecutante <> nil) and
     GMapaDeParaEmpresaExecutante.TryGetValue(Chave, EmpresaPadrao) then
    Exit(EmpresaPadrao);

  if Chave = 'MODELLE' then
    Exit('MODELE');

  if Chave = 'NORDESTE EMERGENCIAS' then
    Exit('NORDESTE EMERGENCIAS');

  Result := Chave;
end;

initialization
  GMapaDeParaEmpresaExecutante := TDictionary<string, string>.Create;

finalization
  GMapaDeParaEmpresaExecutante.Free;

end.
