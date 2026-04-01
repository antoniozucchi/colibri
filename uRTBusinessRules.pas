unit uRTBusinessRules;

interface

uses
  System.SysUtils, System.Classes, Data.DB, ADODB, System.Generics.Collections,
  uProgramacaoRTUtils;

type
  TRegraRecolhimentoRT = record
    ID: Integer;
    Ativa: Boolean;
    Prioridade: Integer;
    Descricao: string;
    Origem: string;
    Destino: string;
    DestinoNaoCriarRT: string; // '', 'SIM', 'NAO'
    RecolherParaTipo: string;  // 'ORIGEM', 'DESTINO', 'FIXO'
    RecolherParaValor: string;
    Observacao: string;
  end;

  TRTBusinessRules = class
  private
    FConnColibri: TADOConnection;
    FConnRT: TADOConnection;
    FConnConsulta: TADOConnection;
    
    function CampoAsString(DS: TDataSet; const NomeCampo: string): string;
    function NormalizarCodigoSAP(const ACodigo: string): string;
    function DeterminarTipoRTAutomatico(const ACodigoSAP: string): string;
    function DeterminarModalAutomatico(const AOrigem, ADestino: string): string;
    function DetermineClasse(const ATipoRT, AOrigem, ADestino: string; 
      AOrigemPlataformas: TStringList; const AClasseAtual: string; 
      ABooleanRecolhimento: Boolean): string;
      
  public
    constructor Create(AConnColibri, AConnRT, AConnConsulta: TADOConnection);
    
    // Core preparation methods
    procedure PrepararProgramacaoExecutanteParaRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean = False);
    procedure SincronizarProgramacaoExecutanteComRTExistente(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
    procedure AplicarRegrasAutomaticasRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
    procedure AplicarRegrasRecolhimentoBanco(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
    procedure RecalcularClassesRTNoPeriodo(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
    procedure AvaliarNecessidadeCriacaoRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
    
    // SAP Integration
    procedure ProcessarRetornoSAP(const LinhaLog: string);
    
    // Reconciliation
    procedure ConciliarRTsOrfasNoPeriodo(const ADataIni, ADataFim: TDateTime);
  end;

implementation

{ TRTBusinessRules }

constructor TRTBusinessRules.Create(AConnColibri, AConnRT, AConnConsulta: TADOConnection);
begin
  FConnColibri := AConnColibri;
  FConnRT := AConnRT;
  FConnConsulta := AConnConsulta;
end;

function TRTBusinessRules.CampoAsString(DS: TDataSet; const NomeCampo: string): string;
begin
  if DS.FindField(NomeCampo) <> nil then
    Result := Trim(DS.FieldByName(NomeCampo).AsString)
  else
    Result := '';
end;

function TRTBusinessRules.NormalizarCodigoSAP(const ACodigo: string): string;
var
  I: Integer;
  ApenasNumeros: Boolean;
begin
  Result := Trim(ACodigo);
  if Result = '' then Exit;

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
end;

function TRTBusinessRules.DeterminarTipoRTAutomatico(const ACodigoSAP: string): string;
var
  I: Integer;
  ApenasNumeros: Boolean;
begin
  if Trim(ACodigoSAP) = '' then
    Exit('R7');

  ApenasNumeros := True;
  for I := 1 to Length(ACodigoSAP) do
  begin
    if not CharInSet(ACodigoSAP[I], ['0'..'9']) then
    begin
      ApenasNumeros := False;
      Break;
    end;
  end;

  if ApenasNumeros then
    Result := 'R3'
  else
    Result := 'R7';
end;

function TRTBusinessRules.DeterminarModalAutomatico(const AOrigem, ADestino: string): string;
begin
  // Lógica simplificada baseada no contexto
  Result := 'AEREO'; // Default
  // Aqui entraria a lógica real de consulta à tblPlataforma
end;

function TRTBusinessRules.DetermineClasse(const ATipoRT, AOrigem, ADestino: string; 
  AOrigemPlataformas: TStringList; const AClasseAtual: string; 
  ABooleanRecolhimento: Boolean): string;
begin
  if SameText(ATipoRT, 'R7') then
    Result := 'COM'
  else if ABooleanRecolhimento and SameText(AOrigem, 'TMIB') then
    Result := 'EV'
  else if SameText(AClasseAtual, 'TR') then
    Result := 'TR'
  else
    Result := 'TT';
end;

procedure TRTBusinessRules.PrepararProgramacaoExecutanteParaRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  SincronizarProgramacaoExecutanteComRTExistente(ADataIni, ADataFim, AForcarTudo);
  AplicarRegrasAutomaticasRT(ADataIni, ADataFim, AForcarTudo);
  AplicarRegrasRecolhimentoBanco(ADataIni, ADataFim, AForcarTudo);
  RecalcularClassesRTNoPeriodo(ADataIni, ADataFim, AForcarTudo);
  AvaliarNecessidadeCriacaoRT(ADataIni, ADataFim, AForcarTudo);
end;

procedure TRTBusinessRules.SincronizarProgramacaoExecutanteComRTExistente(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  // Implementação da sincronização
end;

procedure TRTBusinessRules.AplicarRegrasAutomaticasRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  // Implementação das regras automáticas
end;

procedure TRTBusinessRules.AplicarRegrasRecolhimentoBanco(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  // Implementação das regras de recolhimento
end;

procedure TRTBusinessRules.RecalcularClassesRTNoPeriodo(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  // Implementação do recálculo de classes
end;

procedure TRTBusinessRules.AvaliarNecessidadeCriacaoRT(const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
begin
  // Implementação da avaliação de necessidade
end;

procedure TRTBusinessRules.ProcessarRetornoSAP(const LinhaLog: string);
begin
  // Implementação do processamento de retorno SAP
end;

procedure TRTBusinessRules.ConciliarRTsOrfasNoPeriodo(const ADataIni, ADataFim: TDateTime);
begin
  // Implementação da conciliação de RTs órfãs
end;

end.
