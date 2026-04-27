unit uRTBusinessRules;

interface

uses
  System.SysUtils, Data.DB, ADODB, System.Generics.Collections,
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
