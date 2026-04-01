unit uSimulacaoLogistica;

interface

uses
  System.SysUtils, System.Classes, System.Generics.Collections, Math;

// Tipos para parâmetros de simulação

type
  TSimulacaoParametros = record
    QtdSurfer: Integer;
    QtdSOV: Integer;
    QtdBAE: Integer;
    QtdAQUA: Integer;
    Backlog: Integer; // em movimentações
    DiasDisponiveis: Integer;
    CustoSurfer: Double;
    CustoSOV: Double;
    CustoBAE: Double;
    CustoAQUA: Double;
    ProdutividadeSurfer: Double;
    ProdutividadeSOV: Double;
    ProdutividadeBAE: Double;
    ProdutividadeAQUA: Double;
    HorasUteisSurfer: Double;
    HorasUteisSOV: Double;
    HorasUteisBAE: Double;
    HorasUteisAQUA: Double;
  end;

  TSimulacaoResultado = record
    PrazoEstimado: Integer;
    CustoTotal: Double;
    HorasUteisGeradas: Double;
    CustoPorHoraUtil: Double;
    PlataformasConcluidasPct: Double;
    Relatorio: string;
  end;

  TSimulacaoCenario = record
    Nome: string;
    Parametros: TSimulacaoParametros;
    Resultado: TSimulacaoResultado;
  end;

// Função principal de simulação
type
  TSimulacaoLogistica = class
  public
    class function Simular(const Param: TSimulacaoParametros): TSimulacaoResultado;
    class function GerarRelatorio(const Resultado: TSimulacaoResultado): string;
  end;

implementation

{ TSimulacaoLogistica }

class function TSimulacaoLogistica.Simular(const Param: TSimulacaoParametros): TSimulacaoResultado;
var
  TotalHorasUteis, TotalCusto: Double;
  Prazo, BacklogRestante: Integer;
  MovPorDia: Double;
  PlataformasConcluidas: Double;
begin
  // Cálculo simplificado: soma produtividade ponderada por tipo
  TotalHorasUteis :=
    Param.QtdSurfer * Param.HorasUteisSurfer * Param.ProdutividadeSurfer +
    Param.QtdSOV * Param.HorasUteisSOV * Param.ProdutividadeSOV +
    Param.QtdBAE * Param.HorasUteisBAE * Param.ProdutividadeBAE +
    Param.QtdAQUA * Param.HorasUteisAQUA * Param.ProdutividadeAQUA;

  TotalCusto :=
    Param.QtdSurfer * Param.CustoSurfer * Param.DiasDisponiveis +
    Param.QtdSOV * Param.CustoSOV * Param.DiasDisponiveis +
    Param.QtdBAE * Param.CustoBAE * Param.DiasDisponiveis +
    Param.QtdAQUA * Param.CustoAQUA * Param.DiasDisponiveis;

  // Movimentações/dia: assume 1 mov/hora útil
  MovPorDia := TotalHorasUteis;
  if MovPorDia <= 0 then
    Prazo := 0
  else
    Prazo := Ceil(Param.Backlog / MovPorDia);

  BacklogRestante := Max(0, Param.Backlog - Trunc(MovPorDia * Param.DiasDisponiveis));
  PlataformasConcluidas := 1.0;
  if Param.Backlog > 0 then
    PlataformasConcluidas := (Param.Backlog - BacklogRestante) / Param.Backlog;

  Result.PrazoEstimado := Prazo;
  Result.CustoTotal := TotalCusto;
  Result.HorasUteisGeradas := TotalHorasUteis * Param.DiasDisponiveis;
  if Result.HorasUteisGeradas > 0 then
    Result.CustoPorHoraUtil := TotalCusto / Result.HorasUteisGeradas
  else
    Result.CustoPorHoraUtil := 0;
  Result.PlataformasConcluidasPct := PlataformasConcluidas * 100;
  Result.Relatorio := GerarRelatorio(Result);
end;

class function TSimulacaoLogistica.GerarRelatorio(const Resultado: TSimulacaoResultado): string;
begin
  Result :=
    '--- Relatório de Simulação Logística ---' + sLineBreak +
    Format('Prazo estimado para concluir backlog: %d dias', [Resultado.PrazoEstimado]) + sLineBreak +
    Format('Custo total estimado: R$ %.2f', [Resultado.CustoTotal]) + sLineBreak +
    Format('Horas úteis geradas: %.1f', [Resultado.HorasUteisGeradas]) + sLineBreak +
    Format('Custo por hora útil: R$ %.2f', [Resultado.CustoPorHoraUtil]) + sLineBreak +
    Format('%% do backlog atendido: %.1f%%', [Resultado.PlataformasConcluidasPct]) + sLineBreak;
end;

end.
