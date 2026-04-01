unit uSimulacaoLogistica;

interface

uses
  System.SysUtils, System.Classes, Math;

type
  TSimulacaoParametros = record
    QtdSurfer: Integer;
    QtdSOV: Integer;
    QtdBAE: Integer;
    QtdAQUA: Integer;
    Backlog: Integer; // em movimentacoes
    DiasDisponiveis: Integer;
    CustoSurfer: Double; // custo diario por unidade
    CustoSOV: Double;
    CustoBAE: Double;
    CustoAQUA: Double;
    ProdutividadeSurfer: Double; // movimentacoes por hora util
    ProdutividadeSOV: Double;
    ProdutividadeBAE: Double;
    ProdutividadeAQUA: Double;
    HorasUteisSurfer: Double; // horas planejadas por dia
    HorasUteisSOV: Double;
    HorasUteisBAE: Double;
    HorasUteisAQUA: Double;
    DisponibilidadeSurfer: Double; // percentual efetivo de operacao
    DisponibilidadeSOV: Double;
    DisponibilidadeBAE: Double;
    DisponibilidadeAQUA: Double;
    MobilizacaoSurfer: Double; // custo fixo por unidade no periodo
    MobilizacaoSOV: Double;
    MobilizacaoBAE: Double;
    MobilizacaoAQUA: Double;
  end;

  TSimulacaoResultado = record
    PrazoEstimado: Integer;
    MovimentacoesPorDia: Double;
    HorasPlanejadasPorDia: Double;
    HorasEfetivasPorDia: Double;
    CustoOperacionalTotal: Double;
    CustoMobilizacaoTotal: Double;
    CustoTotal: Double;
    HorasUteisGeradas: Double;
    CustoPorHoraUtil: Double;
    CustoPorMovimentacao: Double;
    IndiceResiliencia: Double;
    BacklogAtendido: Integer;
    BacklogRestante: Integer;
    PercentualBacklogAtendido: Double;
    Relatorio: string;
  end;

  TSimulacaoCenario = record
    Nome: string;
    Parametros: TSimulacaoParametros;
    Resultado: TSimulacaoResultado;
  end;

  TSimulacaoComparativo = record
    PrazoGanhoDias: Integer;
    TornaViavel: Boolean;
    MovimentacoesAdicionaisPorDia: Double;
    HorasUteisAdicionais: Double;
    BacklogAdicionalAtendido: Integer;
    CustoAdicional: Double;
    GanhoCapacidadePercentual: Double;
    ValorOperacionalEquivalente: Double;
    ROIOperacional: Double;
    PaybackDias: Double;
    CustoPorDiaGanho: Double;
  end;

  TSimulacaoLogistica = class
  public
    class function Simular(const Param: TSimulacaoParametros): TSimulacaoResultado;
    class function CompararCenarios(const Base,
      Cenario: TSimulacaoCenario): TSimulacaoComparativo;
    class function PrazoComoTexto(const Resultado: TSimulacaoResultado): string;
    class function PaybackComoTexto(
      const Comparativo: TSimulacaoComparativo): string;
    class function ROIComoTexto(
      const Comparativo: TSimulacaoComparativo): string;
    class function GerarRelatorio(const Resultado: TSimulacaoResultado): string;
  end;

implementation

class function TSimulacaoLogistica.CompararCenarios(const Base,
  Cenario: TSimulacaoCenario): TSimulacaoComparativo;
var
  DiasAnalise: Integer;
  ValorOperacionalDiario: Double;
begin
  DiasAnalise := Max(1, Base.Parametros.DiasDisponiveis);

  Result.PrazoGanhoDias := 0;
  Result.TornaViavel := (Base.Resultado.PrazoEstimado < 0) and
    (Cenario.Resultado.PrazoEstimado >= 0);

  if (Base.Resultado.PrazoEstimado >= 0) and
     (Cenario.Resultado.PrazoEstimado >= 0) then
    Result.PrazoGanhoDias := Base.Resultado.PrazoEstimado -
      Cenario.Resultado.PrazoEstimado;

  Result.MovimentacoesAdicionaisPorDia := Cenario.Resultado.MovimentacoesPorDia -
    Base.Resultado.MovimentacoesPorDia;
  Result.HorasUteisAdicionais := Cenario.Resultado.HorasUteisGeradas -
    Base.Resultado.HorasUteisGeradas;
  Result.BacklogAdicionalAtendido := Cenario.Resultado.BacklogAtendido -
    Base.Resultado.BacklogAtendido;
  Result.CustoAdicional := Cenario.Resultado.CustoTotal -
    Base.Resultado.CustoTotal;

  if Base.Resultado.MovimentacoesPorDia > 0 then
    Result.GanhoCapacidadePercentual :=
      (Result.MovimentacoesAdicionaisPorDia /
       Base.Resultado.MovimentacoesPorDia) * 100
  else
    Result.GanhoCapacidadePercentual := 0;

  Result.ValorOperacionalEquivalente :=
    Max(0.0, Result.MovimentacoesAdicionaisPorDia) *
    Base.Resultado.CustoPorMovimentacao * DiasAnalise;

  if Result.CustoAdicional > 0 then
    Result.ROIOperacional :=
      ((Result.ValorOperacionalEquivalente - Result.CustoAdicional) /
       Result.CustoAdicional) * 100
  else
    Result.ROIOperacional := 0;

  ValorOperacionalDiario :=
    Max(0.0, Result.MovimentacoesAdicionaisPorDia) *
    Base.Resultado.CustoPorMovimentacao;

  if (Result.CustoAdicional > 0) and (ValorOperacionalDiario > 0) then
    Result.PaybackDias := Result.CustoAdicional / ValorOperacionalDiario
  else
    Result.PaybackDias := -1;

  if (Result.CustoAdicional > 0) and (Result.PrazoGanhoDias > 0) then
    Result.CustoPorDiaGanho := Result.CustoAdicional /
      Result.PrazoGanhoDias
  else
    Result.CustoPorDiaGanho := 0;
end;

class function TSimulacaoLogistica.GerarRelatorio(
  const Resultado: TSimulacaoResultado): string;
begin
  Result :=
    '--- Relatorio de Simulacao Logistica ---' + sLineBreak +
    PrazoComoTexto(Resultado) + sLineBreak +
    Format('Movimentacoes estimadas por dia: %.1f',
      [Resultado.MovimentacoesPorDia]) + sLineBreak +
    Format('Horas planejadas por dia: %.1f',
      [Resultado.HorasPlanejadasPorDia]) + sLineBreak +
    Format('Horas efetivas por dia: %.1f',
      [Resultado.HorasEfetivasPorDia]) + sLineBreak +
    Format('Resiliencia estimada: %.1f%%',
      [Resultado.IndiceResiliencia]) + sLineBreak +
    Format('Custo operacional total: R$ %.2f',
      [Resultado.CustoOperacionalTotal]) + sLineBreak +
    Format('Custo de mobilizacao: R$ %.2f',
      [Resultado.CustoMobilizacaoTotal]) + sLineBreak +
    Format('Custo total estimado: R$ %.2f',
      [Resultado.CustoTotal]) + sLineBreak +
    Format('Horas uteis geradas no periodo: %.1f',
      [Resultado.HorasUteisGeradas]) + sLineBreak +
    Format('Custo por hora util: R$ %.2f',
      [Resultado.CustoPorHoraUtil]) + sLineBreak +
    Format('Custo por movimentacao atendida: R$ %.2f',
      [Resultado.CustoPorMovimentacao]) + sLineBreak +
    Format('Backlog atendido no periodo: %d',
      [Resultado.BacklogAtendido]) + sLineBreak +
    Format('Backlog restante: %d',
      [Resultado.BacklogRestante]) + sLineBreak +
    Format('%% do backlog atendido: %.1f%%',
      [Resultado.PercentualBacklogAtendido]) + sLineBreak;
end;

class function TSimulacaoLogistica.PaybackComoTexto(
  const Comparativo: TSimulacaoComparativo): string;
begin
  if Comparativo.CustoAdicional <= 0 then
  begin
    if Comparativo.ValorOperacionalEquivalente > 0 then
      Result := 'ganho sem custo adicional relevante'
    else
      Result := 'nao aplicavel';
    Exit;
  end;

  if Comparativo.PaybackDias < 0 then
    Result := 'nao recupera no modelo atual'
  else
    Result := Format('%.1f dias', [Comparativo.PaybackDias]);
end;

class function TSimulacaoLogistica.PrazoComoTexto(
  const Resultado: TSimulacaoResultado): string;
begin
  if Resultado.PrazoEstimado < 0 then
    Result := 'Prazo estimado para concluir backlog: inviavel com a capacidade atual'
  else
    Result := Format('Prazo estimado para concluir backlog: %d dias',
      [Resultado.PrazoEstimado]);
end;

class function TSimulacaoLogistica.ROIComoTexto(
  const Comparativo: TSimulacaoComparativo): string;
begin
  if Comparativo.CustoAdicional <= 0 then
  begin
    if Comparativo.ValorOperacionalEquivalente > 0 then
      Result := 'ganho sem custo adicional relevante'
    else
      Result := 'neutro'
  end
  else
    Result := Format('%.1f%%', [Comparativo.ROIOperacional]);
end;

class function TSimulacaoLogistica.Simular(
  const Param: TSimulacaoParametros): TSimulacaoResultado;
var
  HorasPlanejadasDia, HorasEfetivasDia: Double;
  MovPorDia, CustoOperacionalTotal, CustoMobilizacaoTotal: Double;
  Prazo, BacklogRestante, BacklogAtendido: Integer;
  PercentualBacklogAtendido: Double;
  DiasAnalise: Integer;
  DispSurfer, DispSOV, DispBAE, DispAQUA: Double;
begin
  DiasAnalise := Max(1, Param.DiasDisponiveis);

  DispSurfer := EnsureRange(Param.DisponibilidadeSurfer, 0.0, 100.0) / 100.0;
  DispSOV := EnsureRange(Param.DisponibilidadeSOV, 0.0, 100.0) / 100.0;
  DispBAE := EnsureRange(Param.DisponibilidadeBAE, 0.0, 100.0) / 100.0;
  DispAQUA := EnsureRange(Param.DisponibilidadeAQUA, 0.0, 100.0) / 100.0;

  HorasPlanejadasDia :=
    Param.QtdSurfer * Param.HorasUteisSurfer +
    Param.QtdSOV * Param.HorasUteisSOV +
    Param.QtdBAE * Param.HorasUteisBAE +
    Param.QtdAQUA * Param.HorasUteisAQUA;

  HorasEfetivasDia :=
    Param.QtdSurfer * Param.HorasUteisSurfer * DispSurfer +
    Param.QtdSOV * Param.HorasUteisSOV * DispSOV +
    Param.QtdBAE * Param.HorasUteisBAE * DispBAE +
    Param.QtdAQUA * Param.HorasUteisAQUA * DispAQUA;

  MovPorDia :=
    Param.QtdSurfer * Param.HorasUteisSurfer * DispSurfer *
      Param.ProdutividadeSurfer +
    Param.QtdSOV * Param.HorasUteisSOV * DispSOV *
      Param.ProdutividadeSOV +
    Param.QtdBAE * Param.HorasUteisBAE * DispBAE *
      Param.ProdutividadeBAE +
    Param.QtdAQUA * Param.HorasUteisAQUA * DispAQUA *
      Param.ProdutividadeAQUA;

  CustoOperacionalTotal :=
    (Param.QtdSurfer * Param.CustoSurfer +
     Param.QtdSOV * Param.CustoSOV +
     Param.QtdBAE * Param.CustoBAE +
     Param.QtdAQUA * Param.CustoAQUA) * DiasAnalise;

  CustoMobilizacaoTotal :=
    Param.QtdSurfer * Param.MobilizacaoSurfer +
    Param.QtdSOV * Param.MobilizacaoSOV +
    Param.QtdBAE * Param.MobilizacaoBAE +
    Param.QtdAQUA * Param.MobilizacaoAQUA;

  if Param.Backlog <= 0 then
    Prazo := 0
  else if MovPorDia <= 0 then
    Prazo := -1
  else
    Prazo := Ceil(Param.Backlog / MovPorDia);

  BacklogRestante := Max(0, Param.Backlog - Trunc(MovPorDia * DiasAnalise));
  BacklogAtendido := Max(0, Param.Backlog - BacklogRestante);

  if Param.Backlog > 0 then
    PercentualBacklogAtendido := (BacklogAtendido / Param.Backlog) * 100
  else
    PercentualBacklogAtendido := 100.0;

  Result.PrazoEstimado := Prazo;
  Result.MovimentacoesPorDia := MovPorDia;
  Result.HorasPlanejadasPorDia := HorasPlanejadasDia;
  Result.HorasEfetivasPorDia := HorasEfetivasDia;
  Result.CustoOperacionalTotal := CustoOperacionalTotal;
  Result.CustoMobilizacaoTotal := CustoMobilizacaoTotal;
  Result.CustoTotal := CustoOperacionalTotal + CustoMobilizacaoTotal;
  Result.HorasUteisGeradas := HorasEfetivasDia * DiasAnalise;

  if Result.HorasUteisGeradas > 0 then
    Result.CustoPorHoraUtil := Result.CustoTotal / Result.HorasUteisGeradas
  else
    Result.CustoPorHoraUtil := 0;

  if BacklogAtendido > 0 then
    Result.CustoPorMovimentacao := Result.CustoTotal / BacklogAtendido
  else
    Result.CustoPorMovimentacao := 0;

  if HorasPlanejadasDia > 0 then
    Result.IndiceResiliencia := (HorasEfetivasDia / HorasPlanejadasDia) * 100
  else
    Result.IndiceResiliencia := 0;

  Result.BacklogAtendido := BacklogAtendido;
  Result.BacklogRestante := BacklogRestante;
  Result.PercentualBacklogAtendido := PercentualBacklogAtendido;
  Result.Relatorio := GerarRelatorio(Result);
end;

end.
