unit uProgramacaoRTUtils;

interface

uses
  System.SysUtils, System.Classes, System.Math, System.Generics.Collections,
  System.Generics.Defaults, System.Variants, DateUtils, Data.DB, Vcl.Grids,
  Vcl.Forms, ADODB, uZucchi, System.StrUtils,Vcl.Dialogs;

  type
    TDadosPlataforma = record
      NomeSAP, CentroCusto, DiagramaRede, OperRede, ElementoPEP, RT_Modal: String;
      booleanProntidao, booleanNaoCriarRT, booleanTurno2: Boolean;
      Latitude, Longitude, CoordX, CoordY: Double;
    end;
  type
    TLegTransbordo = record
      IDProgExec: Integer;
      NomeExecutante: string;
      CodigoSAP: string;
      Origem: string;
      Destino: string;
      TipoEtapaServico: string;
      Funcao: string;
      Empresa: string;
      TipoRT: string;
      ModalOrigem: string;
      ModalDestino: string;
      Turno2Destino: Boolean;
      HoraIdaCalc: TDateTime;
      HoraVoltaCalc: TDateTime;
      TemHoraVolta: Boolean;
      Recolhe: Boolean;
      RecolherPara: string;
      ClasseCalc: string;
    end;

    TLegTransbordoList = class(TList<TLegTransbordo>)
    end;

  TProgramacaoRTUtils = class
  private
    const HoraBase: string = '07:00';

    const VelocidadeKmH: Double = 40;
    const MinTravelMin: Integer = 30;
    const TempoConexaoMin: Integer = 30;
    const DefaultNoDistMin: Integer = 60;
    const TempoPermanenciaMin: Integer = 180;
  public
    class function DadosPlataforma_RT(Plataforma: String): TDadosPlataforma;
    class function GetDistanceBetween(lat1, long1, lat2, long2: Double): Double;
    class function GetDistanceAndTravelMin(
      lat1, long1, lat2, long2: Double;
      out DistKm: Double;
      out TravelMin: Integer
    ): Boolean;
    class procedure ProgramacaoTransbordo(RLGrid: TStringGrid; DataProcura: String);
    class procedure ProgramacaoTransbordosComHora(
      RLGrid: TStringGrid;
      const DataProcura: string;
      const GravarNoBanco: Boolean = False
    );
    class procedure PainelTransbordos(const DataProcura: String);
    class procedure ListaRTEmbarque(RLImpressao: TStringGrid;
      DataProgramacao: TDateTime; Origem: String); static;
  end;

implementation

uses
  untPrincipal, untDataModule, untFrmTabela;

class procedure TProgramacaoRTUtils.ProgramacaoTransbordo(
  RLGrid: TStringGrid; DataProcura: String);
  var
    SQLBase,NomeExecutante1,NomeExecutante2,Origem1,Origem2,Destino1,Destino2,
    TipoEtapaServico1,TipoEtapaServico2,Funcao1,Funcao2,Empresa1,Empresa2: String;
    linha: Integer;
    POrigem1,PDestino1,POrigem2,PDestino2: TDadosPlataforma;
    Distancia1,Distancia2: Double;
begin
  RLGrid.RowCount:= 1;
  RLGrid.ColCount:= 8;
  RLGrid.Rows[0].Clear;
  RLGrid.Cells[0,0]:= 'Origem';
  RLGrid.Cells[1,0]:= 'Destino';
  RLGrid.Cells[2,0]:= 'Nome do Executante';
  RLGrid.Cells[3,0]:= 'Tipo de Etapa de Serviço';
  RLGrid.Cells[4,0]:= 'Função';
  RLGrid.Cells[5,0]:= 'Empresa';
  RLGrid.Cells[6,0]:= 'Hora Início';
  RLGrid.Cells[7,0]:= 'Distânica (km)';
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+
  ')AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado")) ORDER BY NomeExecutante,txtTipoEtapaServico,Origem DESC;';
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  NomeExecutante1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('NomeExecutante').AsString);
  Origem1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Origem').AsString);
  Destino1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtDestino').AsString);
  TipoEtapaServico1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtTipoEtapaServico').AsString);
  Funcao1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Funcao').AsString);
  Empresa1:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Empresa').AsString);
  POrigem1:= DadosPlataforma_RT(Origem1);
  PDestino1:= DadosPlataforma_RT(Destino1);

  Distancia1:= RoundTo(GetDistanceBetween(
    POrigem1.Latitude,POrigem1.Longitude,
    PDestino1.Latitude,PDestino1.Longitude),-3);


  FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    NomeExecutante2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('NomeExecutante').AsString);
    Origem2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Origem').AsString);
    Destino2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtDestino').AsString);
    TipoEtapaServico2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtTipoEtapaServico').AsString);
    Funcao2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Funcao').AsString);
    Empresa2:= TRIM(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Empresa').AsString);
    POrigem2:= DadosPlataforma_RT(Origem2);
    PDestino2:= DadosPlataforma_RT(Destino2);

    Distancia2:= RoundTo(GetDistanceBetween(
      POrigem2.Latitude,POrigem2.Longitude,
      PDestino2.Latitude,PDestino2.Longitude),-3);

    if ((NomeExecutante1 = NomeExecutante2)AND(NomeExecutante1<>'')) then
    begin
      linha:= RLGrid.RowCount;
      RLGrid.RowCount:= RLGrid.RowCount+1;
      RLGrid.Cells[0,linha]:= Origem1;
      RLGrid.Cells[1,linha]:= Destino1;
      RLGrid.Cells[2,linha]:= NomeExecutante1;
      RLGrid.Cells[3,linha]:= TipoEtapaServico1;
      RLGrid.Cells[4,linha]:= Funcao1;
      RLGrid.Cells[5,linha]:= Empresa1;
      RLGrid.Cells[7,linha]:= FormatFloat('0.00km',Distancia1);
      //========================================================
      linha:= RLGrid.RowCount;
      RLGrid.RowCount:= RLGrid.RowCount+1;
      RLGrid.Cells[0,linha]:= Origem2;
      RLGrid.Cells[1,linha]:= Destino2;
      RLGrid.Cells[2,linha]:= NomeExecutante2;
      RLGrid.Cells[3,linha]:= TipoEtapaServico2;
      RLGrid.Cells[4,linha]:= Funcao2;
      RLGrid.Cells[5,linha]:= Empresa2;
      RLGrid.Cells[7,linha]:= FormatFloat('0.00km',Distancia2);
      //========================================================
    end;
    NomeExecutante1:= NomeExecutante2;
    Origem1:= Origem2;
    Destino1:= Destino2;
    TipoEtapaServico1:= TipoEtapaServico2;
    Funcao1:= Funcao2;
    Empresa1:= Empresa2;
    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  end;
  AutoFitGrade(RLGrid);
  try
    RLGrid.FixedRows:= 1;
  except
  end;
end;

class procedure TProgramacaoRTUtils.ProgramacaoTransbordosComHora(
  RLGrid: TStringGrid;
  const DataProcura: string;
  const GravarNoBanco: Boolean = False
);
var
  Conn: TADOConnection;
  QSel: TADOQuery;
  QEdt: TADOQuery;

  SLKeys: TStringList;

  DataDia, DtIni, DtFimMais1: TDateTime;
  ChaveExec: string;
  Legs, Ordered: TLegTransbordoList;
  Leg: TLegTransbordo;
  Dict: TObjectDictionary<string, TLegTransbordoList>;
  Linha: Integer;
  I, J: Integer;

  HoraInicioTurno1, HoraVoltaTurno1: TDateTime;
  GrupoTransbordo: string;
  BaseRetorno: string;

  procedure LimparMarcacoesTransbordoDoDia;
  var
    QReset: TADOQuery;
  begin
    if not GravarNoBanco then
      Exit;

    QReset := TADOQuery.Create(nil);
    try
      QReset.Connection := Conn;
      QReset.ParamCheck := True;
      QReset.SQL.Text :=
        'UPDATE tblProgramacaoExecutante ' +
        'SET ' +
        '  RT_Transbordo = False, ' +
        '  RT_TransbordoAereo = False, ' +
        '  RT_GrupoTransbordo = Null, ' +
        '  RT_BaseRetornoTransbordo = Null, ' +
        '  RT_PrimeiraOrigemTransbordo = Null, ' +
        '  RT_PrimeiroDestinoTransbordo = Null, ' +
        '  RT_SeqTransbordo = Null, ' +
        '  booleanRecolhimento = False, ' +
        '  RecolherPara = Null, ' +
        '  RT_HoraVolta = Null, ' +
        '  DataVolta = Null ' +
        'WHERE idProgramacaoExecutante IN ( ' +
        '  SELECT pe.idProgramacaoExecutante ' +
        '  FROM tblProgramacaoDiaria pd ' +
        '  INNER JOIN tblProgramacaoExecutante pe ' +
        '    ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
        '  WHERE pd.DataProgramacao >= :DtIni ' +
        '    AND pd.DataProgramacao < :DtFimMais1 ' +
        '    AND pe.RT_Transbordo = True ' +
        ')';

      QReset.Parameters.ParamByName('DtIni').DataType := ftDateTime;
      QReset.Parameters.ParamByName('DtFimMais1').DataType := ftDateTime;
      QReset.Parameters.ParamByName('DtIni').Value := DtIni;
      QReset.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;

      QReset.ExecSQL;
    finally
      QReset.Free;
    end;
  end;

  function ParseDateFlexible(const S: string; out ADt: TDateTime): Boolean;
  var
    SS: string;
    FSBR, FSUS: TFormatSettings;
  begin
    SS := Trim(StringReplace(S, '#', '', [rfReplaceAll]));

    FSBR := TFormatSettings.Create('pt-BR');
    FSUS := TFormatSettings.Create('en-US');

    Result :=
      TryStrToDate(SS, ADt, FSBR) or
      TryStrToDate(SS, ADt, FSUS);

    if Result then
      ADt := Trunc(ADt);
  end;

  function TimeFromHHNN(const ADate: TDateTime; const HHNN, ADefault: string): TDateTime;
  var
    T: TDateTime;
    FSBR, FSUS: TFormatSettings;
    S: string;
  begin
    FSBR := TFormatSettings.Create('pt-BR');
    FSUS := TFormatSettings.Create('en-US');

    S := Trim(HHNN);
    if S = '' then
      S := ADefault;

    if (not TryStrToTime(S, T, FSBR)) and (not TryStrToTime(S, T, FSUS)) then
      T := EncodeTime(7, 0, 0, 0);

    Result := Trunc(ADate) + Frac(T);
  end;

  function MakeExecKey(const ANome, ACodigoSAP: string): string;
  begin
    Result := UpperCase(Trim(ANome)) + '|' + UpperCase(Trim(ACodigoSAP));
  end;

  function DiffMinutes(const AStart, AEnd: TDateTime): Double;
  begin
    Result := (AEnd - AStart) * 24 * 60;
  end;

  function CalcClasseTransbordo(const ATipoRT: string): string;
  begin
    if SameText(Trim(ATipoRT), 'R7') then
      Result := 'COM'
    else
      Result := 'TR';
  end;

  function OrderLegsByChain(const L: TLegTransbordoList): TLegTransbordoList;
  var
    DestSet: TStringList;
    Used: array of Boolean;
    HeadIdx, CurrIdx, NextIdx, K: Integer;
    CurrDest: string;
  begin
    Result := TLegTransbordoList.Create;
    DestSet := TStringList.Create;
    try
      DestSet.Sorted := True;
      DestSet.Duplicates := dupIgnore;
      DestSet.CaseSensitive := False;

      for K := 0 to L.Count - 1 do
        if Trim(L[K].Destino) <> '' then
          DestSet.Add(Trim(L[K].Destino));

      HeadIdx := -1;
      for K := 0 to L.Count - 1 do
      begin
        if DestSet.IndexOf(Trim(L[K].Origem)) < 0 then
        begin
          HeadIdx := K;
          Break;
        end;
      end;

      if HeadIdx < 0 then
        HeadIdx := 0;

      SetLength(Used, L.Count);
      CurrIdx := HeadIdx;

      while (CurrIdx >= 0) and (CurrIdx < L.Count) and (not Used[CurrIdx]) do
      begin
        Used[CurrIdx] := True;
        Result.Add(L[CurrIdx]);

        CurrDest := Trim(L[CurrIdx].Destino);
        NextIdx := -1;

        for K := 0 to L.Count - 1 do
        begin
          if (not Used[K]) and SameText(Trim(L[K].Origem), CurrDest) then
          begin
            NextIdx := K;
            Break;
          end;
        end;

        CurrIdx := NextIdx;
      end;

      for K := 0 to L.Count - 1 do
        if not Used[K] then
          Result.Add(L[K]);
    finally
      DestSet.Free;
    end;
  end;

  function IsTransbordoReal(const L: TLegTransbordoList): Boolean;
  var
    K: Integer;
  begin
    Result := L.Count >= 2;
    if not Result then
      Exit;

    for K := 0 to L.Count - 1 do
    begin
      if Trim(L[K].Origem) = '' then
        Exit(False);
      if Trim(L[K].Destino) = '' then
        Exit(False);
      if SameText(Trim(L[K].Origem), Trim(L[K].Destino)) then
        Exit(False);
    end;

    for K := 0 to L.Count - 2 do
      if not SameText(Trim(L[K].Destino), Trim(L[K + 1].Origem)) then
        Exit(False);

    Result := True;
  end;

  function IsTransbordoAereoEspecial(const L: TLegTransbordoList): Boolean;
  begin
    Result := False;
    if L.Count < 2 then
      Exit;

    if SameText(Trim(L[0].ModalOrigem), 'A') then
      Result := True;
  end;

  procedure SetupGrid;
  begin
    RLGrid.RowCount := 1;
    RLGrid.ColCount := 10;
    RLGrid.Rows[0].Clear;

    RLGrid.Cells[0,0] := 'Origem';
    RLGrid.Cells[1,0] := 'Destino';
    RLGrid.Cells[2,0] := 'Nome do Executante';
    RLGrid.Cells[3,0] := 'Tipo de Etapa de Serviço';
    RLGrid.Cells[4,0] := 'Função';
    RLGrid.Cells[5,0] := 'Empresa';
    RLGrid.Cells[6,0] := 'Hora Ida';
    RLGrid.Cells[7,0] := 'Hora Volta';
    RLGrid.Cells[8,0] := 'Recolher Para';
    RLGrid.Cells[9,0] := 'Seq.';
  end;

  procedure AddRowGrid(const Lg: TLegTransbordo; const ASeq: Integer);
  begin
    Linha := RLGrid.RowCount;
    RLGrid.RowCount := Linha + 1;

    RLGrid.Cells[0,Linha] := Lg.Origem;
    RLGrid.Cells[1,Linha] := Lg.Destino;
    RLGrid.Cells[2,Linha] := Lg.NomeExecutante;
    RLGrid.Cells[3,Linha] := Lg.TipoEtapaServico;
    RLGrid.Cells[4,Linha] := Lg.Funcao;
    RLGrid.Cells[5,Linha] := Lg.Empresa;
    RLGrid.Cells[6,Linha] := FormatDateTime('hh:nn', Lg.HoraIdaCalc);

    if Lg.TemHoraVolta then
      RLGrid.Cells[7,Linha] := FormatDateTime('hh:nn', Lg.HoraVoltaCalc)
    else
      RLGrid.Cells[7,Linha] := '';

    RLGrid.Cells[8,Linha] := Lg.RecolherPara;
    RLGrid.Cells[9,Linha] := IntToStr(ASeq);
  end;

  procedure PrepararUpdater;
  begin
    if not GravarNoBanco then
      Exit;

    QEdt := TADOQuery.Create(nil);
    QEdt.Connection := Conn;
    QEdt.ParamCheck := True;
    QEdt.SQL.Text :=
      'SELECT * FROM tblProgramacaoExecutante ' +
      'WHERE idProgramacaoExecutante = :ID';

    QEdt.Parameters.ParamByName('ID').DataType := ftInteger;
  end;

  procedure AplicarNoBanco(
    const Lg: TLegTransbordo;
    const ASeq: Integer;
    const ATransbordo, ATransbordoAereo: Boolean;
    const APrimeiraOrigem, APrimeiroDestino, AGrupoTransbordo,
    ABaseRetorno: string
  );
  var
    F: TField;
    SHoraVolta: string;
  begin
    if not GravarNoBanco then
      Exit;

    if Lg.TemHoraVolta then
      SHoraVolta := FormatDateTime('hh:nn', Lg.HoraVoltaCalc)
    else
      SHoraVolta := '';

    QEdt.Close;
    QEdt.Parameters.ParamByName('ID').Value := Lg.IDProgExec;
    QEdt.Open;

    if QEdt.Eof then
      raise Exception.Create(
        'Registro não encontrado em tblProgramacaoExecutante. ID=' +
        IntToStr(Lg.IDProgExec)
      );

    QEdt.Edit;
    try
      F := QEdt.FindField('RT_Transbordo');
      if F <> nil then
        F.AsBoolean := ATransbordo;

      F := QEdt.FindField('RT_TransbordoAereo');
      if F <> nil then
        F.AsBoolean := ATransbordoAereo;

      F := QEdt.FindField('RT_PrimeiraOrigemTransbordo');
      if F <> nil then
        F.AsString := Copy(Trim(APrimeiraOrigem), 1, F.Size);

      F := QEdt.FindField('RT_PrimeiroDestinoTransbordo');
      if F <> nil then
        F.AsString := Copy(Trim(APrimeiroDestino), 1, F.Size);

      F := QEdt.FindField('RT_GrupoTransbordo');
      if F <> nil then
        F.AsString := Copy(Trim(AGrupoTransbordo), 1, F.Size);

      F := QEdt.FindField('RT_BaseRetornoTransbordo');
      if F <> nil then
        F.AsString := Copy(Trim(ABaseRetorno), 1, F.Size);

      F := QEdt.FindField('RT_SeqTransbordo');
      if F <> nil then
        F.AsInteger := ASeq;

      F := QEdt.FindField('booleanRecolhimento');
      if F <> nil then
        F.AsBoolean := Lg.Recolhe;

      F := QEdt.FindField('RecolherPara');
      if F <> nil then
      begin
        if Trim(Lg.RecolherPara) <> '' then
          F.AsString := Copy(Trim(Lg.RecolherPara), 1, F.Size)
        else
          F.Clear;
      end;

      F := QEdt.FindField('RT_HoraIda');
      if F <> nil then
        F.AsString := FormatDateTime('hh:nn', Lg.HoraIdaCalc);

      F := QEdt.FindField('RT_HoraVolta');
      if F <> nil then
      begin
        if SHoraVolta <> '' then
          F.AsString := SHoraVolta
        else
          F.Clear;
      end;

      F := QEdt.FindField('DataVolta');
      if F <> nil then
      begin
        if Lg.TemHoraVolta then
          F.AsDateTime := Trunc(Lg.HoraVoltaCalc)
        else
          F.Clear;
      end;

      F := QEdt.FindField('RT_Classe');
      if F <> nil then
        F.AsString := Trim(Lg.ClasseCalc);

      QEdt.Post;
    except
      on E: Exception do
      begin
        if QEdt.State in dsEditModes then
          QEdt.Cancel;

        raise Exception.Create(
          'Erro ao atualizar transbordo.' + sLineBreak +
          'ID: ' + IntToStr(Lg.IDProgExec) + sLineBreak +
          'Nome: ' + Lg.NomeExecutante + sLineBreak +
          'Origem: ' + Lg.Origem + sLineBreak +
          'Destino: ' + Lg.Destino + sLineBreak +
          'Seq: ' + IntToStr(ASeq) + sLineBreak +
          'Grupo: ' + AGrupoTransbordo + sLineBreak +
          'BaseRetorno: ' + ABaseRetorno + sLineBreak +
          'HoraIda: ' + FormatDateTime('hh:nn', Lg.HoraIdaCalc) + sLineBreak +
          'HoraVolta: ' + SHoraVolta + sLineBreak +
          'RecolherPara: ' + Lg.RecolherPara + sLineBreak +
          'Classe: ' + Lg.ClasseCalc + sLineBreak +
          'Erro: ' + E.Message
        );
      end;
    end;
  end;

  procedure DistribuirHorariosTransbordoNormal(var L: TLegTransbordoList);
  var
    NumPernas, NumIntervalos: Integer;
    PassoMin, TotalMin: Double;
    K: Integer;
    HoraBaseLeg, HoraFimTurno: TDateTime;
  begin
    NumPernas := L.Count;
    if NumPernas <= 0 then
      Exit;

    NumIntervalos := NumPernas;
    TotalMin := DiffMinutes(HoraInicioTurno1, HoraVoltaTurno1);

    if NumIntervalos <= 0 then
      NumIntervalos := 1;

    PassoMin := TotalMin / NumIntervalos;

    for K := 0 to NumPernas - 1 do
    begin
      Leg := L[K];

      HoraBaseLeg := IncMinute(HoraInicioTurno1, Round(PassoMin * K));
      Leg.HoraIdaCalc := HoraBaseLeg;
      Leg.TemHoraVolta := False;
      Leg.HoraVoltaCalc := 0;
      Leg.Recolhe := False;
      Leg.RecolherPara := '';
      Leg.ClasseCalc := CalcClasseTransbordo(Leg.TipoRT);

      if K = NumPernas - 1 then
      begin
        HoraFimTurno := HoraVoltaTurno1;
        Leg.TemHoraVolta := True;
        Leg.HoraVoltaCalc := HoraFimTurno;
        Leg.Recolhe := True;
        Leg.RecolherPara := Trim(L[0].Origem);
      end;

      L[K] := Leg;
    end;
  end;

  procedure DistribuirHorariosTransbordoAereoEspecial(var L: TLegTransbordoList);
  var
    NumPernas: Integer;
    PassoMin, TotalMin: Double;
    K: Integer;
    HoraRetornoPrimeira, HoraRetornoUltima, HoraBaseLeg: TDateTime;
  begin
    NumPernas := L.Count;
    if NumPernas <= 0 then
      Exit;

    HoraRetornoPrimeira := HoraVoltaTurno1;
    HoraRetornoUltima := IncHour(HoraRetornoPrimeira, -2);

    if HoraRetornoUltima <= HoraInicioTurno1 then
      HoraRetornoUltima := IncHour(HoraInicioTurno1, 1);

    TotalMin := DiffMinutes(HoraInicioTurno1, HoraRetornoUltima);
    PassoMin := TotalMin / NumPernas;

    for K := 0 to NumPernas - 1 do
    begin
      Leg := L[K];

      HoraBaseLeg := IncMinute(HoraInicioTurno1, Round(PassoMin * K));
      Leg.HoraIdaCalc := HoraBaseLeg;
      Leg.TemHoraVolta := False;
      Leg.HoraVoltaCalc := 0;
      Leg.Recolhe := False;
      Leg.RecolherPara := '';
      Leg.ClasseCalc := CalcClasseTransbordo(Leg.TipoRT);

      if K = 0 then
      begin
        Leg.TemHoraVolta := True;
        Leg.HoraVoltaCalc := HoraRetornoPrimeira;
        Leg.Recolhe := True;
        Leg.RecolherPara := Trim(L[0].Origem);
      end;

      if K = NumPernas - 1 then
      begin
        Leg.TemHoraVolta := True;
        Leg.HoraVoltaCalc := HoraRetornoUltima;
        Leg.Recolhe := True;
        Leg.RecolherPara := Trim(L[0].Destino);
      end;

      L[K] := Leg;
    end;
  end;

var
  POrigem, PDestino: TDadosPlataforma;
  EhTransbordo, EhAereoEspecial: Boolean;
begin
  SetupGrid;

  QSel := nil;
  QEdt := nil;

  if not ParseDateFlexible(DataProcura, DataDia) then
    raise Exception.Create('Data inválida em ProgramacaoTransbordosComHora: ' + DataProcura);

  DtIni := Trunc(DataDia);
  DtFimMais1 := IncDay(DtIni, 1);

  if not FrmDataModule.ADOQueryConfigRT.Active then
    FrmDataModule.ADOQueryConfigRT.Active := True;

  HoraInicioTurno1 := TimeFromHHNN(
    DataDia,
    Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraInicio_Turno1').AsString),
    '07:00'
  );

  HoraVoltaTurno1 := TimeFromHHNN(
    DataDia,
    Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraVolta_Turno1').AsString),
    '17:00'
  );

  Conn := FrmDataModule.ADOConnectionColibri;
  if Conn = nil then
    raise Exception.Create('Conexão ADO não encontrada.');

  if not Conn.Connected then
    Conn.Connected := True;

  Dict := TObjectDictionary<string, TLegTransbordoList>.Create([doOwnsValues]);
  SLKeys := TStringList.Create;
  try
    SLKeys.Sorted := True;
    SLKeys.Duplicates := dupIgnore;
    SLKeys.CaseSensitive := False;

    QSel := TADOQuery.Create(nil);
    QSel.Connection := Conn;
    QSel.ParamCheck := True;
    QSel.SQL.Text :=
      'SELECT pd.DataProgramacao, pd.txtDestino, pd.txtTipoEtapaServico, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      '  AND pe.StatusProgramacao = ''Aprovado'' ' +
      'ORDER BY pe.NomeExecutante, pe.CodigoSAP, pe.Origem, pd.txtDestino';

    QSel.Parameters.ParamByName('DtIni').DataType := ftDateTime;
    QSel.Parameters.ParamByName('DtFimMais1').DataType := ftDateTime;

    QSel.Parameters.ParamByName('DtIni').Value := DtIni;
    QSel.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;

    QSel.Open;

    while not QSel.Eof do
    begin
      ChaveExec := MakeExecKey(
        Trim(QSel.FieldByName('NomeExecutante').AsString),
        Trim(QSel.FieldByName('CodigoSAP').AsString)
      );

      if not Dict.TryGetValue(ChaveExec, Legs) then
      begin
        Legs := TLegTransbordoList.Create;
        Dict.Add(ChaveExec, Legs);
      end;

      POrigem := DadosPlataforma_RT(Trim(QSel.FieldByName('Origem').AsString));
      PDestino := DadosPlataforma_RT(Trim(QSel.FieldByName('txtDestino').AsString));

      Leg.IDProgExec := QSel.FieldByName('idProgramacaoExecutante').AsInteger;
      Leg.NomeExecutante := Trim(QSel.FieldByName('NomeExecutante').AsString);
      Leg.CodigoSAP := Trim(QSel.FieldByName('CodigoSAP').AsString);
      Leg.Origem := Trim(QSel.FieldByName('Origem').AsString);
      Leg.Destino := Trim(QSel.FieldByName('txtDestino').AsString);
      Leg.TipoEtapaServico := Trim(QSel.FieldByName('txtTipoEtapaServico').AsString);
      Leg.Funcao := Trim(QSel.FieldByName('Funcao').AsString);
      Leg.Empresa := Trim(QSel.FieldByName('Empresa').AsString);
      Leg.TipoRT := Trim(QSel.FieldByName('RT_Tipo').AsString);
      Leg.ModalOrigem := Trim(POrigem.RT_Modal);
      Leg.ModalDestino := Trim(PDestino.RT_Modal);
      Leg.Turno2Destino := PDestino.booleanTurno2;
      Leg.HoraIdaCalc := 0;
      Leg.HoraVoltaCalc := 0;
      Leg.TemHoraVolta := False;
      Leg.Recolhe := False;
      Leg.RecolherPara := '';
      Leg.ClasseCalc := '';

      Legs.Add(Leg);
      SLKeys.Add(ChaveExec);

      QSel.Next;
    end;

    PrepararUpdater;

    if GravarNoBanco then
      Conn.BeginTrans;
    try
      LimparMarcacoesTransbordoDoDia;


      for I := 0 to SLKeys.Count - 1 do
      begin
        ChaveExec := SLKeys[I];
        Legs := Dict[ChaveExec];

        if Legs.Count < 2 then
          Continue;

        Ordered := OrderLegsByChain(Legs);
        try
          EhTransbordo := IsTransbordoReal(Ordered);
          if not EhTransbordo then
            Continue;

          if Ordered[0].Turno2Destino then
            Continue;

          EhAereoEspecial := IsTransbordoAereoEspecial(Ordered);

          if EhAereoEspecial then
            DistribuirHorariosTransbordoAereoEspecial(Ordered)
          else
            DistribuirHorariosTransbordoNormal(Ordered);

          if EhAereoEspecial then
            BaseRetorno := Trim(Ordered[0].Destino)
          else
            BaseRetorno := Trim(Ordered[0].Origem);

          GrupoTransbordo := Copy(
            FormatDateTime('yyyymmdd', DataDia) + '|' +
            Trim(Ordered[0].NomeExecutante) + '|' +
            Trim(Ordered[0].CodigoSAP) + '|' +
            Trim(Ordered[0].Origem) + '|' +
            Trim(Ordered[0].Destino),
            1,
            80
          );

          for J := 0 to Ordered.Count - 1 do
          begin
            AddRowGrid(Ordered[J], J + 1);

            AplicarNoBanco(
              Ordered[J],
              J + 1,
              True,
              EhAereoEspecial,
              Ordered[0].Origem,
              Ordered[0].Destino,
              GrupoTransbordo,
              BaseRetorno
            );
          end;
        finally
          Ordered.Free;
        end;
      end;

      if GravarNoBanco then
        Conn.CommitTrans;
    except
      if GravarNoBanco then
        Conn.RollbackTrans;
      raise;
    end;

    AutoFitGrade(RLGrid);
    try
      RLGrid.FixedRows := 1;
    except
    end;

  finally
    if Assigned(QEdt) then
      QEdt.Free;
    if Assigned(QSel) then
      QSel.Free;
    SLKeys.Free;
    Dict.Free;
  end;
end;

class procedure TProgramacaoRTUtils.PainelTransbordos(const DataProcura: string);
begin
  if not Assigned(FrmTabela) then
    Application.CreateForm(TFrmTabela, FrmTabela);

  FrmTabela.PanelTitulo.Caption := 'Transbordo de Executantes';
  FrmTabela.Caption := 'Transbordo de Executantes';

  ProgramacaoTransbordosComHora(FrmTabela.RLTabela, DataProcura, False);
  FrmTabela.Show;
end;

class function TProgramacaoRTUtils.DadosPlataforma_RT(
  Plataforma: String): TDadosPlataforma;
var
  SQLBase: String;
begin
  SQLBase := 'SELECT tblPlataforma.* FROM tblPlataforma ' +
             'WHERE Plataforma = "' + Plataforma + '";';

  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Text := SQLBase;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;

  Result.NomeSAP := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('NomeSAP').AsString);
  Result.CentroCusto := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('CentroCusto').AsString);
  Result.DiagramaRede := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('DiagramaRede').AsString);
  Result.OperRede := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('OperRede').AsString);
  Result.ElementoPEP := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('ElementoPEP').AsString);
  Result.RT_Modal := Trim(FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('RT_Modal').AsString);
  Result.booleanProntidao := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('booleanProntidao').AsBoolean;
  Result.booleanNaoCriarRT := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('booleanNaoCriarRT').AsBoolean;
  Result.booleanTurno2 := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('booleanTurno2').AsBoolean;
  Result.Latitude := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('Latitude').AsFloat;
  Result.Longitude := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('Longitude').AsFloat;
  Result.CoordX := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('CoordX').AsFloat;
  Result.CoordY := FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('CoordY').AsFloat;
end;


class procedure TProgramacaoRTUtils.ListaRTEmbarque(RLImpressao: TStringGrid;
  DataProgramacao: TDateTime; Origem: String);
  var
    DataProcura,SQLBase,CPF,CodigoSAP,PRG_TM,
    SQLString: String;
    numLinhas: Integer;
begin
  //Procurar todos os Executantes Programados para a DataProcura especifica
  if InputQuery('Filtrar Origem',
  'Entre com as origens a serem filtradas separadas por ";"',Origem) then
  begin
    SQLString:= '';
    if Origem <> '' then
      SQLString:= FrmPrincipal.palavraBusca(SQLString,'Origem','Contem',Origem,'AND');
    if SQLString <> '' then
      SQLString:= 'AND'+SQLString;
    DataProcura:= DateToStr(DataProgramacao);
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+
    ')AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado")'+SQLString+
    ') ORDER BY NomeExecutante,Empresa,txtDestino,txtTipoEtapaServico;';
    FrmDataModule.ADOQueryTemporarioDBColibri.Close;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
    FrmDataModule.ADOQueryTemporarioDBColibri.Open;
    //==============================================
    RLImpressao.FixedRows:= 0;
    RLImpressao.RowCount:= 1;
    RLImpressao.ColCount:= 7;
    RLImpressao.Cells[0,0]:= 'Empresa';
    RLImpressao.Cells[1,0]:= 'Nome do Executante';
    RLImpressao.Cells[2,0]:= 'Tipo de Etapa de Serviço';
    RLImpressao.Cells[3,0]:= 'Código SAP';
    RLImpressao.Cells[4,0]:= 'CPF';
    RLImpressao.Cells[5,0]:= 'RT';
    RLImpressao.Cells[6,0]:= 'Status';
    //==============================================
    FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount,
    'Carregando executantes programados...');
    while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
    begin
      if FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
      FieldByName('InseridoProgramacaoTransporte').AsBoolean then
        PRG_TM:= 'TM: '
      else
        PRG_TM:= 'PRG: ';

      CPF:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('Documento').AsString;
      CodigoSAP:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('CodigoSAP').AsString;
      if (CPF = '') then
        CPF:= '000.000.000-00';

      numLinhas:= RLImpressao.RowCount;
      RLImpressao.RowCount:= numLinhas+1;
      //Preencher valores
      RLImpressao.Cells[0,numLinhas]:= FrmDataModule.
      DataSourceTemporarioDBColibri.DataSet.FieldByName('Empresa').AsString;
      RLImpressao.Cells[1,numLinhas]:= FrmPrincipal.TextoMaiusculo((FrmDataModule.
      DataSourceTemporarioDBColibri.DataSet.FieldByName('NomeExecutante').AsString));
      RLImpressao.Cells[2,numLinhas]:= FrmDataModule.
      DataSourceTemporarioDBColibri.DataSet.FieldByName('txtTipoEtapaServico').AsString;
      RLImpressao.Cells[3,numLinhas]:= (CodigoSAP);
      RLImpressao.Cells[4,numLinhas]:= (CPF);
      RLImpressao.Cells[5,numLinhas]:= FrmDataModule.
      DataSourceTemporarioDBColibri.DataSet.FieldByName('RT').AsString;
      RLImpressao.Cells[6,numLinhas]:= PRG_TM+FrmDataModule.
      DataSourceTemporarioDBColibri.DataSet.FieldByName('Origem').AsString+'->'+
      FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('txtDestino').AsString;
      //===========================================================================
      FrmPrincipal.ProgressBarIncremento(1);
      FrmDataModule.ADOQueryTemporarioDBColibri.Next;
    end;
    //==============================================
    FrmPrincipal.clasifica(RLImpressao,0,false);
    AutoFitGrade(RLImpressao);
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

class function TProgramacaoRTUtils.GetDistanceBetween(
  lat1, long1, lat2, long2: Double): Double;
var
  F, G, L, dLon: Double;
  SF, SG, SL, CF, CG, CL: Double;
  W1, W2, S, C: Double;
  O, R, D: Double;
  H1, H2: Double;
  ff: Double;
const
  A_EQUATOR_KM = 6378.137;
begin
  Result := 0;

  if (Abs(lat1 - lat2) < 1e-12) and (Abs(long1 - long2) < 1e-12) then
    Exit;

  ff := 1 / 298.257223563;

  dLon := long1 - long2;
  while dLon > 180 do dLon := dLon - 360;
  while dLon < -180 do dLon := dLon + 360;

  F := (lat1 + lat2) / 2;
  G := (lat1 - lat2) / 2;
  L := dLon / 2;

  SF := Sin(F * Pi / 180);
  SG := Sin(G * Pi / 180);
  SL := Sin(L * Pi / 180);
  CF := Cos(F * Pi / 180);
  CG := Cos(G * Pi / 180);
  CL := Cos(L * Pi / 180);

  W1 := Sqr(SG * CL);
  W2 := Sqr(CF * SL);
  S := W1 + W2;

  W1 := Sqr(CG * CL);
  W2 := Sqr(SF * SL);
  C := W1 + W2;

  if (S <= 0) or (C <= 0) then
    Exit;

  O := ArcTan2(Sqrt(S), Sqrt(C));
  if Abs(O) < 1e-15 then
    Exit;

  R := Sqrt(S * C) / O;
  D := 2 * O * A_EQUATOR_KM;

  H1 := (3 * R - 1) / (2 * C);
  H2 := (3 * R + 1) / (2 * S);

  W1 := Sqr(SF * CG) * H1 * ff + 1;
  W2 := Sqr(CF * SG) * H2 * ff;

  Result := D * (W1 - W2);
end;

class function TProgramacaoRTUtils.GetDistanceAndTravelMin(
  lat1, long1, lat2, long2: Double;
  out DistKm: Double;
  out TravelMin: Integer
): Boolean;

  function IsValidCoord(const Lat, Lon: Double): Boolean;
  begin
    Result :=
      (not IsNan(Lat)) and (not IsNan(Lon)) and
      (not IsInfinite(Lat)) and (not IsInfinite(Lon)) and
      (Abs(Lat) > 1e-8) and (Abs(Lon) > 1e-8) and
      (Lat >= -90) and (Lat <= 90) and
      (Lon >= -180) and (Lon <= 180);
  end;

begin
  DistKm := 0;
  TravelMin := DefaultNoDistMin;

  if (not IsValidCoord(lat1, long1)) or (not IsValidCoord(lat2, long2)) then
    Exit(False);

  if (Abs(lat1 - lat2) < 1e-12) and (Abs(long1 - long2) < 1e-12) then
  begin
    DistKm := 0;
    TravelMin := 0;
    Exit(True);
  end;

  DistKm := GetDistanceBetween(lat1, long1, lat2, long2);

  if (DistKm <= 0) or IsNan(DistKm) or IsInfinite(DistKm) then
  begin
    DistKm := 0;
    TravelMin := DefaultNoDistMin;
    Exit(False);
  end;

  if VelocidadeKmH <= 0 then
  begin
    TravelMin := DefaultNoDistMin;
    Exit(True);
  end;

  TravelMin := Ceil((DistKm / VelocidadeKmH) * 60);
  if TravelMin < MinTravelMin then
    TravelMin := MinTravelMin;

  Result := True;
end;

End.
