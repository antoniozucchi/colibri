unit untConsultaExecutantesProgramados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ComCtrls, Vcl.ToolWin, Vcl.CheckLst, Vcl.ExtCtrls, Vcl.Grids, Vcl.DBGrids,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,DateUtils,
  Vcl.ExtDlgs, Vcl.Menus, Vcl.AppEvnts,ComOBJ, uZucchi, ADODB,
  untDBGridFilter, System.StrUtils, Vcl.DBCtrls;

type
  TFrmConsultaExecutantesProgramados = class(TForm)
    PanelTitulo: TPanel;
    ActionManager1: TActionManager;
    actProcurarProgramacaoExecutante: TAction;
    Panel3: TPanel;
    dataInicio: TDateTimePicker;
    Panel4: TPanel;
    dataFim: TDateTimePicker;
    Panel8: TPanel;
    actExcel: TAction;
    actImprmir: TAction;
    actNaoExecutado: TAction;
    actExecutado: TAction;
    actNumDias: TAction;
    actStatusSELECIONADO: TAction;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    DBGridExecutantesProgramados: TFilterDBGrid;
    StatusBarExecutantes: TStatusBar;
    StringGridImpressao: TStringGrid;
    ToolBar1: TToolBar;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    ToolButton3: TToolButton;
    actGRFSalvarImagem: TAction;
    actZoomMais: TAction;
    actZoomMenos: TAction;
    actGRFTitulo: TAction;
    actGRFImprimir: TAction;
    SavePictureDialog1: TSavePictureDialog;
    PopupMenuStatus: TPopupMenu;
    CalcularStatusExecSelecionado1: TMenuItem;
    CalcularStatusExec1: TMenuItem;
    actStatusTODOS: TAction;
    actGraficoTempo: TAction;
    actGraficoPlataforma: TAction;
    actGraficoTipoEtapaServico: TAction;
    actGrafico: TAction;
    actGRFSalvarData: TAction;
    SaveDialog1: TSaveDialog;
    actEtiqueta: TAction;
    ColunasLayout: TStringGrid;
    CheckBoxOrigemDestino: TCheckBox;
    ToolButton11: TToolButton;
    actTransbordo: TAction;
    BitBtn16: TBitBtn;
    actMotivo: TAction;
    TabSheet3: TTabSheet;
    ToolBar2: TToolBar;
    actRelatorioCancelado: TAction;
    actRelatorioMudanca: TAction;
    actRelatorioNaoExecutada: TAction;
    BitBtn1: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    StatusBarRelatorio: TStatusBar;
    actExcelRelatorio: TAction;
    BitBtn8: TBitBtn;
    RLImpressao: TStringGrid;
    RadioGroupTipoEtapaServico: TRadioGroup;
    Splitter1: TSplitter;
    CheckBoxSomenteOrigem: TCheckBox;
    CheckBoxOrigemDestinoDiferente: TCheckBox;
    MemoSAP: TMemo;
    actGerarMultiplasRTs: TAction;
    ToolButton4: TToolButton;
    actConfigurarRT: TAction;
    btnLayout: TToolButton;
    btnClearFiltro: TToolButton;
    ToolButton5: TToolButton;
    btnSelAll: TToolButton;
    btnSelClear: TToolButton;
    BitBtn2: TBitBtn;
    actCancelarRT: TAction;
    TabSheet2: TTabSheet;
    ToolBar3: TToolBar;
    BitBtn10: TBitBtn;
    btnFiltroClearGestaoRT: TToolButton;
    actProcurarGestaoRT: TAction;
    StatusBarGestaoRT: TStatusBar;
    ColunasLayoutGestaoRT: TStringGrid;
    DBGridGestaoRT: TFilterDBGrid;
    actExcluirProgramacaoRT: TAction;
    BitBtn7: TBitBtn;
    procedure actProcurarProgramacaoExecutanteExecute(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actExcelExecute(Sender: TObject);
    procedure DBGridExecutantesProgramadosDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure edtExecutanteKeyPress(Sender: TObject; var Key: Char);
    procedure actImprmirExecute(Sender: TObject);
    procedure dataInicioKeyPress(Sender: TObject; var Key: Char);
    procedure dataFimKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridExecutantesProgramadosCellClick(Column: TColumn);
    procedure actNaoExecutadoExecute(Sender: TObject);
    procedure actExecutadoExecute(Sender: TObject);
    procedure actNumDiasExecute(Sender: TObject);
    procedure actStatusSELECIONADOExecute(Sender: TObject);
    procedure actStatusTODOSExecute(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure DBGridExecutantesProgramadosColumnMoved(Sender: TObject;
      FromIndex, ToIndex: Integer);
    procedure actTransbordoExecute(Sender: TObject);
    procedure actMotivoExecute(Sender: TObject);
    procedure actRelatorioCanceladoExecute(Sender: TObject);
    procedure actRelatorioMudancaExecute(Sender: TObject);
    procedure actRelatorioNaoExecutadaExecute(Sender: TObject);
    procedure actExcelRelatorioExecute(Sender: TObject);
    procedure RLImpressaoFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure actGerarMultiplasRTsExecute(Sender: TObject);
    procedure actConfigurarRTExecute(Sender: TObject);
    procedure actCancelarRTExecute(Sender: TObject);
    procedure actProcurarGestaoRTExecute(Sender: TObject);
    procedure actExcluirProgramacaoRTExecute(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
  type
      TDadosRT = record
        idProgramacaoExecutante, idProgramacaoRT, TipoCusto: Integer;
        DataProgramacao,HoraCobertura: TDateTime;
        RTNumero, TipoRT,DescricaoRT,Requisitante,PessoaContato,TelContato,CentroPlan,GrpPlan,
        MatriculaPax, Documento, Modal,Classe,
        Origem, Destino, DataInicio, HoraInicio, HoraVolta,
        CentroCusto, DiagramaRede, OperRede, ElementoPEP: String;
        RTCobertura, booRTCriada, booleanBateVolta: Boolean;
      end;
  type
      TRT = record
        RT_Numero: String;
        idProgramacaoRT: Integer;
        booRTExiste: Boolean;
      end;
  type
      TCodicoCusto = record
        CentroCusto, DiagramaRede, OperRede, ElementoPEP: String;
      end;
  private
    { Private declarations }
    enderecoColibriRegistro: String;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure RelatorioMotivo(ColunaMotivo: String; TipoFonte:Integer);

    procedure ProcessarRetornoSAP(const LinhaLog: string);
    procedure AtualizarProgramacaoExecutanteComRetornoRT(
      const idProgramacaoExecutante: Integer;
      const MensagemRT, RT_Numero, StatusRT, ErroRT: string);

    function Existe_RT(const Dados: TDadosRT): TRT;
    function ExisteProgramacaoExecutante(DataProcura: TDateTime; Origem, Destino, CodigoSAP, RTNumero: String): Boolean;



    function CriarRegistroTabelaRT(Dados: TDadosRT): Integer;
    procedure AtualizarProgramacaoRTComRetornoRT(
      const idProgramacaoRT: Integer;
      const MensagemRT, RT_Numero, StatusRT, ErroRT: string);

    procedure ParseLogLine(const Linha: string;
      out IdExec, IdRT, Seq: Integer;
      out StatusItem, Msg: string);

    procedure InterpretarMensagemSAP(const Msg: string;
      out RT_Numero, StatusRT, ErroRT: string);

    function CustoExecutante(CodigoSAP: String): TCodicoCusto;
    procedure ExcluirProgramacaoRTSemRT;
    procedure GerarMultiplasRTsArray;


    function NormalizarPlataforma(const Valor: string): string;
    function ExtractOnlyNumbers(const S: string): string;
    procedure Testar_RetornoSAP_SemSAP;
    procedure AtualizarProgramacaoRT_Cancelamento(
      const idProgramacaoRT: Integer; const Cancelada: Boolean; const Mensagem,
      StatusRT, ErroRT: string);
    procedure ProcessarRetornoCancelamentoSAP(const LinhaLog: string);
  public
    ExcelRT,SheetRT: Variant;
    { Public declarations }
  end;

var
  FrmConsultaExecutantesProgramados: TFrmConsultaExecutantesProgramados;

implementation
  uses untPrincipal,untDataModule,untFrmPreview, untFrmConfigRT;
{$R *.dfm}

function TFrmConsultaExecutantesProgramados.ExtractOnlyNumbers(const S: string): string;
var
  i: Integer;
begin
  Result := '';
  for i := 1 to Length(S) do
    if CharInSet(S[i], ['0'..'9']) then
      Result := Result + S[i];
end;

procedure TFrmConsultaExecutantesProgramados.ParseLogLine(
  const Linha: string;
  out IdExec, IdRT, Seq: Integer;
  out StatusItem, Msg: string
);
var
  s: string;
  p, p1, p2, p3, p4: Integer;
begin
  IdExec := 0;
  IdRT   := 0;
  Seq    := 0;
  StatusItem := '';
  Msg    := '';

  // pega só o que vem depois de " | "
  p := Pos(' | ', Linha);
  if p > 0 then
    s := Trim(Copy(Linha, p + 3, MaxInt))
  else
    s := Trim(Linha);

  // tokens: IdExec|IdRT|Seq|Status|Msg
  p1 := Pos('|', s); if p1 = 0 then Exit;
  p2 := PosEx('|', s, p1 + 1); if p2 = 0 then Exit;
  p3 := PosEx('|', s, p2 + 1); if p3 = 0 then Exit;
  p4 := PosEx('|', s, p3 + 1); if p4 = 0 then Exit;

  IdExec := StrToIntDef(Copy(s, 1, p1 - 1), 0);
  IdRT   := StrToIntDef(Copy(s, p1 + 1, p2 - p1 - 1), 0);
  Seq    := StrToIntDef(Copy(s, p2 + 1, p3 - p2 - 1), 0);

  StatusItem := Trim(Copy(s, p3 + 1, p4 - p3 - 1)); // OK / ERRO_FINAL / etc
  Msg        := Trim(Copy(s, p4 + 1, MaxInt));      // "Nota 326... gravada"
end;

procedure TFrmConsultaExecutantesProgramados.InterpretarMensagemSAP(
  const Msg: string;
  out RT_Numero, StatusRT, ErroRT: string);
begin
  RT_Numero := '';
  StatusRT  := '';
  ErroRT    := '';

  // OK típico
  if Pos('Nota', Msg) > 0 then
  begin
    RT_Numero := ExtractOnlyNumbers(Msg); // deve retornar 326422165
    StatusRT  := 'OK';
    Exit;
  end;

  // Erros conhecidos
  if Pos('não está ativo', Msg) > 0 then
  begin
    StatusRT := 'ERRO';
    ErroRT  := 'PAX_INATIVO';
    Exit;
  end;

  if Pos('Contrato obrigatório', Msg) > 0 then
  begin
    StatusRT := 'ERRO';
    ErroRT  := 'CONTRATO_OBRIGATORIO';
    Exit;
  end;

  StatusRT := 'DESCONHECIDO';
end;

procedure TFrmConsultaExecutantesProgramados.ToolButton1Click(Sender: TObject);
begin
Testar_RetornoSAP_SemSAP;
end;

procedure TFrmConsultaExecutantesProgramados.Testar_RetornoSAP_SemSAP;
var
  SL: TStringList;
  i: Integer;
  LogFake: string;
begin
  // Escolha um caminho no PC de DEV
  LogFake := IncludeTrailingPathDelimiter(GetEnvironmentVariable('USERPROFILE')) +
             'Documents\Colibri\sap_log_rts_TESTE.txt';

  // Monta linhas iguais às do log real
  SL := TStringList.Create;
  try
    SL.Add('21/02/2026 18:01:17 | INICIO_LOTE');
    SL.Add('21/02/2026 18:01:17 | LOG_USADO|C:\Users\kbeb\Documents\Colibri\sap_log_rts.txt');

    // 👇 ESSA é a linha que interessa: IdExec|IdRT|Linha|OK|Mensagem
    // troque 1126119 e 44 por IDs reais existentes no seu Access:
    SL.Add('21/02/2026 18:01:25 | 1126115|25|1|OK|Nota 326422165 gravada');

    SL.Add('21/02/2026 18:01:25 | FIM_LOTE|OK=1|ERRO=0');

    SL.SaveToFile(LogFake, TEncoding.ANSI);
  finally
    SL.Free;
  end;

  // Agora simula o “retorno SAP”
  SL := TStringList.Create;
  try
    SL.LoadFromFile(LogFake, TEncoding.ANSI);

    for i := 0 to SL.Count - 1 do
      ProcessarRetornoSAP(SL[i]);

  finally
    SL.Free;
  end;

  ShowMessage('Teste concluído. Verifique tblProgramacaoRT / tblProgramacaoExecutante.');
end;

function TFrmConsultaExecutantesProgramados.NormalizarPlataforma(const Valor: string): string;
var
  Prefixo: string;
  Numero: Integer;
begin
  Prefixo := Copy(Valor, 1, Pos('-', Valor)); // 'PCM-'
  Numero  := StrToIntDef(Copy(Valor, Pos('-', Valor) + 1, MaxInt), -1);

  if Numero >= 0 then
    Result := Prefixo + IntToStr(Numero)
  else
    Result := Valor;
end;


procedure TFrmConsultaExecutantesProgramados.actCancelarRTExecute(Sender: TObject);
var
  EnderecoVbs, EnderecoLog: string;
  DS: TDataSet;
  Bmk: TBookmark;
  LogRetorno: TStringList;
  i: Integer;

  idProgramacaoRT: Integer;
  RTNumero, Origem, Destino, CodigoSAP: string;
  DataInicio: TDateTime;

  function VbsQuote(const S: string): string;
  begin
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"';
  end;

  procedure MemoAdd(const S: string);
  begin
    MemoSAP.Lines.Add(S);
  end;

begin
  EnderecoVbs := ExtractFilePath(ParamStr(0)) + 'RT_Cancelar.vbs';
  EnderecoLog := ExtractFilePath(ParamStr(0)) + 'sap_log_cancel.txt';

  MemoSAP.Lines.Clear;

  DS := FrmDataModule.DataSourceGestaoRT.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    ShowMessage('Não há registros para cancelar.');
    Exit;
  end;

  // ======= VBS: cabeçalho + log robusto =======
  MemoAdd('Option Explicit');
  MemoAdd('Dim SapGuiAuto, application, connection, session');
  MemoAdd('Dim LOGFILE');
  MemoAdd('LOGFILE = ' + VbsQuote(EnderecoLog));
  MemoAdd('');

  MemoAdd('Sub LogLine(ByVal line)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim fso, folderPath, f');
  MemoAdd('  Set fso = CreateObject("Scripting.FileSystemObject")');
  MemoAdd('  folderPath = fso.GetParentFolderName(LOGFILE)');
  MemoAdd('  If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath');
  MemoAdd('  Set f = fso.OpenTextFile(LOGFILE, 8, True)');
  MemoAdd('  f.WriteLine Now & " | " & line');
  MemoAdd('  f.Close');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  MemoAdd('Function StatusText()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  StatusText = session.findById("wnd[0]/sbar").Text');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function StatusType()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  StatusType = session.findById("wnd[0]/sbar").MessageType');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Sub FechaPopupSeExistir()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  If session.Children.Count > 1 Then session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  MemoAdd('Sub CancelarRT(ByVal idProgRT, ByVal rtNumero)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim st, tp');
  MemoAdd('  LogLine idProgRT & "|CANCELAR|INICIO|RT=" & rtNumero');
  MemoAdd('');
  MemoAdd('  session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 250');
  MemoAdd('');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 250');
  MemoAdd('');
  MemoAdd('  ' + ''' >>> ajuste aqui se seu caminho de menu for diferente <<<');
  MemoAdd('  session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[11]/menu[0]").select');
  MemoAdd('  WScript.Sleep 200');
  MemoAdd('  session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('  WScript.Sleep 400');
  MemoAdd('');
  MemoAdd('  st = StatusText()');
  MemoAdd('  If Len(Trim(st)) = 0 Then');
  MemoAdd('    WScript.Sleep 600');
  MemoAdd('    st = StatusText()');
  MemoAdd('  End If');
  MemoAdd('  tp = StatusType()');
  MemoAdd('');
  MemoAdd('  If tp = "E" Then');
  MemoAdd('    LogLine idProgRT & "|CANCELAR|ERRO|" & st');
  MemoAdd('    FechaPopupSeExistir');
  MemoAdd('  Else');
  MemoAdd('    LogLine idProgRT & "|CANCELAR|OK|" & st');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  // ======= conecta no SAP só 1 vez =======
  MemoAdd('LogLine "INICIO_LOTE"');
  MemoAdd('If Not IsObject(application) Then');
  MemoAdd('  Set SapGuiAuto  = GetObject("SAPGUI")');
  MemoAdd('  Set application = SapGuiAuto.GetScriptingEngine');
  MemoAdd('End If');
  MemoAdd('If Not IsObject(connection) Then');
  MemoAdd('  Set connection = application.Children(0)');
  MemoAdd('End If');
  MemoAdd('If Not IsObject(session) Then');
  MemoAdd('  Set session = connection.Children(0)');
  MemoAdd('End If');
  MemoAdd('session.findById("wnd[0]").maximize');
  MemoAdd('');

  // ======= gera chamadas CancelarRT(...) conforme regras =======
  Bmk := DS.GetBookmark;
  try
    DS.DisableControls;
    try
      DS.First;
      while not DS.Eof do
      begin
        idProgramacaoRT := DS.FieldByName('idProgramacaoRT').AsInteger;
        RTNumero        := Trim(DS.FieldByName('RT').AsString);

        DataInicio := DS.FieldByName('DataInicio').AsDateTime;
        Origem     := DS.FieldByName('Origem').AsString;
        Destino    := DS.FieldByName('txtDestino').AsString; // atenção: aqui é txtDestino
        CodigoSAP  := DS.FieldByName('CodigoSAP').AsString;

        // regra 1: se RT está vazia → só marca erro local / loga e pula
        if RTNumero = '' then
        begin
          // você pode logar isso também no SAP log para rastreabilidade
          MemoAdd('LogLine "' + IntToStr(idProgramacaoRT) + '|CANCELAR|ERRO|RT_VAZIA"');
          DS.Next;
          Continue;
        end;

        // regra 2: se NÃO existe mais programação → cancela no SAP
        if not ExisteProgramacaoExecutante(DataInicio, Origem, Destino, CodigoSAP, RTNumero) then
        begin
          MemoAdd('CancelarRT ' + IntToStr(idProgramacaoRT) + ', ' + VbsQuote(RTNumero));
          MemoAdd('WScript.Sleep 250');
        end;

        DS.Next;
      end;
    finally
      DS.EnableControls;
    end;
  finally
    if DS.BookmarkValid(Bmk) then DS.GotoBookmark(Bmk);
    DS.FreeBookmark(Bmk);
  end;

  MemoAdd('LogLine "FIM_LOTE"');

  // salva e executa
  MemoSAP.Lines.SaveToFile(EnderecoVbs, TEncoding.ANSI);
  if FileExists(EnderecoLog) then DeleteFile(EnderecoLog);

  if ExecFileAndWait(EnderecoVbs, SW_SHOWNORMAL) then
  begin
    if FileExists(EnderecoLog) then
    begin
      LogRetorno := TStringList.Create;
      try
        LogRetorno.LoadFromFile(EnderecoLog);
        for i := 0 to LogRetorno.Count - 1 do
          ProcessarRetornoCancelamentoSAP(LogRetorno[i]); // <<< NOVO handler (abaixo)
      finally
        LogRetorno.Free;
      end;
      ShowMessage('Cancelamento concluído. Verifique log e tabela.');
    end
    else
      ShowMessage('Script terminou, mas não encontrei o log: ' + EnderecoLog);
  end
  else
    ShowMessage('Erro ao iniciar o script VBS.');
end;

procedure TFrmConsultaExecutantesProgramados.actConfigurarRTExecute(
  Sender: TObject);
begin
  if not Assigned(FrmConfigRT) then
    Application.CreateForm(TFrmConfigRT, FrmConfigRT);
  FrmConfigRT.Show;   // ou ShowModal conforme o caso
end;

procedure TFrmConsultaExecutantesProgramados.GerarMultiplasRTsArray;
type
  TRTItem = record
    IdExec: string;       // idProgramacaoExecutante
    IdRT: string;         // idProgramacaoRT (tblProgramacaoRT) -> CRIADO antes do VBS
    Linha: string;        // sequencial
    Acao: string;         // "CRIAR" / "MODIFICAR"
    RtNumero: string;     // quando MODIFICAR
    TipoRT: string;       // quando CRIAR (ex: "R3")
    Pernr: string;        // matrícula
    ContratoFix: string;  // "" ou "4600..."
    QmTxt: string;
    Requis: string;
    Pessoa: string;
    Fone: string;
    Iwerk: string;
    Ingrp: string;
    CodTrasg: string;
    Classe: string;       // TT/TR
    Origem: string;
    Destino: string;
    DtViagem: string;     // dd.mm.yyyy
    HViagem: string;      // hh:mm (IDA)
    HoraVolta: string;    // hh:mm (VOLTA)  <<< NOVO
    CCusto: string;
    Cobertura: string;    // "1"/"0"
    BateVolta: string;    // "1"/"0"       <<< NOVO
  end;

var
  EnderecoVbs, EnderecoLog, OrigemPlataformas: string;
  DS, DSColibri: TDataSet;
  Bmk: TBookmark;
  LogRetorno: TStringList;
  Itens: array of TRTItem;
  ItemCount: Integer;

  Dados: TDadosRT;
  TesteRT: TRT;
  CodicoCusto: TCodicoCusto;
  i, seq: Integer;
  A: TRTItem;
  HoraAtual, HoraCobertura: TDateTime;

  function VbsQuote(const S: string): string;
  begin
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"';
  end;

  procedure MemoAdd(const S: string);
  begin
    MemoSAP.Lines.Add(S);
  end;

  function DataSAP(const AData: TDateTime): string;
  begin
    Result := FrmPrincipal.dataSAP(DateToStr(AData));
  end;

  function DetermineClasse(const AOrigem, ADestino, ListaOrigens: string): string;
  begin
    if Pos(AOrigem, ListaOrigens) = 0 then Result := 'TR' else Result := 'TT';
  end;

  function BoolTo01(const B: Boolean): string;
  begin
    if B then Result := '1' else Result := '0';
  end;

  procedure AddItem(const AItem: TRTItem);
  begin
    SetLength(Itens, ItemCount + 1);
    Itens[ItemCount] := AItem;
    Inc(ItemCount);
  end;

  function BuildItemLine(const A: TRTItem): string;
  begin
    // Array(idExec,idRT,linha,acao,rtNumero,tipoRT,pernr,contratoFix,qmtxt,requis,pessoa,fone,
    //       iwerk,ingrp,codTrasg,classe,origem,destino,dt,horaIda,horaVolta,ccusto,cobertura,bateVolta)
    Result :=
      '  Array(' +
        VbsQuote(A.IdExec) + ',' +
        VbsQuote(A.IdRT) + ',' +
        VbsQuote(A.Linha) + ',' +
        VbsQuote(A.Acao) + ',' +
        VbsQuote(A.RtNumero) + ',' +
        VbsQuote(A.TipoRT) + ',' +
        VbsQuote(A.Pernr) + ',' +
        VbsQuote(A.ContratoFix) + ',' +
        VbsQuote(A.QmTxt) + ',' +
        VbsQuote(A.Requis) + ',' +
        VbsQuote(A.Pessoa) + ',' +
        VbsQuote(A.Fone) + ',' +
        VbsQuote(A.Iwerk) + ',' +
        VbsQuote(A.Ingrp) + ',' +
        VbsQuote(A.CodTrasg) + ',' +
        VbsQuote(A.Classe) + ',' +
        VbsQuote(A.Origem) + ',' +
        VbsQuote(A.Destino) + ',' +
        VbsQuote(A.DtViagem) + ',' +
        VbsQuote(A.HViagem) + ',' +
        VbsQuote(A.HoraVolta) + ',' +
        VbsQuote(A.CCusto) + ',' +
        VbsQuote(A.Cobertura) + ',' +
        VbsQuote(A.BateVolta) +
      ')';
  end;

  function CriarRegistroTabelaRT_Obrigatorio(const ADados: TDadosRT): Integer;
  begin
    Result := CriarRegistroTabelaRT(ADados); // deve retornar idProgramacaoRT
  end;

begin
  EnderecoVbs := ExtractFilePath(ParamStr(0)) + 'RT_R3.vbs';
  EnderecoLog := 'C:\Users\kbeb\Documents\Colibri\sap_log_rts.txt';
  OrigemPlataformas := FrmPrincipal.OrigemPlataformas;

  MemoSAP.Lines.Clear;

  FrmDataModule.ADOQueryColibri.Active := False;
  FrmDataModule.ADOQueryColibri.Active := True;

  DS := FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet;
  DSColibri := FrmDataModule.DataSourceColibri.DataSet;

  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    ShowMessage('Não há registros para emitir RT.');
    Exit;
  end;

  ItemCount := 0;
  SetLength(Itens, 0);
  seq := 0;

  // =========================
  // 1) MONTA ITENS + CRIA tblProgramacaoRT ANTES DO VBS
  // =========================
  Bmk := DS.GetBookmark;
  try
    DS.DisableControls;
    try
      DS.First;
      while not DS.Eof do
      begin
        // ---- fixos (colibri) ----
        Dados.TipoRT        := DSColibri.FieldByName('RT_Tipo').AsString;
        Dados.CentroPlan    := DSColibri.FieldByName('RT_CentroPlan').AsString;
        Dados.GrpPlan       := DSColibri.FieldByName('RT_GrpPlan').AsString;
        Dados.Requisitante  := DSColibri.FieldByName('RT_Requisitante').AsString;
        Dados.TelContato    := DSColibri.FieldByName('RT_TelContato').AsString;
        Dados.PessoaContato := DSColibri.FieldByName('RT_PessoaContato').AsString;
        Dados.Modal         := DSColibri.FieldByName('RT_Modal').AsString;

        Dados.HoraCobertura := DSColibri.FieldByName('RT_HoraCobertura').AsDateTime;
        HoraAtual     := Frac(Now);
        HoraCobertura := Frac(Dados.HoraCobertura);
        Dados.RTCobertura := (HoraAtual > HoraCobertura);

        Dados.HoraVolta := DSColibri.FieldByName('RT_HoraVolta').AsString; // <<< NOVO

        // ---- por passageiro ----
        Dados.idProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger;
        Dados.Origem            := NormalizarPlataforma(DS.FieldByName('Origem').AsString);
        Dados.Destino           := NormalizarPlataforma(DS.FieldByName('txtDestino').AsString);
        Dados.MatriculaPax      := TRIM(DS.FieldByName('CodigoSAP').AsString);
        Dados.Documento         := TRIM(DS.FieldByName('Documento').AsString);
        Dados.booleanBateVolta := DS.FieldByName('booleanBateVolta').AsBoolean;

        Dados.DataInicio := DataSAP(DS.FieldByName('DataProgramacao').AsDateTime);

        Dados.HoraInicio := DS.FieldByName('RT_HoraInicio').AsString;        // <<< IDA (vem do DS)
        if Trim(Dados.HoraVolta) = '' then Dados.HoraVolta := '16:30';
        if Trim(Dados.HoraInicio) = '' then Dados.HoraInicio := '07:00';

        Dados.DescricaoRT := Dados.Origem + ': ' + DS.FieldByName('NomeExecutante').AsString;

        CodicoCusto := CustoExecutante(Dados.MatriculaPax);
        Dados.CentroCusto := CodicoCusto.CentroCusto;

        if Dados.Origem <> Dados.Destino then
        begin
          Dados.Classe := DetermineClasse(Dados.Origem, Dados.Destino, OrigemPlataformas);

          TesteRT := Existe_RT(Dados);
          Inc(seq);

          FillChar(A, SizeOf(A), 0);
          A.IdExec      := IntToStr(Dados.idProgramacaoExecutante);
          A.IdRT        := '';
          A.Linha       := IntToStr(seq);
          A.TipoRT      := '';
          A.RtNumero    := '';
          A.ContratoFix := '';
          A.Pernr       := Dados.MatriculaPax;
          A.QmTxt       := Dados.DescricaoRT;
          A.Requis      := Dados.Requisitante;
          A.Pessoa      := Dados.PessoaContato;
          A.Fone        := Dados.TelContato;
          A.Iwerk       := Dados.CentroPlan;
          A.Ingrp       := Dados.GrpPlan;
          A.CodTrasg    := Dados.Modal;
          A.Classe      := Dados.Classe;
          A.Origem      := Dados.Origem;
          A.Destino     := Dados.Destino;
          A.DtViagem    := Dados.DataInicio;
          A.HViagem     := Dados.HoraInicio;     // IDA
          A.HoraVolta   := Dados.HoraVolta;      // VOLTA <<< NOVO
          A.CCusto      := Dados.CentroCusto;
          A.Cobertura   := BoolTo01(Dados.RTCobertura);
          A.BateVolta   := BoolTo01(Dados.booleanBateVolta); // <<< NOVO

          if TesteRT.booRTExiste and (Trim(TesteRT.RT_Numero) <> '') then
          begin
            // (se for usar MODIFICAR no futuro)
          end
          else
          begin
            A.Acao   := 'CRIAR';
            A.TipoRT := Dados.TipoRT;

            // cria registro-base tblProgramacaoRT com HoraInicio/HoraVolta/booleanBateVolta
            A.IdRT := IntToStr(CriarRegistroTabelaRT_Obrigatorio(Dados));
            AddItem(A);
          end;
        end;

        DS.Next;
      end;
    finally
      DS.EnableControls;
    end;
  finally
    if DS.BookmarkValid(Bmk) then DS.GotoBookmark(Bmk);
    DS.FreeBookmark(Bmk);
  end;

  if ItemCount = 0 then
  begin
    ShowMessage('Nenhuma RT a processar.');
    Exit;
  end;

  // =========================
  // 2) ESCREVE O VBS (motor + itens)
  // =========================
  MemoAdd('Option Explicit');
  MemoAdd('');
  MemoAdd('Dim SapGuiAuto, application, connection, session');
  MemoAdd('Dim LOGFILE');
  MemoAdd('');
  MemoAdd('LOGFILE = ' + VbsQuote(EnderecoLog));
  MemoAdd('');

  MemoAdd('Sub LogLine(ByVal line)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim fso, folderPath, f');
  MemoAdd('  Set fso = CreateObject("Scripting.FileSystemObject")');
  MemoAdd('  folderPath = fso.GetParentFolderName(LOGFILE)');
  MemoAdd('  If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath');
  MemoAdd('  Set f = fso.OpenTextFile(LOGFILE, 8, True)');
  MemoAdd('  f.WriteLine Now & " | " & line');
  MemoAdd('  f.Close');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  MemoAdd('Function GetStatusText(sess)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  GetStatusText = sess.findById("wnd[0]/sbar").Text');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function GetStatusType(sess)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  GetStatusType = sess.findById("wnd[0]/sbar").MessageType');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Sub FechaPopupSeExistir(sess)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  If sess.Children.Count > 1 Then');
  MemoAdd('    sess.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  MemoAdd('Function ResolverContratoObrigatorio(ByVal idExec, ByVal idRT, ByVal linha, ByVal contratoFix)');
  MemoAdd('  ResolverContratoObrigatorio = False');
  MemoAdd('  Dim pathContrato, st, stType');
  MemoAdd('  pathContrato = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                 "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CONTRATO[19,0]"');
  MemoAdd('  LogLine idExec & "|" & idRT & "|" & linha & "|INFO_CONTRATO_OBRIGATORIO|Campo Contrato obrigatório."');
  MemoAdd('  If Len(Trim(contratoFix)) > 0 Then');
  MemoAdd('    session.findById(pathContrato).text = contratoFix');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 300');
  MemoAdd('  Else');
  MemoAdd('    session.findById(pathContrato).setFocus');
  MemoAdd('    session.findById("wnd[0]").sendVKey 4');
  MemoAdd('    WScript.Sleep 300');
  MemoAdd('    session.findById("wnd[1]/usr/lbl[1,5]").setFocus');
  MemoAdd('    session.findById("wnd[1]").sendVKey 14');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('    WScript.Sleep 250');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 300');
  MemoAdd('  End If');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_CONTRATO|" & st');
  MemoAdd('    ResolverContratoObrigatorio = False');
  MemoAdd('  Else');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|OK_CONTRATO|" & st');
  MemoAdd('    ResolverContratoObrigatorio = True');
  MemoAdd('  End If');
  MemoAdd('End Function');
  MemoAdd('');

  // ========= PROCESSAR 1 RT (com bate-volta na LINHA 1 do grid) =========
  MemoAdd('Function ProcessarRT(ByVal idExec, ByVal idRT, ByVal linha, ByVal acao, ByVal rtNumero, ByVal tipoRT, _');
  MemoAdd('                     ByVal pernr, ByVal contratoFix, ByVal qmtxt, ByVal requis, ByVal pessoaContato, ByVal foneContato, _');
  MemoAdd('                     ByVal iwerk, ByVal ingrp, ByVal codTrasg, ByVal classePas, _');
  MemoAdd('                     ByVal origem, ByVal destino, ByVal dtViagem, ByVal hIda, ByVal hVolta, _');
  MemoAdd('                     ByVal ccusto, ByVal cobertura, ByVal bateVolta, ByVal salvar)');
  MemoAdd('  ProcessarRT = False');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim st, stType, pathPernr');
  MemoAdd('  LogLine idExec & "|" & idRT & "|" & linha & "|INICIO_ITEM|ACAO=" & acao & "|RT=" & rtNumero & "|PERNR=" & pernr');
  MemoAdd('');

  MemoAdd('  If UCase(acao) = "MODIFICAR" Then');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 250');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 250');
  MemoAdd('  Else');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw51"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 250');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = tipoRT');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 200');
  MemoAdd('  End If');
  MemoAdd('');

  MemoAdd('  session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/subNOTIF_TYPE:SAPLIQS0:1061/txtVIQMEL-QMTXT").text = qmtxt');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 150');
  MemoAdd('');

  MemoAdd('  If (Trim(cobertura) = "1") Or (UCase(Trim(cobertura)) = "TRUE") Then');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnANWENDERSTATUS").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True');
  MemoAdd('    session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    WScript.Sleep 200');
  MemoAdd('  End If');
  MemoAdd('');

  MemoAdd('  session.findById("wnd[0]/shellcont/shell").selectItem "010","Column01"');
  MemoAdd('  session.findById("wnd[0]/shellcont/shell").clickLink "010","Column01"');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtYSCSRTCB-REQUIS").text = requis');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-PESSOA_CONTATO").text = pessoaContato');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-TELEFONE_CONTATO").text = foneContato');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-IWERK").text = iwerk');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-INGRP").text = ingrp');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('');

  // PERNR (linha 0)
  MemoAdd('  pathPernr = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('            "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-PERNR[3,0]"');
  MemoAdd('  session.findById(pathPernr).text = pernr');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 300');
  MemoAdd('');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    If InStr(1, st, "não está ativo", vbTextCompare) > 0 Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_PAX|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('    If InStr(1, st, "Contrato obrigatório", vbTextCompare) > 0 Then');
  MemoAdd('      On Error GoTo 0');
  MemoAdd('      If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix) Then');
  MemoAdd('        FechaPopupSeExistir session');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      On Error Resume Next');
  MemoAdd('    Else');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_PERNR|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('');

  // LINHA 0 (IDA)
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-CODTRASG[8,0]").text = codTrasg');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CLASSEPAS[9,0]").text = classePas');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_ORIGEM[28,0]").text = origem');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_DESTINO[38,0]").text = destino');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-DTVIAGEM[52,0]").text = dtViagem');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-HVIAGEM[53,0]").text = hIda');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CCUSTO[63,0]").Text = ccusto');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('');

  // LINHA 1 (VOLTA) — só quando bateVolta=1
  MemoAdd('  If (Trim(bateVolta) = "1") Or (UCase(Trim(bateVolta)) = "TRUE") Then');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-PERNR[3,1]").text = pernr');

  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-CODTRASG[8,1]").text = codTrasg');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');

  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CLASSEPAS[9,1]").text = classePas');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');

  // origem/destino invertidos
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_ORIGEM[28,1]").text = destino');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');

  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_DESTINO[38,1]").text = origem');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');

  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-DTVIAGEM[52,1]").text = dtViagem');

  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-HVIAGEM[53,1]").text = hVolta');


  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CCUSTO[63,1]").Text = ccusto');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  End If');
  MemoAdd('');

  // salvar 2x
  MemoAdd('  If salvar Then');
  MemoAdd('    session.findById("wnd[0]").sendVKey 11');
  MemoAdd('    WScript.Sleep 500');
  MemoAdd('    session.findById("wnd[0]").sendVKey 11');
  MemoAdd('    WScript.Sleep 500');
  MemoAdd('  End If');
  MemoAdd('');

  // status final + log em formato compatível: idExec|idRT|linha|OK|mensagem
  MemoAdd('  WScript.Sleep 400');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  If Len(Trim(st)) = 0 Then');
  MemoAdd('    WScript.Sleep 600');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('  End If');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('    FechaPopupSeExistir session');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  LogLine idExec & "|" & idRT & "|" & linha & "|OK|" & st');
  MemoAdd('  ProcessarRT = True');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  // MAIN
  MemoAdd('LogLine "INICIO_LOTE"');
  MemoAdd('LogLine "LOG_USADO|" & LOGFILE');
  MemoAdd('');
  MemoAdd('If Not IsObject(application) Then');
  MemoAdd('  Set SapGuiAuto  = GetObject("SAPGUI")');
  MemoAdd('  Set application = SapGuiAuto.GetScriptingEngine');
  MemoAdd('End If');
  MemoAdd('If Not IsObject(connection) Then');
  MemoAdd('  Set connection = application.Children(0)');
  MemoAdd('End If');
  MemoAdd('If Not IsObject(session) Then');
  MemoAdd('  Set session = connection.Children(0)');
  MemoAdd('End If');
  MemoAdd('session.findById("wnd[0]").maximize');
  MemoAdd('');

  MemoAdd('Dim itens, i, okCount, errCount, salvar');
  MemoAdd('salvar = True');
  MemoAdd('');

  MemoAdd('itens = Array( _');
  for i := 0 to ItemCount - 1 do
  begin
    if i < ItemCount - 1 then
      MemoAdd(BuildItemLine(Itens[i]) + ', _')
    else
      MemoAdd(BuildItemLine(Itens[i]) + ')');
  end;

  MemoAdd('');
  MemoAdd('okCount = 0');
  MemoAdd('errCount = 0');
  MemoAdd('');
  MemoAdd('For i = 0 To UBound(itens)');
  MemoAdd('  Dim a, n');
  MemoAdd('  a = itens(i)');
  MemoAdd('  n = UBound(a)');
  MemoAdd('  If n < 23 Then'); // 0..23 = 24 campos
  MemoAdd('    LogLine "ERRO_ITEM_TAMANHO|" & i & "|UBOUND=" & n');
  MemoAdd('    errCount = errCount + 1');
  MemoAdd('  Else');
  MemoAdd('    If ProcessarRT(a(0),a(1),a(2),a(3),a(4),a(5),a(6),a(7),a(8),a(9), _');
  MemoAdd('                   a(10),a(11),a(12),a(13),a(14),a(15),a(16),a(17),a(18),a(19), _');
  MemoAdd('                   a(20),a(21),a(22),a(23), salvar) Then');
  MemoAdd('      okCount = okCount + 1');
  MemoAdd('    Else');
  MemoAdd('      errCount = errCount + 1');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('  WScript.Sleep 250');
  MemoAdd('Next');
  MemoAdd('');
  MemoAdd('LogLine "FIM_LOTE|OK=" & okCount & "|ERRO=" & errCount');

  // =========================
  // 3) SALVA, EXECUTA, LÊ LOG
  // =========================
  MemoSAP.Lines.SaveToFile(EnderecoVbs, TEncoding.ANSI);

  if FileExists(EnderecoLog) then
    DeleteFile(EnderecoLog);

  if ExecFileAndWait(EnderecoVbs, SW_SHOWNORMAL) then
  begin
    if FileExists(EnderecoLog) then
    begin
      LogRetorno := TStringList.Create;
      try
        LogRetorno.LoadFromFile(EnderecoLog, TEncoding.ANSI);
        for i := 0 to LogRetorno.Count - 1 do
          ProcessarRetornoSAP(LogRetorno[i]);
      finally
        LogRetorno.Free;
      end;
      ShowMessage('Processamento SAP concluído!');
    end
    else
      ShowMessage('Script terminou, mas não encontrei o log: ' + EnderecoLog);
  end
  else
    ShowMessage('Erro ao iniciar o script VBS.');
end;

function TFrmConsultaExecutantesProgramados.CriarRegistroTabelaRT(Dados: TDadosRT): Integer;
var
  Q: TADOQuery;
  QID: TADOQuery;
  NewID: Variant;
begin
  Q := FrmDataModule.ADOQueryAuxTabelaRT;

  // garante conexão
  if not Assigned(Q.Connection) then
    Q.Connection := FrmDataModule.ADOConnectionColibri;

  // garante dataset aberto (se for query/table vinculada à tblProgramacaoRT)
  if not Q.Active then
    Q.Open;

  Q.Insert;
  try
    Q.FieldByName('RT_Tipo').AsString         := Trim(Dados.TipoRT);
    Q.FieldByName('RT_Descricao').AsString    := Trim(Dados.DescricaoRT);
    Q.FieldByName('RT_Requisitante').AsString := Trim(Dados.Requisitante);
    Q.FieldByName('RT_PessoaContato').AsString:= Trim(Dados.PessoaContato);
    Q.FieldByName('RT_TelContato').AsString   := Trim(Dados.TelContato);
    Q.FieldByName('RT_CentroPlan').AsString   := Trim(Dados.CentroPlan);
    Q.FieldByName('RT_GrpPlan').AsString      := Trim(Dados.GrpPlan);

    Q.FieldByName('CodigoSAP').AsString       := Trim(Dados.MatriculaPax);
    Q.FieldByName('RT_Modal').AsString        := Trim(Dados.Modal);
    Q.FieldByName('RT_Classe').AsString       := Trim(Dados.Classe);
    Q.FieldByName('Origem').AsString          := Trim(Dados.Origem);
    Q.FieldByName('txtDestino').AsString      := Trim(Dados.Destino);

    Q.FieldByName('DataInicio').AsDateTime    := FrmPrincipal.DataSAP_ToDate(Dados.DataInicio);

    // ✅ ida/volta + flag
    Q.FieldByName('HoraInicio').AsString      := Trim(Dados.HoraInicio);
    Q.FieldByName('HoraVolta').AsString       := Trim(Dados.HoraVolta);
    Q.FieldByName('booleanBateVolta').AsBoolean := Dados.booleanBateVolta;

    Q.FieldByName('RT_CentroCusto').AsString  := Trim(Dados.CentroCusto);
    Q.FieldByName('RT_Cobertura').AsBoolean   := Dados.RTCobertura;

    Q.Post;

    // 1) tenta pegar do próprio dataset (às vezes já vem preenchido)
    try
      Result := Q.FieldByName('idProgramacaoRT').AsInteger;
    except
      Result := 0;
    end;

    // 2) se não veio, pega via @@IDENTITY na mesma conexão
    if Result <= 0 then
    begin
      QID := TADOQuery.Create(nil);
      try
        QID.Connection := FrmDataModule.ADOConnectionColibri;
        QID.SQL.Text := 'SELECT @@IDENTITY AS NewID';
        QID.Open;

        NewID := QID.FieldByName('NewID').Value;
        if not VarIsNull(NewID) then
          Result := Integer(NewID);
      finally
        QID.Free;
      end;
    end;

  except
    on E: Exception do
    begin
      if Q.State in dsEditModes then
        Q.Cancel;
      raise;
    end;
  end;
end;

function TFrmConsultaExecutantesProgramados.CustoExecutante(
  CodigoSAP: String): TCodicoCusto;
  var
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblExecutante.* FROM tblExecutante '+
  'WHERE CodigoSAP = "'+CodigoSAP+'";';
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Text:= SQLBase;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  Result.CentroCusto:=  FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('CentroCusto').AsString;
  Result.DiagramaRede:=  FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('DiagramaRede').AsString;
  Result.OperRede:=  FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('OperRede').AsString;
  Result.ElementoPEP:=  FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('ElementoPEP').AsString;
end;

procedure TFrmConsultaExecutantesProgramados.ProcessarRetornoSAP(const LinhaLog: string);
var
  IdExec, IdRT, Seq: Integer;
  StatusItem: string;
  Msg, RT_Numero, StatusRT, ErroRT: string;
begin
  ParseLogLine(LinhaLog, IdExec, IdRT, Seq, StatusItem, Msg);

  // ignora linhas que não são itens
  if (IdExec <= 0) and (IdRT <= 0) then
    Exit;

  // Aqui Msg é só "Nota 326... gravada"
  InterpretarMensagemSAP(Msg, RT_Numero, StatusRT, ErroRT);

  // Se quiser guardar o status do item do VBS junto:
  // Msg := StatusItem + '|' + Msg;

  if IdExec > 0 then
    AtualizarProgramacaoExecutanteComRetornoRT(IdExec, Msg, RT_Numero, StatusRT, ErroRT);

  if IdRT > 0 then
    AtualizarProgramacaoRTComRetornoRT(IdRT, Msg, RT_Numero, StatusRT, ErroRT);
end;


procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoExecutanteComRetornoRT(
  const idProgramacaoExecutante: Integer;
  const MensagemRT, RT_Numero, StatusRT, ErroRT: string);
var
  Q: TADOQuery;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante '+
      'SET RT_Mensagem = :RT_Mensagem, '+
      '    RT         = :RT, '+
      '    RT_Status  = :RT_Status, '+
      '    RT_Erro    = :RT_Erro '+
      'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemRT;
    Q.Parameters.ParamByName('RT').Value         := Trim(RT_Numero);
    Q.Parameters.ParamByName('RT_Status').Value  := StatusRT;
    Q.Parameters.ParamByName('RT_Erro').Value    := ErroRT;
    Q.Parameters.ParamByName('idProgramacaoExecutante').Value := idProgramacaoExecutante;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;


procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoRTComRetornoRT(
  const idProgramacaoRT: Integer;
  const MensagemRT, RT_Numero, StatusRT, ErroRT: string);
var
  Q: TADOQuery;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT '+
      'SET RT = :RT, RT_Mensagem = :RT_Mensagem, RT_Status = :RT_Status, RT_Erro = :RT_Erro '+
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.Parameters.ParamByName('RT').Value := RT_Numero;
    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemRT;
    Q.Parameters.ParamByName('RT_Status').Value := StatusRT;
    Q.Parameters.ParamByName('RT_Erro').Value := ErroRT;
    Q.Parameters.ParamByName('idProgramacaoRT').Value := idProgramacaoRT;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;


procedure TFrmConsultaExecutantesProgramados.actExcelExecute(Sender: TObject);
begin
  FrmPrincipal.GerarExcel(DBGridExecutantesProgramados,'Executantes Programados');
end;

procedure TFrmConsultaExecutantesProgramados.actExcelRelatorioExecute(
  Sender: TObject);
begin
  FrmPrincipal.ExcelStringGrid(RLImpressao,'Relatório de Impressão',
  '','',1);
end;

procedure TFrmConsultaExecutantesProgramados.actExecutadoExecute(
  Sender: TObject);
  var
    NumRegistros,idProgramacaoExecutante: Integer;
begin
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
  NumRegistros:= FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecordCount;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  FrmPrincipal.ProgressBarIncializa(NumRegistros,
  'Atribuindo Status de Execução...');
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('booleanSelecao').AsBoolean then
    begin
      idProgramacaoExecutante:= FrmDataModule.DataSourceConsultaExecutantesProgramados.
      DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
      FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,1,
      '','Executado','');
    end;
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  //==========================================================
  FrmPrincipal.ProgressBarAtualizar;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= true;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
end;

procedure TFrmConsultaExecutantesProgramados.actGerarMultiplasRTsExecute(
  Sender: TObject);
begin
  GerarMultiplasRTsArray;
  btnClearFiltro.Click;
end;

procedure TFrmConsultaExecutantesProgramados.actImprmirExecute(Sender: TObject);
  var
    SelRegistro,NumLinhas: Integer;
    strDataInicio,strDataFim,txt2,txt1: String;
begin
  StringGridImpressao.RowCount:= 1;
  StringGridImpressao.ColCount:= 4;
  StringGridImpressao.Cells[0,0]:= 'Nome';
  StringGridImpressao.Cells[1,0]:= 'Função';
  StringGridImpressao.Cells[2,0]:= 'Empresa';
  StringGridImpressao.Cells[3,0]:= 'Assinatura';
  SelRegistro:= FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    numLinhas:= StringGridImpressao.RowCount;
    StringGridImpressao.RowCount:= numLinhas+1;
    StringGridImpressao.Cells[0,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('NomeExecutante').AsString;
    StringGridImpressao.Cells[1,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('Funcao').AsString;
    StringGridImpressao.Cells[2,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('Empresa').AsString;
    StringGridImpressao.Cells[3,numLinhas]:= '';
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
  end;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo:= SelRegistro;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
  FrmPrincipal.AutoFitGrade(StringGridImpressao);
  StringGridImpressao.ColWidths[0]:= StringGridImpressao.ColWidths[0]+20;
  StringGridImpressao.ColWidths[1]:= StringGridImpressao.ColWidths[1]+20;
  StringGridImpressao.ColWidths[2]:= StringGridImpressao.ColWidths[2]+20;
  StringGridImpressao.ColWidths[3]:= StringGridImpressao.ColWidths[3]+20;
  //===================================================================
  if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;

  strDataInicio:= FrmPrincipal.corrigirData(dataInicio.DateTime);
  strDataFim:= FrmPrincipal.corrigirData(dataInicio.DateTime);
  if strDataInicio <> strDataFim then
    txt2:=  strDataInicio + ' à ' + strDataFim
  else
    txt2:= strDataInicio;

  txt1:= '';
  if InputQuery('Impressão Lista de Assinaturas','Entre com o titulo da lista:',txt1) then
  begin
    FrmPreview.ListaTituloRelatorio1:= TStringList.Create;
    FrmPreview.ListaTituloRelatorio1.Add(txt1);
    FrmPreview.ListaTituloRelatorio1.Add(txt2);
    FrmPreview.RadioGroupRelatorio.ItemIndex:= 2;
    FrmPreview.PrintPreview.PrintJobTitle:= txt1;
    FrmPreview.GeneratePages;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actExcluirProgramacaoRTExecute(Sender: TObject);
begin
  ExcluirProgramacaoRTSemRT;
  btnFiltroClearGestaoRT.Click;
end;

procedure TFrmConsultaExecutantesProgramados.ExcluirProgramacaoRTSemRT;
var
  Qry: TADOQuery;
begin
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FrmDataModule.ADOConnectionColibri;

    Qry.SQL.Text :=
      'DELETE FROM tblProgramacaoRT ' +
      'WHERE RT IS NULL OR Trim(RT) = ''''';

    Qry.ExecSQL;
  finally
    Qry.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actMotivoExecute(Sender: TObject);
  var
    MotivoProgramacao,MotivoMudanca,PalavraChaveProgramacao,PalavraChaveMudanca,
    StatusProgramacao: String;
    idProgramacaoExecutante: Integer;
begin
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    StatusProgramacao:=
    FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString;
    MotivoProgramacao:=
    FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('MotivoProgramacao').AsString;
    MotivoMudanca:=
    FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('MotivoMudanca').AsString;
    PalavraChaveProgramacao:=
    FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('PalavraChaveProgramacao').AsString;
    PalavraChaveMudanca:=
    FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('PalavraChaveMudanca').AsString;
    //=======================================================
    if (StatusProgramacao = 'Mudança') then
    begin
      idProgramacaoExecutante:= FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
      FieldByName('idProgramacaoExecutante').AsInteger;

      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= false;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:= idProgramacaoExecutante;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= true;

      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
      FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('MotivoProgramacao').AsString:= MotivoMudanca;
      FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('PalavraChaveProgramacao').AsString:= PalavraChaveMudanca;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
    end;
    //=======================================================
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actNaoExecutadoExecute(
  Sender: TObject);
begin
  if not FrmDataModule.ADOQueryConsultaExecutantesProgramados.IsEmpty then
  begin
    FrmPrincipal.PainelCancelamentoMudanca(1,'Motivo da Não Execução');
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actNumDiasExecute(Sender: TObject);
  var
    DataProgramacao1,DataProgramacao2: String;
    TotalDias: Integer;
begin
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.Sort := 'DataProgramacao ASC';

  DataProgramacao2:= 'aa';
  TotalDias:= 0;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    DataProgramacao1:= FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('DataProgramacao').AsString;
    if DataProgramacao1 <> DataProgramacao2 then
    begin
      TotalDias:= TotalDias+1;
    end;
    DataProgramacao2:= DataProgramacao1;
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
  end;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
  ShowMessage('Total de dias dos registros filtrados: '+IntToStr(TotalDias));
end;

procedure TFrmConsultaExecutantesProgramados.actProcurarGestaoRTExecute(
  Sender: TObject);
  var
    SQLString,SQLBase,
    DataProcuraIncio,DataProcuraFim: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  //=======================================================
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,false);
  //Query de procura
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:=
  'SELECT tblProgramacaoRT.* FROM tblProgramacaoRT '+
  'WHERE (DataInicio BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'+
  SQLString+' ORDER BY txtDestino,Origem;';
  FrmDataModule.ProcuraQueryCompleta(SQLBase,FrmDataModule.
  ADOQueryGestaoRT,StatusBarGestaoRT);
end;

procedure TFrmConsultaExecutantesProgramados.actProcurarProgramacaoExecutanteExecute(Sender: TObject);
  var
    SQLString,SQLBase,SQL_OrigemDestino,
    DataProcuraIncio,DataProcuraFim: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  //=======================================================
  SQL_OrigemDestino:= '';
  if CheckBoxOrigemDestino.Checked then
    SQL_OrigemDestino:= 'AND (tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem)';
  //=======================================================
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,false);
  //Query de procura
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:=
    {'SELECT tblRoteamento.*, tblAux_Rota_Distribuicao.*, tblProgramacaoExecutante.*, '+
    'tblProgramacaoDiaria.* FROM tblProgramacaoDiaria INNER JOIN (tblRoteamento INNER JOIN '+
    '(tblProgramacaoExecutante INNER JOIN tblAux_Rota_Distribuicao ON tblProgramacaoExecutante.'+
    'idProgramacaoExecutante = tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ON '+
    'tblRoteamento.idRoteamento = tblAux_Rota_Distribuicao.CodigoRota) ON tblProgramacaoDiaria.'+
    'idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+ }

    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE (DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'+
    SQLString+SQL_OrigemDestino+' ORDER BY DataProgramacao,txtDestino,Origem,txtTipoEtapaServico;';
  FrmDataModule.ProcuraQueryCompleta(SQLBase,FrmDataModule.
  ADOQueryConsultaExecutantesProgramados,StatusBarExecutantes);
end;

procedure TFrmConsultaExecutantesProgramados.actRelatorioCanceladoExecute(
  Sender: TObject);
begin
  RelatorioMotivo('Motivo Cancelamento',0);
end;

procedure TFrmConsultaExecutantesProgramados.actRelatorioMudancaExecute(
  Sender: TObject);
begin
  RelatorioMotivo('Motivo Mudança',1);
end;

procedure TFrmConsultaExecutantesProgramados.actRelatorioNaoExecutadaExecute(
  Sender: TObject);
begin
  RelatorioMotivo('Motivo Mudança',2);
end;

procedure TFrmConsultaExecutantesProgramados.actStatusSELECIONADOExecute(Sender: TObject);
  var
    idProgramacaoDiaria,NumCancelados,NumAprovados,NumExecutantes,SelRegistro: Integer;
begin
  idProgramacaoDiaria:= FrmDataModule.DataSourceConsultaExecutantesProgramados.
  DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
  NumCancelados:= FrmPrincipal.CalcNumCanceladosAprovado(idProgramacaoDiaria,0);
  NumAprovados:= FrmPrincipal.CalcNumCanceladosAprovado(idProgramacaoDiaria,1);
  NumExecutantes:= FrmPrincipal.CalcNumExecutantes(idProgramacaoDiaria);
  //========================================================
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:=idProgramacaoDiaria;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Edit;
  FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
  FieldByName('NumExecutantes').AsInteger:= NumExecutantes;
  FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
  FieldByName('NumCancelados').AsInteger:= NumCancelados;
  FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
  FieldByName('NumAprovados').AsInteger:= NumAprovados;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Post;
  //========================================================
  SelRegistro:= FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= true;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo:= SelRegistro;
end;

procedure TFrmConsultaExecutantesProgramados.actStatusTODOSExecute(
  Sender: TObject);
  var
    idProgramacaoDiaria,NumCancelados,NumAprovados,NumExecutantes: Integer;
begin
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.
  ADOQueryConsultaExecutantesProgramados.RecordCount,
  'Calculando "N° Exec.", "N° Apro." e "N° Canc."...');

  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    idProgramacaoDiaria:= FrmDataModule.DataSourceConsultaExecutantesProgramados.
    DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
    NumCancelados:= FrmPrincipal.CalcNumCanceladosAprovado(idProgramacaoDiaria,0);
    NumAprovados:= FrmPrincipal.CalcNumCanceladosAprovado(idProgramacaoDiaria,1);
    NumExecutantes:= FrmPrincipal.CalcNumExecutantes(idProgramacaoDiaria);
    //========================================================
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:=idProgramacaoDiaria;
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;

    FrmDataModule.ADOQueryConsultaProgramacao_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumExecutantes').AsInteger:= NumExecutantes;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumCancelados').AsInteger:= NumCancelados;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumAprovados').AsInteger:= NumAprovados;
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Post;
    //========================================================
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmConsultaExecutantesProgramados.actTransbordoExecute(
  Sender: TObject);
  var
    Origem: String;
begin
  Origem:= FrmPrincipal.OrigemPlataformas;
  btnClearFiltro.Click;
  FrmPrincipal.buscaFiledGrid1('InseridoProgramacaoTransporte','False','Exato',ColunasLayout,4,false);
  //FrmPrincipal.buscaFiledGrid1('Origem',Origem,'Diferente',ColunasLayout,4,false);
  FrmPrincipal.buscaFiledGrid1('StatusProgramacao','Aprovado','Exato',ColunasLayout,4,false);
  actProcurarProgramacaoExecutante.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.dataFimKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.dataInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.DBGridExecutantesProgramadosCellClick(
  Column: TColumn);
begin
  try
    if (Self.DBGridExecutantesProgramados.SelectedField.DataType = ftBoolean)AND
    (Column.Field.FieldName = 'booleanSelecao') then
    begin
      DBGridExecutantesProgramados.Options:=
      [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
      FrmDataModule.ADOqueryConsultaExecutantesProgramados.Edit;
      FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
      FieldByName('booleanSelecao').AsBoolean:= not Self.DBGridExecutantesProgramados.SelectedField.AsBoolean;
      FrmDataModule.ADOqueryConsultaExecutantesProgramados.Post;
    end
    else
      DBGridExecutantesProgramados.Options:=
      [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
  except

  end;
end;

procedure TFrmConsultaExecutantesProgramados.DBGridExecutantesProgramadosColumnMoved(
  Sender: TObject; FromIndex, ToIndex: Integer);
begin
  FrmPrincipal.CarregarRLColunasAtivasGRID(DBGridExecutantesProgramados);
  FrmPrincipal.GravarRLColunas('ExecutantesProgramados.txt');
end;

procedure TFrmConsultaExecutantesProgramados.DBGridExecutantesProgramadosDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridExecutantesProgramados,ColunasLayout,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'RequisitantePT') then
  begin
    if (FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('RequisitantePT').AsBoolean) then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 92);
    end
    else
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 93);
    end;
  end
  else if (Column.Field.FieldName = 'Kit_PS') then
  begin
    if (FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('Kit_PS').AsBoolean) then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 92);
    end
    else
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 93);
    end;
  end
  else if (Column.Field.FieldName = 'booleanSelecao') then
  begin
    Self.DBGridExecutantesProgramados.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridExecutantesProgramados.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end
  else if (Column.Field.FieldName = 'InseridoProgramacaoTransporte') then
  begin
    if (FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('InseridoProgramacaoTransporte').AsBoolean) then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 92);
    end
    else
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +15,Rect.Top + 1, 93);
    end;
  end
  else if (Column.Field.FieldName = 'StatusProgramacao') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Aprovado' then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clLime;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Cancelado' then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clRed;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Mudança' then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clYellow;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'StatusExecucao') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusExecucao').AsString = 'Executado' then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clLime;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('StatusExecucao').AsString = 'Não Executada' then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clRed;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumCancelados') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('NumCancelados').AsInteger > 0 then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clRed;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'NumAprovados') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
    FieldByName('NumAprovados').AsInteger > 0 then
    begin
      DBGridExecutantesProgramados.Canvas.Brush.Color:= clLime;
      DBGridExecutantesProgramados.Font.Color:= clBlack;
      DBGridExecutantesProgramados.Canvas.FillRect(Rect);
      DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.edtExecutanteKeyPress(
  Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

function TFrmConsultaExecutantesProgramados.ExisteProgramacaoExecutante(
  DataProcura: TDateTime; Origem, Destino, CodigoSAP, RTNumero: String): Boolean;
begin

  with FrmDataModule.ADOQueryTemporarioDBColibri do
  begin
    Close;
    SQL.Clear;
    SQL.Add(
      'SELECT TOP 1 tblProgramacaoExecutante.idProgramacaoExecutante '+
      'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
      'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE tblProgramacaoExecutante.Origem = :Origem '+
      '  AND tblProgramacaoDiaria.txtDestino = :Destino '+
      '  AND tblProgramacaoExecutante.CodigoSAP = :CodigoSAP '+
      '  AND tblProgramacaoDiaria.DataProgramacao = :DataInicio'
    );

    Parameters.ParamByName('Origem').Value    := Origem;
    Parameters.ParamByName('Destino').Value   := Destino;
    Parameters.ParamByName('CodigoSAP').Value := CodigoSAP;
    Parameters.ParamByName('DataInicio').DataType := ftDateTime;
    Parameters.ParamByName('DataInicio').Value := DataProcura;

    Open;
    Result := not IsEmpty;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ProcessarRetornoCancelamentoSAP(const LinhaLog: string);
var
  idProgRT: Integer;
  Acao, Status, Msg: string;
  p1, p2, p3: Integer;
begin
  // ignora linhas gerais
  if Pos('|CANCELAR|', LinhaLog) = 0 then
    Exit;

  // pega só a parte depois do " | "
  p1 := Pos(' | ', LinhaLog);
  if p1 <= 0 then Exit;
  Msg := Copy(LinhaLog, p1 + 3, MaxInt);

  // id|acao|status|msg
  p1 := Pos('|', Msg); if p1 <= 0 then Exit;
  p2 := PosEx('|', Msg, p1+1); if p2 <= 0 then Exit;
  p3 := PosEx('|', Msg, p2+1); if p3 <= 0 then Exit;

  idProgRT := StrToIntDef(Copy(Msg, 1, p1-1), 0);
  Acao     := Copy(Msg, p1+1, p2-p1-1);
  Status   := Copy(Msg, p2+1, p3-p2-1);
  Msg      := Trim(Copy(Msg, p3+1, MaxInt));

  if (idProgRT <= 0) or (Acao <> 'CANCELAR') then Exit;

  // sucesso → marca cancelada
  if SameText(Status, 'OK') then
    AtualizarProgramacaoRT_Cancelamento(idProgRT, True, Msg, 'CANCELADA', '')
  else
    AtualizarProgramacaoRT_Cancelamento(idProgRT, False, Msg, 'ERRO', 'CANCELAMENTO_FALHOU');
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoRT_Cancelamento(
  const idProgramacaoRT: Integer;
  const Cancelada: Boolean;
  const Mensagem, StatusRT, ErroRT: string);
var
  Q: TADOQuery;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT SET '+
      '  RT_Cancelada = :RT_Cancelada, '+
      '  RT_Mensagem  = :RT_Mensagem, '+
      '  RT_Status    = :RT_Status, '+
      '  RT_Erro      = :RT_Erro '+
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.Parameters.ParamByName('RT_Cancelada').Value := Cancelada;
    Q.Parameters.ParamByName('RT_Mensagem').Value  := Mensagem;
    Q.Parameters.ParamByName('RT_Status').Value    := StatusRT;
    Q.Parameters.ParamByName('RT_Erro').Value      := ErroRT;
    Q.Parameters.ParamByName('idProgramacaoRT').Value := idProgramacaoRT;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.Existe_RT(
  const Dados: TDadosRT
): TRT;
begin
  Result.booRTExiste := False;
  Result.RT_Numero := '';
  Result.idProgramacaoRT := 0;

  with FrmDataModule.ADOQueryTemporarioDBColibri do
  begin
    Close;
    SQL.Clear;
    SQL.Add(
      'SELECT TOP 1 '+
      '  idProgramacaoRT, RT, RT_Mensagem, RT_Status, RT_Erro '+
      'FROM tblProgramacaoRT '+
      'WHERE Origem     = :Origem '+
      '  AND txtDestino = :Destino '+
      '  AND CodigoSAP  = :CodigoSAP '+
      '  AND DataInicio = :DataInicio '+
      '  AND (RT_Status IS NULL OR RT_Status <> ''ERRO'')'
    );

    Parameters.ParamByName('Origem').Value     := Dados.Origem;
    Parameters.ParamByName('Destino').Value    := Dados.Destino;       // (vem do txtDestino do executante)
    Parameters.ParamByName('CodigoSAP').Value  := Dados.MatriculaPax;
    Parameters.ParamByName('DataInicio').Value := FrmPrincipal.DataSAP_ToDate(Dados.DataInicio);
    // se Dados.DataInicio já for DateTime, melhor ainda: := Dados.DataInicioDT;

    Open;

    if not IsEmpty then
    begin
      Result.RT_Numero       := FieldByName('RT').AsString;
      Result.idProgramacaoRT := FieldByName('idProgramacaoRT').AsInteger;
      Result.booRTExiste     := True;

      AtualizarProgramacaoExecutanteComRetornoRT(Dados.idProgramacaoExecutante,
        FieldByName('RT_Mensagem').AsString,
        FieldByName('RT').AsString,
        FieldByName('RT_Status').AsString,
        FieldByName('RT_Erro').AsString);
    end;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmConsultaExecutantesProgramados:=nil;
end;

procedure TFrmConsultaExecutantesProgramados.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  FrmDataModule.setFilterDBGrid(DBGridExecutantesProgramados);
  FrmDataModule.setFilterDBGrid(DBGridGestaoRT);
  enderecoColibriRegistro:= registroEndereco('Banco de dados');
  dataInicio.Date:= IncDay(now,1);
  dataFim.Date:= IncDay(now,1);
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  //=====================================
  FrmDataModule.ADOQueryProgramacaoCalendario1.Active:= true;

  if FrmPrincipal.logPerfil = 'Administrador' then
  begin
    actNaoExecutado.Enabled:= true;
    actExecutado.Enabled:= true;
    actStatusSELECIONADO.Enabled:= true;
    actStatusTODOS.Enabled:= true;
    DBGridExecutantesProgramados.ReadOnly:= false;
  end
  else if ((FrmPrincipal.logPerfil = 'Programação')OR
  (FrmPrincipal.logPerfil = 'Supervisor')) then
  begin
    actNaoExecutado.Enabled:= true;
    actExecutado.Enabled:= true;
    actStatusSELECIONADO.Enabled:= false;
    actStatusTODOS.Enabled:= false;
    DBGridExecutantesProgramados.ReadOnly:= true;
  end
  else
  begin
    actNaoExecutado.Enabled:= false;
    actExecutado.Enabled:= false;
    actStatusSELECIONADO.Enabled:= false;
    actStatusTODOS.Enabled:= false;
    DBGridExecutantesProgramados.ReadOnly:= true;
  end;
  //Incicialização
  FrmPrincipal.carregarRadioGroup(FrmDataModule.ADOQueryTipoEtapaServico_Carteira,
  RadioGroupTipoEtapaServico,'TipoEtapaServico','TODOS');
  RadioGroupTipoEtapaServico.ItemIndex:= 0;
  actProcurarProgramacaoExecutante.Execute;
  actProcurarGestaoRT.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmConsultaExecutantesProgramados.PageControl1Change(
  Sender: TObject);
begin
  case PageControl1.TabIndex of
    0: actProcurarProgramacaoExecutante.Execute;
    2: actProcurarGestaoRT.Execute;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.RelatorioMotivo(ColunaMotivo: String;
  TipoFonte:Integer);
  var
    SQLBase,SQLTipoEtapaServico,DataProcuraIncio,DataProcuraFim,SQLOrigem,SQL_OrigemDestino: String;
    ListaMotivo,ListaData,ListaPlataforma,ListaContar: TStringList;
    i,j,Linha,TotalContar: Integer;
begin
  if CheckBoxSomenteOrigem.Checked then
  begin
    SQLOrigem:= FrmPrincipal.OrigemPlataformas;
    if SQLOrigem <> '' then
      SQLOrigem:= 'AND'+FrmPrincipal.palavraBusca('','Origem','Exato',SQLOrigem,'OR');
  end
  else
    SQLOrigem:= '';
  //=========================================
  SQL_OrigemDestino:= '';
  if CheckBoxOrigemDestinoDiferente.Checked then
    SQL_OrigemDestino:= 'AND((tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem))'
  else
    SQL_OrigemDestino:= '';
  //=========================================
  SQLOrigem:= SQLOrigem+SQL_OrigemDestino;
  //====================================================================
  RLImpressao.RowCount:= 1;
  RLImpressao.ColCount:= 5;
  RLImpressao.Cells[0,0]:= 'Data Programação';
  RLImpressao.Cells[1,0]:= ColunaMotivo;
  RLImpressao.Cells[2,0]:= 'Instalação';
  RLImpressao.Cells[3,0]:= 'N° Exec.';
  RLImpressao.Cells[4,0]:= 'Serviço';
  if RadioGroupTipoEtapaServico.ItemIndex = 0 then
    SQLTipoEtapaServico:= ''
  else
  begin
    SQLTipoEtapaServico:= RadioGroupTipoEtapaServico.Items.Strings[RadioGroupTipoEtapaServico.ItemIndex];
    SQLTipoEtapaServico:= 'AND(txtTipoEtapaServico LIKE '+QuotedStr(SQLTipoEtapaServico)+')';
  end;
  //=====================================================
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  case TipoFonte of
    0: //Cancelamento
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
    '#)AND(MotivoProgramacao <> "")AND(StatusProgramacao LIKE "Cancelado")'+SQLTipoEtapaServico+SQLOrigem+
    ') ORDER BY MotivoProgramacao;';
    1: //Mudança
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
    '#)AND(MotivoProgramacao <> "")AND(StatusProgramacao LIKE "Mudança")'+SQLTipoEtapaServico+SQLOrigem+
    ') ORDER BY MotivoProgramacao;';
    2: //Não Executada
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
    '#)AND(MotivoNaoExecucao <> "")AND(StatusExecucao LIKE "Não Executada")'+SQLTipoEtapaServico+SQLOrigem+
    ') ORDER BY MotivoNaoExecucao;';
  end;
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  ListaMotivo:= TStringList.Create;
  ListaMotivo.Clear;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    case TipoFonte of
      0: ListaMotivo.Add(FrmDataModule.ADOQueryTemporarioDBColibri.
      FieldByName('MotivoProgramacao').AsString);
      1: ListaMotivo.Add(FrmDataModule.ADOQueryTemporarioDBColibri.
      FieldByName('MotivoProgramacao').AsString);
      2: ListaMotivo.Add(FrmDataModule.ADOQueryTemporarioDBColibri.
      FieldByName('MotivoNaoExecucao').AsString);
    end;
    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  end;
  FrmPrincipal.deleteRepetidosList(ListaMotivo);
  FrmPrincipal.ProgressBarIncializa(ListaMotivo.Count,'Relatório...');
  ListaData:= TStringList.Create;
  ListaData.Clear;
  ListaPlataforma:= TStringList.Create;
  ListaPlataforma.Clear;
  ListaContar:= TStringList.Create;
  ListaContar.Clear;
  for i := 0 to ListaMotivo.Count-1 do
  begin
    case TipoFonte of
      0: SQLBase:=
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
      'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
      'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
      '#)AND(MotivoProgramacao LIKE '+QuotedStr(ListaMotivo[i])+')AND(StatusProgramacao LIKE "Cancelado")'
      +SQLTipoEtapaServico+')'+SQLOrigem+' ORDER BY DataProgramacao,txtDestino;';
      1: SQLBase:=
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
      'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
      'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
      '#)AND(MotivoProgramacao LIKE '+QuotedStr(ListaMotivo[i])+')AND(StatusProgramacao LIKE "Mudança")'
      +SQLTipoEtapaServico+')'+SQLOrigem+' ORDER BY DataProgramacao,txtDestino;';
      2: SQLBase:=
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
      'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
      'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE ((DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+
      '#)AND(MotivoNaoExecucao LIKE '+QuotedStr(ListaMotivo[i])+')AND(StatusExecucao LIKE "Não Executada")'
      +SQLTipoEtapaServico+')'+SQLOrigem+' ORDER BY DataProgramacao,txtDestino;';
    end;
    FrmDataModule.ADOQueryTemporarioDBColibri.Close;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
    FrmDataModule.ADOQueryTemporarioDBColibri.Open;
    //========================================================
    ListaData.Clear;
    ListaPlataforma.Clear;
    ListaContar.Clear;
    while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
    begin
      ListaData.Add(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('DataProgramacao').AsString);
      ListaPlataforma.Add(FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtDestino').AsString);
      FrmDataModule.ADOQueryTemporarioDBColibri.Next;
    end;
    TotalContar:= 1;
    for j := ListaData.Count - 1 downto 1 do
    begin
      if ((UPPERCASE(ListaData[j-1])) = (UPPERCASE(ListaData[j])))AND
      ((UPPERCASE(ListaPlataforma[j-1])) = (UPPERCASE(ListaPlataforma[j]))) then
      begin
        TotalContar:= TotalContar+1;
        ListaData.Delete(j);
        ListaPlataforma.Delete(j);
      end
      else
      begin
        ListaContar.Insert(0,IntToStr(TotalContar));
        TotalContar:= 1;
      end;
    end;
    ListaContar.Insert(0,IntToStr(TotalContar));
    for j := 0 to ListaData.Count - 1 do
    begin
      Linha:= RLImpressao.RowCount;
      RLImpressao.RowCount:= RLImpressao.RowCount+1;
      RLImpressao.Cells[0,Linha]:= ListaData[j];
      RLImpressao.Cells[1,Linha]:= ListaMotivo[i];
      RLImpressao.Cells[2,Linha]:= ListaPlataforma[j];
      RLImpressao.Cells[3,Linha]:= ListaContar[j];
      RLImpressao.Cells[4,Linha]:= '';
    end;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  FrmPrincipal.AutoFitGrade(RLImpressao);
  FrmPrincipal.ProgressBarAtualizar;
  StatusBarRelatorio.Panels[0].Text:= 'N° registros: '+intToStr(RLImpressao.RowCount-1);
  FrmPrincipal.AutoFitStatusBar(StatusBarRelatorio);
  try
    RLImpressao.FixedRows:= 1;
  except
  end;
end;

procedure TFrmConsultaExecutantesProgramados.RLImpressaoFixedCellClick(
  Sender: TObject; ACol, ARow: Integer);
begin
  FrmPrincipal.clasifica(RLImpressao,ACol,true);
  FrmPrincipal.AutoFitGrade(RLImpressao);
end;

procedure TFrmConsultaExecutantesProgramados.WMMDIACTIVATE(
  var msg: TWMMDIACTIVATE);
 var
  active: TWinControl;
  idx: Integer;
begin
  active := FindControl(msg.ActiveWnd) ;

  if Assigned(active) then
  begin
    idx := FrmPrincipal.mdiChildrenTabs.Tabs.IndexOfObject(TObject(msg.ActiveWnd));
    FrmPrincipal.mdiChildrenTabs.Tag := -1;
    FrmPrincipal.mdiChildrenTabs.TabIndex := idx;
    FrmPrincipal.mdiChildrenTabs.Tag := 0;
  end;
end;

end.
