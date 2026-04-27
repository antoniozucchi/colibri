unit untAplatAnexos;

interface

uses
  Winapi.Windows, Winapi.Messages, Winapi.ShellAPI, System.SysUtils,
  System.Classes, System.Variants, System.StrUtils, System.NetEncoding,
  System.IOUtils, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.ExtCtrls,
  Vcl.Buttons, Vcl.Clipbrd, System.UITypes;

type
  TPTPdfItem = class
  private
    FArquivo: string;
    FCaminhoCompleto: string;
    FNumeroArquivo: string;
    FAno: string;
    FNumeroPesquisa: string;
  public
    constructor Create(const ACaminhoCompleto, ANumeroArquivo, AAno: string);
    property Arquivo: string read FArquivo;
    property CaminhoCompleto: string read FCaminhoCompleto;
    property NumeroArquivo: string read FNumeroArquivo;
    property Ano: string read FAno;
    property NumeroPesquisa: string read FNumeroPesquisa;
  end;

  TFrmAplatAnexos = class(TForm)
    PanelTitulo: TPanel;
    PanelConteudo: TPanel;
    MemoOrientacoes: TMemo;
    OpenDialogPdf: TOpenDialog;
    OpenDialogEdge: TOpenDialog;
    GridPanel1: TGridPanel;
    LabelUrl: TLabel;
    edtUrlAplat: TEdit;
    LabelEdge: TLabel;
    edtEdgePath: TEdit;
    LabelNumeroPT: TLabel;
    edtNumeroPT: TEdit;
    LabelPdf: TLabel;
    edtCaminhoPdf: TEdit;
    GridPanel2: TGridPanel;
    btnAbrirAplat: TButton;
    btnCarregarPT_PDF: TButton;
    btnAbrirPastaPdf: TButton;
    btnSalvarConfig: TButton;
    btnPrepararFluxo: TButton;
    ListBoxPT_PDF: TListBox;
    btnAnexarPT_APLAT: TButton;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure btnSelecionarPdfClick(Sender: TObject);
    procedure btnSelecionarEdgeClick(Sender: TObject);
    procedure btnSalvarConfigClick(Sender: TObject);
    procedure btnAbrirAplatClick(Sender: TObject);
    procedure btnAbrirPastaPdfClick(Sender: TObject);
    procedure btnPrepararFluxoClick(Sender: TObject);
    procedure btnAnexarPT_APLATClick(Sender: TObject);
    procedure btnCarregarPT_PDFClick(Sender: TObject);
  private
    procedure CarregarConfiguracoes;
    procedure SalvarConfiguracoes;
    procedure LimparListaPdf;
    procedure RegistrarResultado(const AMensagem: string);
    function ObterUrlAplat: string;
    function ObterExecutavelEdge: string;
    function ObterPastaPdf: string;
    function CarregarPdfsDaPasta: Integer;
    function TentarExtrairNumeroPT(const ANomeArquivo: string;
      out ANumeroPesquisa, ANumeroArquivo, AAno: string): Boolean;
    function ObterHostAplat: string;
    function ObterPastaPerfilEdgeAplat: string;
    function ObterScriptAutomacaoPath: string;
    function PortaDevToolsAtiva: Boolean;
    function AguardarPortaDevTools(const ATimeoutMs: Cardinal): Boolean;
    function ExecutarFluxoAnexoPorScript(const AItem: TPTPdfItem): Boolean;
    function ObterUrlEmissaoPT: string;
    function LocalizarJanelaAplat: HWND;
    function AtivarJanela(const AHandle: HWND): Boolean;
    procedure DigitarTextoUnicode(const ATexto: string);
    function EscapeJavaScriptString(const AValue: string): string;
    function ExecutarJavaScriptNoEdge(const AEdgeHandle: HWND;
      const AJavaScript: string; const ADelayMs: Cardinal = 900): Boolean;
    function ClicarXPathNoEdge(const AEdgeHandle: HWND; const AXPath: string;
      const ADelayMs: Cardinal = 900; const ADuploClique: Boolean = False): Boolean;
    function PreencherCampoXPathNoEdge(const AEdgeHandle: HWND; const AXPath,
      AValor: string; const ADelayMs: Cardinal = 700): Boolean;
    function IrParaUrlNoEdge(const AEdgeHandle: HWND; const AUrl: string): Boolean;
    function ClicarElementoAcessivel(const AEdgeHandle: HWND; const ATexto: string;
      ARole: Integer; const ATimeoutMs: Cardinal; const AParcial: Boolean = True;
      const ADuploClique: Boolean = False): Boolean;
    function PreencherCampoNumeroPT(const AEdgeHandle: HWND;
      const ANumeroPT: string): Boolean;
    function PreencherDialogoArquivo(const ACaminhoArquivo: string): Boolean;
    function FecharDetalhePT(const AEdgeHandle: HWND): Boolean;
    function ExecutarFluxoAnexo(const AEdgeHandle: HWND;
      const AItem: TPTPdfItem): Boolean;
  public
  end;

var
  FrmAplatAnexos: TFrmAplatAnexos;

implementation

uses
  Winapi.ActiveX, Winapi.oleacc, Winapi.WinSock, untPrincipal;

{$R *.dfm}

const
  REG_APLAT_URL = 'APLAT_URL';
  REG_APLAT_EDGE = 'APLAT_EDGE_EXE';
  REG_APLAT_PDF = 'APLAT_ULTIMO_PDF';
  REG_APLAT_PT = 'APLAT_ULTIMA_PT';
  EDGE_REMOTE_DEBUG_PORT = 9222;
  XPATH_EXIBIR_OPCOES =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/app-expando-filter/button';
  XPATH_CAMPO_NUMERO_PT =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/app-expando-filter/div/div[11]/div[1]/div/div/input';
  XPATH_BOTAO_PESQUISAR =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[2]/form/div/div[5]/div/button';
  XPATH_CARD_PT =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/div[3]/div';
  XPATH_ABA_ANEXOS =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[2]/div/app-cadastrar-pt/app-tabs/ul/li[13]/a';
  XPATH_BOTAO_ADICIONAR_ANEXO =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[2]/div/app-cadastrar-pt/app-tabs/app-tab[13]/div/app-anexos/div/div/div/app-fileuploader/label';
  XPATH_FECHAR_MODAL =
    '/html/body/app-root/div/app-permissao-trabalho/div/app-emissao-pt/app-dynamic-component/app-form-cadastrar-pt/app-modal/div/div/div/div[1]/button/span[1]';

type
  TAccessibleNode = record
    Acc: IAccessible;
    Child: OleVariant;
  end;

  TEdgeWindowSearch = record
    HostParte: string;
    Encontrado: HWND;
    Fallback: HWND;
  end;

function GetClasseJanela(const AWnd: HWND): string; forward;

function QuoteCmdArg(const AValue: string): string;
begin
  Result := '"' + StringReplace(AValue, '"', '\"', [rfReplaceAll]) + '"';
end;

function ConteudoScriptAutomacaoAplat: string;
const
  SCRIPT_BASE64 =
{$I codex_aplat_anexar_base64.inc}
    ;
begin
  Result := TNetEncoding.Base64.Decode(SCRIPT_BASE64);
end;

function CriarChildSelf: OleVariant;
begin
  Result := CHILDID_SELF;
end;

function NodeValido(const ANode: TAccessibleNode): Boolean;
begin
  Result := Assigned(ANode.Acc);
end;

function ValorInteiroVar(const AValor: OleVariant): Integer;
begin
  Result := 0;
  if VarIsOrdinal(AValor) then
    Result := AValor
  else if VarIsNumeric(AValor) then
    Result := AValor;
end;

function TextoDoNode(const AAcc: IAccessible; const AChild: OleVariant): string;
var
  Nome, Valor, Descricao: WideString;
begin
  Result := '';
  if not Assigned(AAcc) then
    Exit;

  Nome := '';
  Valor := '';
  Descricao := '';

  try
    AAcc.Get_accName(AChild, Nome);
  except
    Nome := '';
  end;

  try
    AAcc.Get_accValue(AChild, Valor);
  except
    Valor := '';
  end;

  try
    AAcc.Get_accDescription(AChild, Descricao);
  except
    Descricao := '';
  end;

  Result := Trim(Nome);
  if (Trim(Valor) <> '') and (not ContainsText(Result, Trim(Valor))) then
    Result := Trim(Result + ' ' + Trim(Valor));
  if (Trim(Descricao) <> '') and (not ContainsText(Result, Trim(Descricao))) then
    Result := Trim(Result + ' ' + Trim(Descricao));
end;

function RoleDoNode(const AAcc: IAccessible; const AChild: OleVariant): Integer;
var
  Papel: OleVariant;
begin
  Result := 0;
  if not Assigned(AAcc) then
    Exit;

  Papel := Unassigned;
  try
    if Succeeded(AAcc.Get_accRole(AChild, Papel)) then
      Result := ValorInteiroVar(Papel);
  except
    Result := 0;
  end;
end;

function NodeCombina(const ANode: TAccessibleNode; const ATexto: string;
  const ARole: Integer; const AParcial: Boolean): Boolean;
var
  TextoNode: string;
  RoleNode: Integer;
begin
  Result := False;
  if not NodeValido(ANode) then
    Exit;

  RoleNode := RoleDoNode(ANode.Acc, ANode.Child);
  if (ARole <> 0) and (RoleNode <> ARole) then
    Exit;

  if ATexto = '' then
    Exit(True);

  TextoNode := TextoDoNode(ANode.Acc, ANode.Child);
  if AParcial then
    Result := ContainsText(TextoNode, ATexto)
  else
    Result := SameText(Trim(TextoNode), Trim(ATexto));
end;

function EncontrarFilhoRecursivoPorClasseParcial(const AParent: HWND;
  const AClasseParte: string): HWND;
var
  Child, Encontrado: HWND;
  ClasseJanela: string;
begin
  Result := 0;
  Child := 0;
  repeat
    Child := FindWindowEx(AParent, Child, nil, nil);
    if Child = 0 then
      Break;

    ClasseJanela := GetClasseJanela(Child);
    if ContainsText(ClasseJanela, AClasseParte) then
      Exit(Child);

    Encontrado := EncontrarFilhoRecursivoPorClasseParcial(Child, AClasseParte);
    if Encontrado <> 0 then
      Exit(Encontrado);
  until False;
end;

function TentarObterRootAcessivelDaJanela(const AWnd: HWND;
  out ANode: TAccessibleNode): Boolean;
var
  Acc: IAccessible;
begin
  ANode.Acc := nil;
  ANode.Child := CriarChildSelf;
  Acc := nil;

  Result := (AWnd <> 0) and
    Succeeded(AccessibleObjectFromWindow(AWnd, OBJID_CLIENT, IID_IAccessible, Acc)) and
    Assigned(Acc);

  if not Result then
    Result := (AWnd <> 0) and
      Succeeded(AccessibleObjectFromWindow(AWnd, OBJID_WINDOW, IID_IAccessible, Acc)) and
      Assigned(Acc);

  if Result then
  begin
    ANode.Acc := Acc;
    ANode.Child := CriarChildSelf;
  end;
end;

function ObterRootAcessivel(const AWnd: HWND; out ANode: TAccessibleNode): Boolean;
var
  WndConteudo: HWND;
begin
  Result := False;
  if TentarObterRootAcessivelDaJanela(AWnd, ANode) then
    Exit(True);

  WndConteudo := EncontrarFilhoRecursivoPorClasseParcial(AWnd, 'Chrome_RenderWidgetHostHWND');
  if TentarObterRootAcessivelDaJanela(WndConteudo, ANode) then
    Exit(True);

  WndConteudo := EncontrarFilhoRecursivoPorClasseParcial(AWnd, 'Chrome_WidgetWin_0');
  if TentarObterRootAcessivelDaJanela(WndConteudo, ANode) then
    Exit(True);

  WndConteudo := EncontrarFilhoRecursivoPorClasseParcial(AWnd, 'Intermediate D3D Window');
  if TentarObterRootAcessivelDaJanela(WndConteudo, ANode) then
    Exit(True);
end;

function EncontrarNodeRecursivo(const ANode: TAccessibleNode; const ATexto: string;
  const ARole: Integer; const AParcial: Boolean; const AProfundidade: Integer;
  out AEncontrado: TAccessibleNode): Boolean;
var
  QuantidadeFilhos, I: Integer;
  FilhoNode: TAccessibleNode;
  ChildId: OleVariant;
  Disp: IDispatch;
begin
  Result := False;
  AEncontrado.Acc := nil;
  AEncontrado.Child := CriarChildSelf;

  if not NodeValido(ANode) then
    Exit;

  if NodeCombina(ANode, ATexto, ARole, AParcial) then
  begin
    AEncontrado := ANode;
    Exit(True);
  end;

  if AProfundidade <= 0 then
    Exit;

  QuantidadeFilhos := 0;
  try
    ANode.Acc.Get_accChildCount(QuantidadeFilhos);
  except
    QuantidadeFilhos := 0;
  end;

  if QuantidadeFilhos <= 0 then
    Exit;

  for I := 1 to QuantidadeFilhos do
  begin
    FilhoNode.Acc := nil;
    FilhoNode.Child := CriarChildSelf;
    ChildId := I;

    Disp := nil;
    try
      ANode.Acc.Get_accChild(ChildId, Disp);
    except
      Disp := nil;
    end;

    if Assigned(Disp) and Supports(Disp, IAccessible, FilhoNode.Acc) then
    begin
      FilhoNode.Child := CriarChildSelf;
    end
    else
    begin
      FilhoNode.Acc := ANode.Acc;
      FilhoNode.Child := ChildId;
    end;

    if EncontrarNodeRecursivo(FilhoNode, ATexto, ARole, AParcial,
      AProfundidade - 1, AEncontrado) then
      Exit(True);
  end;
end;

function TentarAtivarNode(const ANode: TAccessibleNode;
  const ADuploClique: Boolean): Boolean;
var
  Flags: Integer;
begin
  Result := False;
  if not NodeValido(ANode) then
    Exit;

  try
    Result := Succeeded(ANode.Acc.accDoDefaultAction(ANode.Child));
    if Result and ADuploClique then
    begin
      Sleep(250);
      ANode.Acc.accDoDefaultAction(ANode.Child);
    end;
  except
    Result := False;
  end;

  if Result then
    Exit;

  Flags := SELFLAG_TAKEFOCUS or SELFLAG_TAKESELECTION;
  try
    Result := Succeeded(ANode.Acc.accSelect(Flags, ANode.Child));
    if Result and ADuploClique then
    begin
      Sleep(250);
      ANode.Acc.accSelect(Flags, ANode.Child);
    end;
  except
    Result := False;
  end;
end;

function TentarDefinirValorNode(const ANode: TAccessibleNode;
  const AValor: string): Boolean;
begin
  Result := False;
  if not NodeValido(ANode) then
    Exit;

  try
    Result := Succeeded(ANode.Acc.Set_accValue(ANode.Child, AValor));
  except
    Result := False;
  end;
end;

procedure TeclaDown(const AVK: Word);
var
  Input: TInput;
begin
  ZeroMemory(@Input, SizeOf(Input));
  Input.Itype := INPUT_KEYBOARD;
  Input.ki.wVk := AVK;
  SendInput(1, Input, SizeOf(Input));
end;

procedure TeclaUp(const AVK: Word);
var
  Input: TInput;
begin
  ZeroMemory(@Input, SizeOf(Input));
  Input.Itype := INPUT_KEYBOARD;
  Input.ki.wVk := AVK;
  Input.ki.dwFlags := KEYEVENTF_KEYUP;
  SendInput(1, Input, SizeOf(Input));
end;

procedure PressionarTecla(const AVK: Word);
begin
  TeclaDown(AVK);
  Sleep(60);
  TeclaUp(AVK);
end;

procedure PressionarCtrlTecla(const AVK: Word);
begin
  TeclaDown(VK_CONTROL);
  Sleep(50);
  PressionarTecla(AVK);
  Sleep(50);
  TeclaUp(VK_CONTROL);
end;

procedure ColarTexto(const ATexto: string);
begin
  Clipboard.AsText := ATexto;
  Sleep(80);
  PressionarCtrlTecla(Ord('V'));
  Sleep(120);
end;

procedure SelecionarTudoECopiarTexto(const ATexto: string);
begin
  PressionarCtrlTecla(Ord('A'));
  Sleep(80);
  PressionarTecla(VK_DELETE);
  Sleep(80);
  ColarTexto(ATexto);
end;

procedure TFrmAplatAnexos.DigitarTextoUnicode(const ATexto: string);
var
  Inputs: array of TInput;
  I, J: Integer;
begin
  if ATexto = '' then
    Exit;

  SetLength(Inputs, Length(ATexto) * 2);
  J := 0;
  for I := 1 to Length(ATexto) do
  begin
    ZeroMemory(@Inputs[J], SizeOf(TInput));
    Inputs[J].Itype := INPUT_KEYBOARD;
    Inputs[J].ki.wVk := 0;
    Inputs[J].ki.wScan := Word(ATexto[I]);
    Inputs[J].ki.dwFlags := KEYEVENTF_UNICODE;
    Inc(J);

    ZeroMemory(@Inputs[J], SizeOf(TInput));
    Inputs[J].Itype := INPUT_KEYBOARD;
    Inputs[J].ki.wVk := 0;
    Inputs[J].ki.wScan := Word(ATexto[I]);
    Inputs[J].ki.dwFlags := KEYEVENTF_UNICODE or KEYEVENTF_KEYUP;
    Inc(J);
  end;

  if J > 0 then
    SendInput(J, Inputs[0], SizeOf(TInput));
end;

function TFrmAplatAnexos.EscapeJavaScriptString(const AValue: string): string;
begin
  Result := StringReplace(AValue, '\', '\\', [rfReplaceAll]);
  Result := StringReplace(Result, '''', '\x27', [rfReplaceAll]);
  Result := StringReplace(Result, #13, '\r', [rfReplaceAll]);
  Result := StringReplace(Result, #10, '\n', [rfReplaceAll]);
end;

function TFrmAplatAnexos.ExecutarJavaScriptNoEdge(const AEdgeHandle: HWND;
  const AJavaScript: string; const ADelayMs: Cardinal): Boolean;
var
  ScriptUrl: string;
begin
  Result := False;
  if not AtivarJanela(AEdgeHandle) then
    Exit;

  ScriptUrl := 'javascript:(function(){' + AJavaScript + '})();';
  PressionarCtrlTecla(Ord('L'));
  Sleep(250);
  DigitarTextoUnicode(ScriptUrl);
  Sleep(120);
  PressionarTecla(VK_RETURN);
  Sleep(ADelayMs);
  Result := True;
end;

function TFrmAplatAnexos.ClicarXPathNoEdge(const AEdgeHandle: HWND;
  const AXPath: string; const ADelayMs: Cardinal; const ADuploClique: Boolean): Boolean;
var
  JS: string;
begin
  JS :=
    'var e=document.evaluate(''' + EscapeJavaScriptString(AXPath) +
    ''',document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;' +
    'if(e){e.scrollIntoView({block:"center",inline:"center"});' +
    'var fire=function(t){e.dispatchEvent(new MouseEvent(t,{bubbles:true,cancelable:true,view:window}));};' +
    'try{e.focus();}catch(_e){}' +
    'fire("mouseover");fire("mousedown");fire("mouseup");fire("click");' +
    'try{e.click();}catch(_e){}';

  if ADuploClique then
    JS := JS + 'fire("mousedown");fire("mouseup");fire("click");fire("dblclick");';

  JS := JS + '}';
  Result := ExecutarJavaScriptNoEdge(AEdgeHandle, JS, ADelayMs);
end;

function TFrmAplatAnexos.PreencherCampoXPathNoEdge(const AEdgeHandle: HWND;
  const AXPath, AValor: string; const ADelayMs: Cardinal): Boolean;
var
  JS: string;
begin
  JS :=
    'var e=document.evaluate(''' + EscapeJavaScriptString(AXPath) +
    ''',document,null,XPathResult.FIRST_ORDERED_NODE_TYPE,null).singleNodeValue;' +
    'if(e){e.scrollIntoView({block:"center",inline:"center"});' +
    'try{e.focus();}catch(_e){}' +
    'e.value="";' +
    'e.dispatchEvent(new Event("input",{bubbles:true}));' +
    'e.value=''' + EscapeJavaScriptString(AValor) + ''';' +
    'e.dispatchEvent(new Event("input",{bubbles:true}));' +
    'e.dispatchEvent(new Event("change",{bubbles:true}));' +
    'try{e.blur();}catch(_e){}}';
  Result := ExecutarJavaScriptNoEdge(AEdgeHandle, JS, ADelayMs);
end;

function EncontrarJanelaTopLevelPorClasse(const AClasse: string): HWND;
begin
  Result := 0;
  repeat
    Result := FindWindowEx(0, Result, PChar(AClasse), nil);
    if (Result <> 0) and IsWindowVisible(Result) then
      Exit;
  until Result = 0;
end;

function GetCaptionJanela(const AWnd: HWND): string;
var
  Buffer: array[0..511] of Char;
begin
  Buffer[0] := #0;
  GetWindowText(AWnd, Buffer, Length(Buffer));
  Result := Buffer;
end;

function GetClasseJanela(const AWnd: HWND): string;
var
  Buffer: array[0..255] of Char;
begin
  Buffer[0] := #0;
  GetClassName(AWnd, Buffer, Length(Buffer));
  Result := Buffer;
end;

function EnumWindowsEdgeProc(AWnd: HWND; AParam: LPARAM): BOOL; stdcall;
var
  Info: ^TEdgeWindowSearch;
  ClasseJanela, CaptionJanela: string;
begin
  Result := True;
  Info := Pointer(AParam);
  if not IsWindowVisible(AWnd) then
    Exit;

  ClasseJanela := GetClasseJanela(AWnd);
  if not SameText(ClasseJanela, 'Chrome_WidgetWin_1') then
    Exit;

  CaptionJanela := GetCaptionJanela(AWnd);
  if (Info^.HostParte <> '') and
     (ContainsText(CaptionJanela, Info^.HostParte) or
      ContainsText(CaptionJanela, 'APLAT') or
      ContainsText(CaptionJanela, 'PETROBRAS')) then
  begin
    Info^.Encontrado := AWnd;
    Result := False;
    Exit;
  end;

  if Info^.Fallback = 0 then
    Info^.Fallback := AWnd;
end;

function EncontrarChildRecursivo(const AParent: HWND; const AClasse: string;
  const ACaptionContem: string = ''): HWND;
var
  Child, Encontrado: HWND;
  ClasseJanela, CaptionJanela: string;
begin
  Result := 0;
  Child := 0;
  repeat
    Child := FindWindowEx(AParent, Child, nil, nil);
    if Child = 0 then
      Break;

    ClasseJanela := GetClasseJanela(Child);
    CaptionJanela := GetCaptionJanela(Child);

    if SameText(ClasseJanela, AClasse) and
       ((ACaptionContem = '') or ContainsText(CaptionJanela, ACaptionContem)) then
      Exit(Child);

    Encontrado := EncontrarChildRecursivo(Child, AClasse, ACaptionContem);
    if Encontrado <> 0 then
      Exit(Encontrado);
  until False;
end;

function AguardarDialogoArquivo(const ATimeoutMs: Cardinal): HWND;
var
  Inicio: Cardinal;
  Wnd: HWND;
begin
  Result := 0;
  Inicio := GetTickCount;
  repeat
    Wnd := 0;
    repeat
      Wnd := FindWindowEx(0, Wnd, '#32770', nil);
      if (Wnd <> 0) and IsWindowVisible(Wnd) then
        Exit(Wnd);
    until Wnd = 0;

    Sleep(200);
  until GetTickCount - Inicio >= ATimeoutMs;
end;

constructor TPTPdfItem.Create(const ACaminhoCompleto, ANumeroArquivo,
  AAno: string);
begin
  inherited Create;
  FCaminhoCompleto := ACaminhoCompleto;
  FArquivo := ExtractFileName(ACaminhoCompleto);
  FNumeroArquivo := ANumeroArquivo;
  FAno := AAno;
  FNumeroPesquisa := ANumeroArquivo;
  while (Length(FNumeroPesquisa) > 1) and (FNumeroPesquisa[1] = '0') do
    Delete(FNumeroPesquisa, 1, 1);
  FNumeroPesquisa := FNumeroPesquisa + '/' + FAno;
end;

procedure TFrmAplatAnexos.btnAbrirAplatClick(Sender: TObject);
var
  UrlAplat: string;
  EdgePath: string;
  Parametros: string;
  CmdLine: string;
  PerfilDir: string;
  SI: TStartupInfo;
  PI: TProcessInformation;
begin
  UrlAplat := ObterUrlAplat;
  if UrlAplat = '' then
  begin
    MessageDlg('Informe a URL do APLAT antes de abrir o navegador.', mtWarning, [mbOK], 0);
    Exit;
  end;

  SalvarConfiguracoes;

  EdgePath := ObterExecutavelEdge;
  if EdgePath <> '' then
  begin
    PerfilDir := ObterPastaPerfilEdgeAplat;
    Parametros :=
      '--remote-debugging-port=' + IntToStr(EDGE_REMOTE_DEBUG_PORT) + ' ' +
      '--remote-allow-origins=* ' +
      '--user-data-dir=' + QuoteCmdArg(PerfilDir) + ' ' +
      '--no-first-run --no-default-browser-check ' +
      '--new-window ' + QuoteCmdArg(UrlAplat);
    CmdLine := QuoteCmdArg(EdgePath) + ' ' + Parametros;

    FillChar(SI, SizeOf(SI), 0);
    SI.cb := SizeOf(SI);
    SI.dwFlags := STARTF_USESHOWWINDOW;
    SI.wShowWindow := SW_SHOWNORMAL;
    FillChar(PI, SizeOf(PI), 0);

    if not CreateProcess(nil, PChar(CmdLine), nil, nil, False, 0, nil,
      PChar(ExtractFilePath(EdgePath)), SI, PI) then
    begin
      MessageDlg('Falha ao abrir o Edge: ' + SysErrorMessage(GetLastError),
        mtError, [mbOK], 0);
      Exit;
    end;

    CloseHandle(PI.hThread);
    CloseHandle(PI.hProcess);

    RegistrarResultado('Abrindo Edge com porta de depuração ' +
      IntToStr(EDGE_REMOTE_DEBUG_PORT) + '.');
    if AguardarPortaDevTools(15000) then
      RegistrarResultado('Porta de depuração do Edge ativa.')
    else
      RegistrarResultado(
        'Aviso: a porta de depuração do Edge ainda não respondeu. ' +
        'Se o Edge já estava aberto antes, feche-o e abra novamente por este botão.'
      );
  end
  else
  begin
    ShellExecute(Handle, 'open', PChar(UrlAplat), nil, nil, SW_SHOWNORMAL);
    MessageDlg(
      'O APLAT foi aberto no navegador padrão, mas a automação em lote precisa do caminho do Edge configurado para usar o DevTools Protocol.',
      mtWarning, [mbOK], 0);
  end;
end;

procedure TFrmAplatAnexos.btnAbrirPastaPdfClick(Sender: TObject);
var
  CaminhoPdf: string;
  Parametros: string;
begin
  CaminhoPdf := Trim(edtCaminhoPdf.Text);
  if CaminhoPdf = '' then
  begin
    MessageDlg('Informe a pasta dos PDFs antes de abrir.', mtWarning, [mbOK], 0);
    Exit;
  end;

  CaminhoPdf := ExpandFileName(CaminhoPdf);
  if DirectoryExists(CaminhoPdf) then
  begin
    ShellExecute(Handle, 'open', PChar(CaminhoPdf), nil, nil, SW_SHOWNORMAL);
    Exit;
  end;

  if FileExists(CaminhoPdf) then
  begin
    Parametros := '/select,"' + CaminhoPdf + '"';
    ShellExecute(Handle, 'open', 'explorer.exe', PChar(Parametros), nil, SW_SHOWNORMAL);
    Exit;
  end;

  MessageDlg('O caminho informado para os PDFs não foi encontrado.', mtError, [mbOK], 0);
end;

procedure TFrmAplatAnexos.btnAnexarPT_APLATClick(Sender: TObject);
var
  I: Integer;
  Item: TPTPdfItem;
begin
  SalvarConfiguracoes;

  if ListBoxPT_PDF.Count = 0 then
  begin
    if CarregarPdfsDaPasta <= 0 then
    begin
      MessageDlg('Carregue os PDFs da pasta antes de iniciar o anexo no APLAT.', mtWarning, [mbOK], 0);
      Exit;
    end;
  end;

  if ObterExecutavelEdge = '' then
  begin
    MessageDlg(
      'Configure o caminho do Edge e abra o APLAT pelo botão "Abrir APLAT". A automação em lote depende do Edge com porta de depuração habilitada.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  RegistrarResultado('Validando porta de depuração do Edge.');
  if not AguardarPortaDevTools(8000) then
  begin
    RegistrarResultado(
      'ERRO: a porta de depuração do Edge não respondeu na ' +
      IntToStr(EDGE_REMOTE_DEBUG_PORT) + '.'
    );
    MessageDlg(
      'O Edge não está com a porta de depuração ativa.' + sLineBreak +
      'Abra o APLAT pelo botão "Abrir APLAT", aguarde a janela carregar e faça o login nessa mesma janela dedicada.',
      mtError, [mbOK], 0);
    Exit;
  end;

  RegistrarResultado('Iniciando anexo em lote de ' + IntToStr(ListBoxPT_PDF.Count) + ' arquivo(s).');
  for I := 0 to ListBoxPT_PDF.Count - 1 do
  begin
    ListBoxPT_PDF.ItemIndex := I;
    Application.ProcessMessages;

    Item := TPTPdfItem(ListBoxPT_PDF.Items.Objects[I]);
    if Item = nil then
      Continue;

    RegistrarResultado('Processando PT ' + Item.NumeroPesquisa + ' -> ' + Item.Arquivo);
    if not ExecutarFluxoAnexoPorScript(Item) then
    begin
      MessageDlg(
        'Falha ao anexar a PT ' + Item.NumeroPesquisa + '.' + sLineBreak +
        'Veja o log na tela para o detalhe da etapa que falhou.' + sLineBreak +
        'Se a porta de depuração não estiver ativa, feche o Edge e reabra pelo botão "Abrir APLAT".',
        mtError, [mbOK], 0);
      Exit;
    end;
  end;

  RegistrarResultado('Lote concluído com sucesso.');
  MessageDlg('Anexo em lote finalizado.', mtInformation, [mbOK], 0);
end;

procedure TFrmAplatAnexos.btnCarregarPT_PDFClick(Sender: TObject);
var
  Quantidade: Integer;
begin
  SalvarConfiguracoes;
  Quantidade := CarregarPdfsDaPasta;
  if Quantidade > 0 then
    MessageDlg(IntToStr(Quantidade) + ' PDF(s) carregado(s) para anexo.', mtInformation, [mbOK], 0);
end;

procedure TFrmAplatAnexos.btnPrepararFluxoClick(Sender: TObject);
var
  Mensagem: TStringList;
begin
  SalvarConfiguracoes;

  Mensagem := TStringList.Create;
  try
    Mensagem.Add('Fluxo sugerido para anexar PTs no APLAT:');
    Mensagem.Add('');
    Mensagem.Add('1. Clique em "Abrir APLAT" e faça o login manual no Edge.');
    Mensagem.Add('2. Informe a pasta dos PDFs em edtCaminhoPdf.');
    Mensagem.Add('3. Clique em "Carregar PDFs" para montar a fila no ListBox.');
    Mensagem.Add('4. Revise a lista de arquivos PT.ANO.pdf carregados.');
    Mensagem.Add('5. Clique em "Anexar no APLAT" para executar o lote.');
    Mensagem.Add('6. Em caso de falha, reabra a tela de Emissão de PT no APLAT e repita a partir do item pendente.');
    MessageDlg(Mensagem.Text, mtInformation, [mbOK], 0);
  finally
    Mensagem.Free;
  end;
end;

procedure TFrmAplatAnexos.btnSalvarConfigClick(Sender: TObject);
begin
  SalvarConfiguracoes;
  MessageDlg('Configurações do APLAT salvas no perfil do usuário.', mtInformation, [mbOK], 0);
end;

procedure TFrmAplatAnexos.btnSelecionarEdgeClick(Sender: TObject);
begin
  if edtEdgePath.Text <> '' then
    OpenDialogEdge.FileName := edtEdgePath.Text;

  if OpenDialogEdge.Execute then
    edtEdgePath.Text := OpenDialogEdge.FileName;
end;

procedure TFrmAplatAnexos.btnSelecionarPdfClick(Sender: TObject);
begin
  if edtCaminhoPdf.Text <> '' then
    OpenDialogPdf.FileName := edtCaminhoPdf.Text;

  if OpenDialogPdf.Execute then
    edtCaminhoPdf.Text := OpenDialogPdf.FileName;
end;

function TFrmAplatAnexos.AtivarJanela(const AHandle: HWND): Boolean;
begin
  Result := IsWindow(AHandle);
  if not Result then
    Exit;

  if IsIconic(AHandle) then
    ShowWindow(AHandle, SW_RESTORE);

  Result := SetForegroundWindow(AHandle);
  if not Result then
    Result := BringWindowToTop(AHandle);

  Sleep(350);
end;

function TFrmAplatAnexos.CarregarPdfsDaPasta: Integer;
var
  Pasta, NumeroPesquisa, NumeroArquivo, Ano: string;
  SearchRec: TSearchRec;
  Item: TPTPdfItem;
begin
  Result := 0;
  Pasta := ObterPastaPdf;
  if Pasta = '' then
    Exit;

  LimparListaPdf;
  if FindFirst(IncludeTrailingPathDelimiter(Pasta) + '*.pdf', faAnyFile, SearchRec) = 0 then
  begin
    try
      repeat
        if ((SearchRec.Attr and faDirectory) = 0) and
           TentarExtrairNumeroPT(SearchRec.Name, NumeroPesquisa, NumeroArquivo, Ano) then
        begin
          Item := TPTPdfItem.Create(
            IncludeTrailingPathDelimiter(Pasta) + SearchRec.Name,
            NumeroArquivo,
            Ano
          );
          ListBoxPT_PDF.Items.AddObject(SearchRec.Name, Item);
          Inc(Result);
        end;
      until FindNext(SearchRec) <> 0;
    finally
      FindClose(SearchRec);
    end;
  end;

  if Result = 0 then
  begin
    MessageDlg(
      'Nenhum PDF no formato PT.ANO.pdf foi encontrado na pasta informada.',
      mtWarning, [mbOK], 0);
    Exit;
  end;

  if ListBoxPT_PDF.Count > 0 then
  begin
    ListBoxPT_PDF.ItemIndex := 0;
    Item := TPTPdfItem(ListBoxPT_PDF.Items.Objects[0]);
    if Item <> nil then
      edtNumeroPT.Text := Item.NumeroPesquisa;
  end;

  RegistrarResultado(IntToStr(Result) + ' arquivo(s) carregado(s) da pasta ' + Pasta + '.');
end;

procedure TFrmAplatAnexos.CarregarConfiguracoes;
begin
  edtUrlAplat.Text := Trim(FrmPrincipal.registroEndereco(REG_APLAT_URL));
  if edtUrlAplat.Text = '' then
    edtUrlAplat.Text := 'https://aplat.petrobras.com.br/#/login';

  edtEdgePath.Text := Trim(FrmPrincipal.registroEndereco(REG_APLAT_EDGE));
  if edtEdgePath.Text = '' then
    edtEdgePath.Text := 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe';

  edtCaminhoPdf.Text := Trim(FrmPrincipal.registroEndereco(REG_APLAT_PDF));
  edtNumeroPT.Text := Trim(FrmPrincipal.registroEndereco(REG_APLAT_PT));
end;

function TFrmAplatAnexos.ObterHostAplat: string;
begin
  Result := ObterUrlAplat;
  Result := StringReplace(Result, 'https://', '', [rfIgnoreCase]);
  Result := StringReplace(Result, 'http://', '', [rfIgnoreCase]);
  if Pos('/', Result) > 0 then
    Result := Copy(Result, 1, Pos('/', Result) - 1);
  if Pos('#', Result) > 0 then
    Result := Copy(Result, 1, Pos('#', Result) - 1);
end;

function TFrmAplatAnexos.ObterPastaPerfilEdgeAplat: string;
var
  BaseDir: string;
begin
  BaseDir := GetEnvironmentVariable('LOCALAPPDATA');
  if BaseDir = '' then
    BaseDir := ExtractFilePath(ParamStr(0));

  Result := TPath.Combine(BaseDir, 'Colibri\EdgeAplatProfile');
  ForceDirectories(Result);
end;

function TFrmAplatAnexos.PortaDevToolsAtiva: Boolean;
var
  WsaData: TWSAData;
  Sock: TSocket;
  Addr: TSockAddrIn;
begin
  Result := False;
  if WSAStartup($0202, WsaData) <> 0 then
    Exit;
  try
    Sock := socket(AF_INET, SOCK_STREAM, IPPROTO_TCP);
    if Sock = INVALID_SOCKET then
      Exit;
    try
      FillChar(Addr, SizeOf(Addr), 0);
      Addr.sin_family := AF_INET;
      Addr.sin_port := htons(EDGE_REMOTE_DEBUG_PORT);
      Addr.sin_addr.S_addr := inet_addr(PAnsiChar(AnsiString('127.0.0.1')));
      Result := Winapi.WinSock.connect(Sock, Addr, SizeOf(Addr)) <> SOCKET_ERROR;
    finally
      closesocket(Sock);
    end;
  finally
    WSACleanup;
  end;
end;

function TFrmAplatAnexos.AguardarPortaDevTools(
  const ATimeoutMs: Cardinal): Boolean;
var
  Inicio: Cardinal;
begin
  Inicio := GetTickCount;
  repeat
    if PortaDevToolsAtiva then
      Exit(True);
    Sleep(250);
    Application.ProcessMessages;
  until GetTickCount - Inicio >= ATimeoutMs;

  Result := PortaDevToolsAtiva;
end;

function TFrmAplatAnexos.ObterScriptAutomacaoPath: string;
var
  BaseExe: string;
  TempDir: string;
  Candidatos: array[0..5] of string;
  I: Integer;
  Conteudo: TStringList;
begin
  BaseExe := ExtractFilePath(ParamStr(0));
  TempDir := GetEnvironmentVariable('TEMP');
  if TempDir = '' then
    TempDir := BaseExe;

  Candidatos[0] := IncludeTrailingPathDelimiter(BaseExe) + 'codex_aplat_anexar.ps1';
  Candidatos[1] := ExpandFileName(BaseExe + '..\codex_aplat_anexar.ps1');
  Candidatos[2] := ExpandFileName(BaseExe + '..\..\codex_aplat_anexar.ps1');
  Candidatos[3] := ExpandFileName(BaseExe + '..\..\..\codex_aplat_anexar.ps1');
  Candidatos[4] := ExpandFileName(GetCurrentDir + '\codex_aplat_anexar.ps1');
  Candidatos[5] := IncludeTrailingPathDelimiter(TempDir) + 'codex_aplat_anexar_embutido.ps1';

  Result := '';
  for I := Low(Candidatos) to High(Candidatos) do
    if FileExists(Candidatos[I]) then
      Exit(Candidatos[I]);

  Conteudo := TStringList.Create;
  try
    Conteudo.Text := ConteudoScriptAutomacaoAplat;
    Conteudo.SaveToFile(Candidatos[5], TEncoding.UTF8);
    Result := Candidatos[5];
  finally
    Conteudo.Free;
  end;
end;

function TFrmAplatAnexos.ExecutarFluxoAnexoPorScript(const AItem: TPTPdfItem): Boolean;
var
  ScriptPath, TempDir, LogPath, CmdLine: string;
  SI: TStartupInfo;
  PI: TProcessInformation;
  WaitResult, ExitCode: DWORD;
  LinhasLog: TStringList;
  I: Integer;
begin
  Result := False;
  if (AItem = nil) or (Trim(AItem.CaminhoCompleto) = '') then
    Exit;

  if not PortaDevToolsAtiva then
  begin
    RegistrarResultado(
      'ERRO: a porta de depuração do Edge não está ativa na ' +
      IntToStr(EDGE_REMOTE_DEBUG_PORT) + '.'
    );
    Exit;
  end;

  ScriptPath := ObterScriptAutomacaoPath;
  if ScriptPath = '' then
  begin
    RegistrarResultado('ERRO: script codex_aplat_anexar.ps1 não encontrado.');
    Exit;
  end;

  TempDir := GetEnvironmentVariable('TEMP');
  if TempDir = '' then
    TempDir := ExtractFilePath(ParamStr(0));
  LogPath := IncludeTrailingPathDelimiter(TempDir) +
    Format('aplat_anexo_%d_%d.log', [GetCurrentProcessId, GetTickCount]);

  CmdLine :=
    'powershell.exe -NoProfile -ExecutionPolicy Bypass -File ' +
    QuoteCmdArg(ScriptPath) +
    ' -Port ' + IntToStr(EDGE_REMOTE_DEBUG_PORT) +
    ' -AplatHost ' + QuoteCmdArg(ObterHostAplat) +
    ' -EmissionUrl ' + QuoteCmdArg(ObterUrlEmissaoPT) +
    ' -NumeroPt ' + QuoteCmdArg(AItem.NumeroPesquisa) +
    ' -PdfPath ' + QuoteCmdArg(AItem.CaminhoCompleto) +
    ' -LogPath ' + QuoteCmdArg(LogPath);

  FillChar(SI, SizeOf(SI), 0);
  SI.cb := SizeOf(SI);
  SI.dwFlags := STARTF_USESHOWWINDOW;
  SI.wShowWindow := SW_HIDE;
  FillChar(PI, SizeOf(PI), 0);

  if not CreateProcess(nil, PChar(CmdLine), nil, nil, False, CREATE_NO_WINDOW,
    nil, PChar(ExtractFilePath(ScriptPath)), SI, PI) then
  begin
    RegistrarResultado('ERRO: falha ao iniciar automação PowerShell: ' +
      SysErrorMessage(GetLastError));
    Exit;
  end;

  try
    repeat
      WaitResult := WaitForSingleObject(PI.hProcess, 200);
      if WaitResult = WAIT_TIMEOUT then
        Application.ProcessMessages;
    until WaitResult <> WAIT_TIMEOUT;

    GetExitCodeProcess(PI.hProcess, ExitCode);
  finally
    CloseHandle(PI.hThread);
    CloseHandle(PI.hProcess);
  end;

  if FileExists(LogPath) then
  begin
    LinhasLog := TStringList.Create;
    try
      LinhasLog.LoadFromFile(LogPath, TEncoding.UTF8);
      for I := 0 to LinhasLog.Count - 1 do
        if Trim(LinhasLog[I]) <> '' then
          MemoOrientacoes.Lines.Add(LinhasLog[I]);
      MemoOrientacoes.SelStart := Length(MemoOrientacoes.Text);
      SendMessage(MemoOrientacoes.Handle, EM_SCROLLCARET, 0, 0);
    finally
      LinhasLog.Free;
      DeleteFile(LogPath);
    end;
  end;

  Result := ExitCode = 0;
  if not Result then
    RegistrarResultado('ERRO: automação PowerShell retornou código ' + IntToStr(ExitCode) + '.');
end;

function TFrmAplatAnexos.ClicarElementoAcessivel(const AEdgeHandle: HWND;
  const ATexto: string; ARole: Integer; const ATimeoutMs: Cardinal;
  const AParcial, ADuploClique: Boolean): Boolean;
var
  Inicio: Cardinal;
  RootNode, Encontrado: TAccessibleNode;
begin
  Result := False;
  Inicio := GetTickCount;

  repeat
    if not AtivarJanela(AEdgeHandle) then
      Exit(False);

    if ObterRootAcessivel(AEdgeHandle, RootNode) then
    begin
      if EncontrarNodeRecursivo(RootNode, ATexto, ARole, AParcial, 18, Encontrado) then
      begin
        if TentarAtivarNode(Encontrado, ADuploClique) then
          Exit(True);
      end;

      if (ARole <> 0) and
         EncontrarNodeRecursivo(RootNode, ATexto, 0, AParcial, 18, Encontrado) and
         TentarAtivarNode(Encontrado, ADuploClique) then
        Exit(True);
    end;

    Sleep(300);
  until GetTickCount - Inicio >= ATimeoutMs;
end;

function TFrmAplatAnexos.ExecutarFluxoAnexo(const AEdgeHandle: HWND;
  const AItem: TPTPdfItem): Boolean;
begin
  Result := False;
  if (AItem = nil) or (AItem.CaminhoCompleto = '') then
    Exit;

  if not IrParaUrlNoEdge(AEdgeHandle, ObterUrlEmissaoPT) then
  begin
    RegistrarResultado('Falha ao navegar para a tela Emissão de PT.');
    Exit;
  end;
  Sleep(1800);

  RegistrarResultado('Abrindo filtros da pesquisa.');
  ClicarXPathNoEdge(AEdgeHandle, XPATH_EXIBIR_OPCOES, 1200, False);
  Sleep(900);

  RegistrarResultado('Preenchendo PT ' + AItem.NumeroPesquisa + '.');
  if not PreencherCampoXPathNoEdge(AEdgeHandle, XPATH_CAMPO_NUMERO_PT, AItem.NumeroPesquisa, 900) then
    if not PreencherCampoNumeroPT(AEdgeHandle, AItem.NumeroPesquisa) then
    begin
      RegistrarResultado('Não foi possível preencher o número da PT.');
      Exit;
    end;
  Sleep(500);

  RegistrarResultado('Pesquisando PT.');
  if not ClicarXPathNoEdge(AEdgeHandle, XPATH_BOTAO_PESQUISAR, 2200, False) then
  begin
    AtivarJanela(AEdgeHandle);
    PressionarTecla(VK_RETURN);
    Sleep(2200);
  end;

  RegistrarResultado('Abrindo o resultado da PT.');
  if not ClicarXPathNoEdge(AEdgeHandle, XPATH_CARD_PT, 2500, True) then
    if not ClicarElementoAcessivel(AEdgeHandle, AItem.NumeroPesquisa, 0, 12000, True, True) then
    begin
      RegistrarResultado('Não encontrei o item da PT ' + AItem.NumeroPesquisa + '.');
      Exit;
    end;
  Sleep(2500);

  RegistrarResultado('Abrindo a aba Anexos.');
  if not ClicarXPathNoEdge(AEdgeHandle, XPATH_ABA_ANEXOS, 1200, False) then
    if not ClicarElementoAcessivel(AEdgeHandle, 'Anexos', 0, 10000, True) then
    begin
      RegistrarResultado('Não encontrei a aba "Anexos".');
      Exit;
    end;
  Sleep(1400);

  RegistrarResultado('Acionando "Adicionar Anexo".');
  if not ClicarXPathNoEdge(AEdgeHandle, XPATH_BOTAO_ADICIONAR_ANEXO, 1200, False) then
    if not ClicarElementoAcessivel(AEdgeHandle, 'Adicionar Anexo', 0, 10000, True) then
    begin
      RegistrarResultado('Não encontrei o botão "Adicionar Anexo".');
      Exit;
    end;

  if not PreencherDialogoArquivo(AItem.CaminhoCompleto) then
  begin
    RegistrarResultado('Falha ao preencher o diálogo de arquivo para ' + AItem.Arquivo + '.');
    Exit;
  end;

  Sleep(2500);
  FecharDetalhePT(AEdgeHandle);
  Sleep(800);
  Result := True;
end;

function TFrmAplatAnexos.FecharDetalhePT(const AEdgeHandle: HWND): Boolean;
begin
  Result := ClicarXPathNoEdge(AEdgeHandle, XPATH_FECHAR_MODAL, 900, False) or
            ClicarElementoAcessivel(AEdgeHandle, 'x', ROLE_SYSTEM_PUSHBUTTON, 3000, False) or
            ClicarElementoAcessivel(AEdgeHandle, 'x', 0, 3000, False);
  if not Result then
  begin
    AtivarJanela(AEdgeHandle);
    PressionarTecla(VK_ESCAPE);
    Result := True;
  end;
end;

procedure TFrmAplatAnexos.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
  FrmAplatAnexos := nil;
end;

procedure TFrmAplatAnexos.FormCreate(Sender: TObject);
begin
  FrmPrincipal.MDIChildCreated(Self.Handle);
  CarregarConfiguracoes;
end;

procedure TFrmAplatAnexos.FormDestroy(Sender: TObject);
begin
  LimparListaPdf;
  FrmPrincipal.MDIChildDestroyed(Self.Handle);
end;

function TFrmAplatAnexos.IrParaUrlNoEdge(const AEdgeHandle: HWND;
  const AUrl: string): Boolean;
begin
  Result := False;
  if not AtivarJanela(AEdgeHandle) then
    Exit;

  PressionarCtrlTecla(Ord('L'));
  Sleep(250);
  SelecionarTudoECopiarTexto(AUrl);
  Sleep(150);
  PressionarTecla(VK_RETURN);
  Sleep(3500);
  Result := True;
end;

procedure TFrmAplatAnexos.LimparListaPdf;
var
  I: Integer;
begin
  for I := 0 to ListBoxPT_PDF.Items.Count - 1 do
    ListBoxPT_PDF.Items.Objects[I].Free;
  ListBoxPT_PDF.Items.Clear;
end;

function TFrmAplatAnexos.LocalizarJanelaAplat: HWND;
var
  Info: TEdgeWindowSearch;
  Host: string;
begin
  FillChar(Info, SizeOf(Info), 0);
  Host := ObterUrlAplat;
  Host := StringReplace(Host, 'https://', '', [rfIgnoreCase]);
  Host := StringReplace(Host, 'http://', '', [rfIgnoreCase]);
  if Pos('/', Host) > 0 then
    Host := Copy(Host, 1, Pos('/', Host) - 1);
  if Pos('#', Host) > 0 then
    Host := Copy(Host, 1, Pos('#', Host) - 1);

  Info.HostParte := Host;
  EnumWindows(@EnumWindowsEdgeProc, LPARAM(@Info));

  if Info.Encontrado <> 0 then
    Exit(Info.Encontrado);

  if Info.Fallback <> 0 then
    Exit(Info.Fallback);

  Result := EncontrarJanelaTopLevelPorClasse('Chrome_WidgetWin_1');
end;

function TFrmAplatAnexos.ObterExecutavelEdge: string;
begin
  Result := Trim(edtEdgePath.Text);
  if (Result <> '') and not FileExists(Result) then
  begin
    MessageDlg(
      'O executável do Edge informado não foi encontrado. Ajuste o caminho ou deixe em branco para usar o navegador padrão.',
      mtWarning, [mbOK], 0);
    Result := '';
  end;
end;

function TFrmAplatAnexos.ObterPastaPdf: string;
begin
  Result := Trim(edtCaminhoPdf.Text);
  if Result = '' then
  begin
    MessageDlg('Informe a pasta onde estão os PDFs das PTs.', mtWarning, [mbOK], 0);
    Exit('');
  end;

  Result := ExpandFileName(Result);
  if FileExists(Result) then
    Result := ExtractFilePath(Result);

  if not DirectoryExists(Result) then
  begin
    MessageDlg('A pasta informada para os PDFs não foi encontrada.', mtError, [mbOK], 0);
    Result := '';
  end;
end;

function TFrmAplatAnexos.ObterUrlAplat: string;
begin
  Result := Trim(edtUrlAplat.Text);
end;

function TFrmAplatAnexos.ObterUrlEmissaoPT: string;
var
  UrlBase: string;
  PosHash: Integer;
begin
  UrlBase := ObterUrlAplat;
  PosHash := Pos('#', UrlBase);
  if PosHash > 0 then
    UrlBase := Copy(UrlBase, 1, PosHash - 1);
  Result := UrlBase + '#/permissaotrabalho/SMAR/emissaopt';
end;

function TFrmAplatAnexos.PreencherCampoNumeroPT(const AEdgeHandle: HWND;
  const ANumeroPT: string): Boolean;
const
  TentativasNome: array[0..4] of string = (
    'numeroPT',
    'inputUnidade',
    'Nº da PT',
    'Número da PT',
    'Numero da PT'
  );
var
  RootNode, NodeCampo: TAccessibleNode;
  I: Integer;
begin
  Result := False;
  if PreencherCampoXPathNoEdge(AEdgeHandle, XPATH_CAMPO_NUMERO_PT, ANumeroPT, 700) then
    Exit(True);

  if not AtivarJanela(AEdgeHandle) then
    Exit;

  if ObterRootAcessivel(AEdgeHandle, RootNode) then
  begin
    for I := Low(TentativasNome) to High(TentativasNome) do
    begin
      if EncontrarNodeRecursivo(RootNode, TentativasNome[I], ROLE_SYSTEM_TEXT, True, 18, NodeCampo) or
         EncontrarNodeRecursivo(RootNode, TentativasNome[I], 0, True, 18, NodeCampo) then
      begin
        if TentarAtivarNode(NodeCampo, False) then
        begin
          Sleep(150);
          if not TentarDefinirValorNode(NodeCampo, ANumeroPT) then
            SelecionarTudoECopiarTexto(ANumeroPT);
          Exit(True);
        end;
      end;
    end;

    if EncontrarNodeRecursivo(RootNode, '', ROLE_SYSTEM_TEXT, True, 18, NodeCampo) and
       TentarAtivarNode(NodeCampo, False) then
    begin
      Sleep(150);
      if not TentarDefinirValorNode(NodeCampo, ANumeroPT) then
        SelecionarTudoECopiarTexto(ANumeroPT);
      Exit(True);
    end;
  end;

  PressionarTecla(VK_TAB);
  Sleep(200);
  SelecionarTudoECopiarTexto(ANumeroPT);
  Result := True;
end;

function TFrmAplatAnexos.PreencherDialogoArquivo(
  const ACaminhoArquivo: string): Boolean;
var
  Dialogo, CampoArquivo, BotaoAbrir: HWND;
  Inicio: Cardinal;
begin
  Result := False;
  Dialogo := AguardarDialogoArquivo(10000);
  if Dialogo = 0 then
    Exit;

  AtivarJanela(Dialogo);
  CampoArquivo := EncontrarChildRecursivo(Dialogo, 'Edit');
  if CampoArquivo <> 0 then
  begin
    SendMessage(CampoArquivo, WM_SETTEXT, 0, LPARAM(PChar(ACaminhoArquivo)));
    Sleep(150);
  end
  else
  begin
    SelecionarTudoECopiarTexto(ACaminhoArquivo);
  end;

  BotaoAbrir := EncontrarChildRecursivo(Dialogo, 'Button', 'Abrir');
  if BotaoAbrir = 0 then
    BotaoAbrir := EncontrarChildRecursivo(Dialogo, 'Button', 'Open');

  if BotaoAbrir <> 0 then
    SendMessage(BotaoAbrir, BM_CLICK, 0, 0)
  else
    PostMessage(Dialogo, WM_COMMAND, IDOK, 0);

  Inicio := GetTickCount;
  repeat
    if not IsWindow(Dialogo) then
      Exit(True);
    Sleep(150);
  until GetTickCount - Inicio >= 8000;

  Result := not IsWindow(Dialogo);
end;

procedure TFrmAplatAnexos.RegistrarResultado(const AMensagem: string);
begin
  MemoOrientacoes.Lines.Add(FormatDateTime('hh:nn:ss', Now) + ' - ' + AMensagem);
  MemoOrientacoes.SelStart := Length(MemoOrientacoes.Text);
  SendMessage(MemoOrientacoes.Handle, EM_SCROLLCARET, 0, 0);
  Application.ProcessMessages;
end;

procedure TFrmAplatAnexos.SalvarConfiguracoes;
begin
  FrmPrincipal.registroEscrever(REG_APLAT_URL, Trim(edtUrlAplat.Text));
  FrmPrincipal.registroEscrever(REG_APLAT_EDGE, Trim(edtEdgePath.Text));
  FrmPrincipal.registroEscrever(REG_APLAT_PDF, Trim(edtCaminhoPdf.Text));
  FrmPrincipal.registroEscrever(REG_APLAT_PT, Trim(edtNumeroPT.Text));
end;

function TFrmAplatAnexos.TentarExtrairNumeroPT(const ANomeArquivo: string;
  out ANumeroPesquisa, ANumeroArquivo, AAno: string): Boolean;
var
  BaseNome: string;
  PosPonto: Integer;
  NumeroAux: Integer;
begin
  Result := False;
  ANumeroPesquisa := '';
  ANumeroArquivo := '';
  AAno := '';

  BaseNome := ChangeFileExt(ExtractFileName(ANomeArquivo), '');
  PosPonto := LastDelimiter('.', BaseNome);
  if PosPonto <= 1 then
    Exit;

  ANumeroArquivo := Copy(BaseNome, 1, PosPonto - 1);
  AAno := Copy(BaseNome, PosPonto + 1, MaxInt);
  if (ANumeroArquivo = '') or (AAno = '') then
    Exit;

  if not TryStrToInt(ANumeroArquivo, NumeroAux) then
    Exit(False);
  if not TryStrToInt(AAno, NumeroAux) then
    Exit(False);

  ANumeroPesquisa := ANumeroArquivo;
  while (Length(ANumeroPesquisa) > 1) and (ANumeroPesquisa[1] = '0') do
    Delete(ANumeroPesquisa, 1, 1);
  ANumeroPesquisa := ANumeroPesquisa + '/' + AAno;
  Result := True;
end;

end.

