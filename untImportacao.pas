unit untImportacao;

interface

uses
  Winapi.Windows, Winapi.Messages, SysUtils, System.Variants, System.Classes,
  Vcl.Graphics,Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ComCtrls, Vcl.ToolWin,
  Vcl.PlatformDefaultStyleActnCtrls, System.Actions, Vcl.ActnList, Vcl.ActnMan,
  Vcl.ExtCtrls, Data.DB, Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls,ComOBJ,
  Data.Win.ADODB, Vcl.Buttons,UITYPES, Vcl.DBCtrls,DateUtils,Math,ShellAPI,
  Vcl.Menus,ActiveX, Vcl.Mask, Vcl.TabNotBk, untDBGridFilter, uZucchi;

type
  TFrmImportacao = class(TForm)
    PageControlImportacao: TPageControl;
    TabSheet1: TTabSheet;
    ToolBar1: TToolBar;
    ActionManager1: TActionManager;
    Panel7: TPanel;
    actEnderecoExecutanteAPLAT: TAction;
    OpenDialog1: TOpenDialog;
    actImportarExecutanteAPLAT: TAction;
    actExecutanteAPLAT: TAction;
    Panel2: TPanel;
    actLocalizarAPLATExecutante: TAction;
    actLocalizarSAPExecutante: TAction;
    Edit4: TEdit;
    btnAbrirAPLAT: TSpeedButton;
    edtEnderecoExecutanteAPLAT: TEdit;
    actAnalisaExecutanteAPLAT: TAction;
    btnExecAPLAT: TBitBtn;
    btnNaoDefinidoAPLAT: TBitBtn;
    actExcluirAPLAT_Executantes: TAction;
    btnExcluirAPLAT: TBitBtn;
    actND_APLAT: TAction;
    btnAnaliseExecAPLAT: TBitBtn;
    PanelAjuda: TPanel;
    PanelTituloAjuda: TPanel;
    SpeedButton4: TSpeedButton;
    ToolBar4: TToolBar;
    StringGridND: TStringGrid;
    actExcluirLinha: TAction;
    actInserirExecutantes: TAction;
    ComboBoxEmpresa: TComboBox;
    ComboBoxTipoEtapaServico: TComboBox;
    ComboBoxFuncao: TComboBox;
    RadioGroupOrigem: TRadioGroup;
    BitBtn4: TBitBtn;
    Panel4: TPanel;
    StatusBarExecAPLAT: TStatusBar;
    DBGridExecutanteAPLAT: TFilterDBGrid;
    Splitter1: TSplitter;
    BitBtn11: TBitBtn;
    BitBtn13: TBitBtn;
    actInterromper: TAction;
    ColunasAPLAT: TStringGrid;
    actProcurar: TAction;
    actSubstituirMemoriaLocal: TAction;
    actCopiarMemoriaLocal: TAction;
    btnClearFiltroAPLAT: TToolButton;
    btnExcelAPLAT: TToolButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actEnderecoExecutanteAPLATExecute(Sender: TObject);
    procedure actImportarExecutanteAPLATExecute(Sender: TObject);
    procedure DBGridExecutanteAPLATDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actExecutanteAPLATExecute(Sender: TObject);
    procedure actLocalizarAPLATExecutanteExecute(Sender: TObject);
    procedure actAnalisaExecutanteAPLATExecute(Sender: TObject);
    procedure actExcluirAPLAT_ExecutantesExecute(Sender: TObject);
    procedure actND_APLATExecute(Sender: TObject);
    procedure SpeedButton4Click(Sender: TObject);
    procedure PanelTituloAjudaMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure StringGridNDSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ComboBoxTipoEtapaServicoCloseUp(Sender: TObject);
    procedure actExcluirLinhaExecute(Sender: TObject);
    procedure actInserirExecutantesExecute(Sender: TObject);
    procedure ComboBoxEmpresaCloseUp(Sender: TObject);
    procedure ComboBoxFuncaoCloseUp(Sender: TObject);
    procedure StringGridNDFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure actInterromperExecute(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
    procedure actSubstituirMemoriaLocalExecute(Sender: TObject);
    procedure actCopiarMemoriaLocalExecute(Sender: TObject);
    type
      TExecCadastro = record
        NomeExecutante,TipoEtapaServico,Funcao,Empresa,Documento: String;
      end;
  private
    { Private declarations }
    Interromper,tempoExcedido: Boolean;
    TabIndexPageControl: Integer;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure editEndereco(edt: TEdit);
    function AnalisaExecutanteAPLAT(CodigoSAP: String): TExecCadastro;
    function VerificaCodigoTipoEtapaServico(TipoEtapaServico: String): Integer;

  public
    { Public declarations }
  end;

var
  FrmImportacao: TFrmImportacao;

implementation
  uses untDataModule,untPrincipal;

{$R *.dfm}

procedure TFrmImportacao.actCopiarMemoriaLocalExecute(Sender: TObject);
  var
    Caminho_Copia,enderecoMemoria: String;
begin
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  enderecoMemoria := FrmPrincipal.Caption;
  enderecoMemoria:= ExtractFilePath(enderecoMemoria)+'\dbMemoria.mdb';
  Caminho_Copia := ExtractFilePath(Application.ExeName) + 'TEMPMemoria.mdb';
  DeleteFile(Caminho_Copia);
  CopyFile(PChar(enderecoMemoria), PChar(Caminho_Copia), false);
  FrmPrincipal.conectarBDDireto(Caminho_Copia,FrmDataModule.ADOConnectionMemoria);
end;

procedure TFrmImportacao.actEnderecoExecutanteAPLATExecute(Sender: TObject);
begin
  editEndereco(edtEnderecoExecutanteAPLAT);
  FrmPrincipal.registroEscrever('POB_APLAT', edtEnderecoExecutanteAPLAT.Text);
end;

procedure TFrmImportacao.actExcluirAPLAT_ExecutantesExecute(Sender: TObject);
begin
  if Application.MessageBox(PChar(
  'Deseja realmente excluir todos os registros de Executantes APLAT?'),
  '.::ATENÇÃO::.',36) = 6 then
  begin
    FrmPrincipal.deleteQueryRapido(FrmDataModule.ADOQueryImportarExecutanteAPLAT,
    'tblExecutanteAPLAT');
  end;
end;

procedure TFrmImportacao.actInserirExecutantesExecute(Sender: TObject);
  var
    i,codigoTipoEtapaServico: Integer;
    NomeExecutnate,CPF,OutroDocumento: String;
    EncontrouTodos: Boolean;
begin
  EncontrouTodos:= true;
  FrmDataModule.ADOQueryInserirExecutanteCadastro.Active:= false;
  FrmDataModule.ADOQueryInserirExecutanteCadastro.Active:= true;
  //Verificar preenchimento do Tipo de Etapa de Serviço
  for I := 1 to StringGridND.RowCount-1 do
  begin
    codigoTipoEtapaServico:= VerificaCodigoTipoEtapaServico(StringGridND.Cells[1,i]);
    if (codigoTipoEtapaServico = 0) then
    begin
      NomeExecutnate:= FrmPrincipal.TextoMaiusculo(StringGridND.Cells[0,i]);
      MessageBox(0,PChar('O Tipo de Etapa de Serviço inserido para o Executante: '+
      NomeExecutnate+' não é valido! Corrigir e tentar novamente!'),'Colibri', MB_ICONERROR);
      EncontrouTodos:= false;
      Break
    end;
  end;
  //Todos preenchido corretamente
  if EncontrouTodos = true then
  begin
    for I := 1 to StringGridND.RowCount-1 do
    begin
      //Verifica se é um CPF e corrige mascara de CPF
      CPF:= FrmPrincipal.VerificaCPF(StringGridND.Cells[5,i]);
      if CPF = '' then
        OutroDocumento:= StringGridND.Cells[6,i]
      else
        OutroDocumento:= '';
      //=====================================================================
      codigoTipoEtapaServico:= VerificaCodigoTipoEtapaServico(StringGridND.Cells[1,i]);
      NomeExecutnate:= FrmPrincipal.TextoMaiusculo(StringGridND.Cells[0,i]);
      FrmDataModule.ADOQueryInserirExecutanteCadastro.Insert;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('NomeExecutante').AsString:= NomeExecutnate;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('txtTipoEtapaServico').AsString:= StringGridND.Cells[1,i];
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('txtFuncao').AsString:= FrmPrincipal.TextoMaiusculo(StringGridND.Cells[2,i]);
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('txtEmpresa').AsString:= FrmPrincipal.TextoMaiusculo(StringGridND.Cells[3,i]);
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('CodigoSAP').AsString:= StringGridND.Cells[4,i];
      if StringGridND.ColCount = 7 then
      begin
        FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
        FieldByName('CPF').AsString:= CPF;
        FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
        FieldByName('OutroDocumento').AsString:= OutroDocumento;
      end;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('codigoTipoEtapaServico').AsInteger:= codigoTipoEtapaServico;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('AvaliadoPor').AsString:= FrmPrincipal.logChave;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('DataAtualizacao').AsDateTime:= now;
      FrmDataModule.DataSourceInserirExecutanteCadastro.DataSet.
      FieldByName('Computador').AsString:= FrmPrincipal.logMaquina;
      FrmDataModule.ADOQueryInserirExecutanteCadastro.Post;
    end;
    //Finalização
    PanelAjuda.Visible:= false;
    Splitter1.Visible:= false;
    case RadioGroupOrigem.ItemIndex of
      0:
      begin
        actAnalisaExecutanteAPLAT.Execute;
        btnClearFiltroAPLAT.Click;
      end;
    end;
    MessageBox(0,'Executantes inseridos com sucesso!','Colibri', MB_ICONINFORMATION);
  end;
end;

procedure TFrmImportacao.actInterromperExecute(Sender: TObject);
begin
  Interromper:= true;
end;

procedure TFrmImportacao.actExcluirLinhaExecute(Sender: TObject);
begin
  FrmPrincipal.DeleteRow(StringGridND,StringGridND.Row);
  AutoFitGrade(StringGridND);
end;

procedure TFrmImportacao.actLocalizarAPLATExecutanteExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasAPLAT,true);
  SQLBase:= 'SELECT tblExecutanteAPLAT.* FROM tblExecutanteAPLAT '+
  SQLString+' ORDER BY NomeExecutante;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryImportarExecutanteAPLAT,
  StatusBarExecAPLAT);
end;

procedure TFrmImportacao.actND_APLATExecute(Sender: TObject);
  var
    numLinhas: Integer;
    CodigoSAP,NomeExecutante: String;
begin
  RadioGroupOrigem.ItemIndex:= 0;
  FrmDataModule.ADOQueryExecutante.Active:= false;
  FrmDataModule.ADOQueryExecutante.Active:= true;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Close;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.SQL.Clear;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.SQL.Add(
  'SELECT tblExecutanteAPLAT.* '+
  'FROM tblExecutanteAPLAT '+
  'WHERE (txtTipoEtapaServico like '+QuotedStr('NÃO DEFINIDO')+');');
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Open;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Active:= true;
  if not FrmDataModule.ADOQueryImportarExecutanteAPLAT.IsEmpty then
  begin
    StringGridND.FixedRows:= 0;
    StringGridND.RowCount:= 1;
    StringGridND.ColCount:= 5;
    StringGridND.Cells[0,0]:= 'Executante';
    StringGridND.Cells[1,0]:= 'Tipo de Etapa de Serviço';
    StringGridND.Cells[2,0]:= 'Função';
    StringGridND.Cells[3,0]:= 'Empresa';
    StringGridND.Cells[4,0]:= 'Código SAP';
    FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'TipoEtapaServico',
    'SELECT tblTipoEtapaServico.* FROM tblTipoEtapaServico ORDER BY TipoEtapaServico;',
    ComboBoxTipoEtapaServico);
    FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'txtFuncao',
    'SELECT tblExecutante.txtFuncao FROM tblExecutante GROUP BY tblExecutante.txtFuncao;',
    ComboBoxFuncao);
    FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'txtEmpresa',
    'SELECT tblExecutante.txtEmpresa FROM tblExecutante GROUP BY tblExecutante.txtEmpresa;',
    ComboBoxEmpresa);

    FrmDataModule.ADOQueryImportarExecutanteAPLAT.First;
    while not FrmDataModule.ADOQueryImportarExecutanteAPLAT.Eof do
    begin
      CodigoSAP:= FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('CodigoSAP').AsString;
      NomeExecutante:= FrmPrincipal.TextoMaiusculo(FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('NomeExecutante').AsString);
      //Caso vazio, procurar se existe executante no
      //cadastro e atualiza o a planilha do APLAT
      if CodigoSAP = '' then
      begin
        FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
        FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
        FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(
        'SELECT tblExecutante.* FROM tblExecutante '+
        'WHERE (((NomeExecutante like '+QuotedStr('%'+NomeExecutante+'%')+')));');
        FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
        //====================================================
        if not FrmDataModule.ADOQueryTemporarioDBConsulta1.IsEmpty then
        begin
          FrmDataModule.ADOQueryImportarExecutanteAPLAT.Edit;
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('txtTipoEtapaServico').AsString:=
          FrmDataModule.DataSourceTemporarioDBConsulta1.DataSet.
          FieldByName('txtTipoEtapaServico').AsString;
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('txtFuncao').AsString:=
          FrmDataModule.DataSourceTemporarioDBConsulta1.DataSet.
          FieldByName('txtFuncao').AsString;
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('txtEmpresa').AsString:=
          FrmDataModule.DataSourceTemporarioDBConsulta1.DataSet.
          FieldByName('txtEmpresa').AsString;
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('CodigoSAP').AsString:=
          FrmDataModule.DataSourceTemporarioDBConsulta1.DataSet.
          FieldByName('CodigoSAP').AsString;
          FrmDataModule.ADOQueryImportarExecutanteAPLAT.Post;
          MessageBox(0,PChar(
          'O código SAP não definido na planilha de importação do APLAT foi corrigido!'+#13+
          'Necessário corrigir somente a base de dados do APLAT para futuras importações!'),
          'Colibri', MB_ICONINFORMATION);
        end;
      end
      else
      begin
        numLinhas:= StringGridND.RowCount;
        StringGridND.RowCount:= numLinhas+1;
        StringGridND.Cells[0,numLinhas]:= NomeExecutante;
        StringGridND.Cells[1,numLinhas]:= '';
        StringGridND.Cells[2,numLinhas]:= FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
        FieldByName('Funcao').AsString;
        StringGridND.Cells[3,numLinhas]:= FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
        FieldByName('Empresa').AsString;
        StringGridND.Cells[4,numLinhas]:= CodigoSAP;
      end;
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Next;
    end;
    if StringGridND.RowCount > 1 then
    begin
      try
        StringGridND.FixedRows:= 1;
      except
        StringGridND.FixedRows:= 0;
        StringGridND.RowCount:= 1;
      end;
      AutoFitGrade(StringGridND);
      StringGridND.ColWidths[2]:= 120;
      StringGridND.ColWidths[3]:= 120;
      PanelAjuda.Visible:= true;
      Splitter1.Visible:= true;
      Splitter1.Top:= 145;
    end;
  end
  else
  begin
    MessageBox(0,'Não existem Executantes "NÃO DEFINIDO"','Colibri', MB_ICONINFORMATION);
    btnClearFiltroAPLAT.Click;
  end;
end;

procedure TFrmImportacao.actProcurarExecute(Sender: TObject);
begin
  if PageControlImportacao.ActivePage.Caption = 'Executantes: Hotelaria [APLAT]' then
  begin
    actLocalizarAPLATExecutante.Execute;
    edtEnderecoExecutanteAPLAT.Text:= FrmPrincipal.registroEndereco('POB_APLAT');
    TabIndexPageControl:= 0;
  end;
end;

procedure TFrmImportacao.actSubstituirMemoriaLocalExecute(Sender: TObject);
  var
    Caminho_Copia,enderecoMemoria: String;
begin
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  enderecoMemoria := FrmPrincipal.Caption;
  enderecoMemoria:= ExtractFilePath(enderecoMemoria)+'\dbMemoria.mdb';
  Caminho_Copia := ExtractFilePath(Application.ExeName) + 'TEMPMemoria.mdb';
  FrmPrincipal.compactarDB(Caminho_Copia,false,true,FrmDataModule.ADOConnectionColibri);
  FrmDataModule.ADOConnectionMemoria.Connected:= false;
  CopyFile(PChar(Caminho_Copia), PChar(enderecoMemoria), false);
  FrmPrincipal.conectarBDDireto(enderecoMemoria,FrmDataModule.ADOConnectionMemoria);
  btnClearFiltroAPLAT.Click;
end;

procedure TFrmImportacao.actAnalisaExecutanteAPLATExecute(Sender: TObject);
  var
    ExecutanteAPLAT: TExecCadastro;
begin
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.
  ADOQueryImportarExecutanteAPLAT.RecordCount,
  'Analisando Executantes Cadastrados...');
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Active:= false;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Active:= true;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.Enabled:= false;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.First;
  while not FrmDataModule.ADOQueryImportarExecutanteAPLAT.Eof do
  begin
    try
      ExecutanteAPLAT:= AnalisaExecutanteAPLAT(FrmDataModule.
      DataSourceImportarExecutanteAPLAT.DataSet.FieldByName('CodigoSAP').asString);

      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Edit;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('txtTipoetapaServico').asString:= ExecutanteAPLAT.TipoEtapaServico;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('txtFuncao').asString:= ExecutanteAPLAT.Funcao;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('txtEmpresa').asString:= ExecutanteAPLAT.Empresa;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('txtNomeExecutante').asString:= ExecutanteAPLAT.NomeExecutante;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
      FieldByName('txtDocumento').asString:= ExecutanteAPLAT.Documento;
      //================================================
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Post;
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    except
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Cancel;
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.Next;
    end;
  end;
  ExecutanteAPLAT:= AnalisaExecutanteAPLAT(FrmDataModule.
  DataSourceImportarExecutanteAPLAT.DataSet.FieldByName('CodigoSAP').asString);
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Edit;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('txtTipoetapaServico').asString:= ExecutanteAPLAT.TipoEtapaServico;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('txtFuncao').asString:= ExecutanteAPLAT.Funcao;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('txtEmpresa').asString:= ExecutanteAPLAT.Empresa;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('txtNomeExecutante').asString:= ExecutanteAPLAT.NomeExecutante;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
  FieldByName('txtDocumento').asString:= ExecutanteAPLAT.Documento;
  //=============================================================
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.First;
  FrmDataModule.DataSourceImportarExecutanteAPLAT.Enabled:= true;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmImportacao.actImportarExecutanteAPLATExecute(Sender: TObject);
begin
  if fileexists(edtEnderecoExecutanteAPLAT.Text) then
  begin
    FrmPrincipal.registroEscrever('POB_APLAT', edtEnderecoExecutanteAPLAT.Text);
    if Application.MessageBox(PChar(
    'Deseja realmente importar a planilha de executantes disponíveis no APLAT?'),
    '.::ATENÇÃO::.',36) = 6 then
    begin
      FrmPrincipal.registroEscrever('POB_APLAT',edtEnderecoExecutanteAPLAT.Text);
      btnClearFiltroAPLAT.Click;
      actExecutanteAPLAT.Execute;
      actAnalisaExecutanteAPLAT.Execute;
      edtEnderecoExecutanteAPLAT.Text:= FrmPrincipal.registroEndereco('POB_APLAT');
    end;
  end
  else
  begin
    MessageBox(0,'O endereço da planilha Excel não existe e operação foi cancelada!',
    'Colibri', MB_ICONWARNING);
  end;
end;

procedure TFrmImportacao.actExecutanteAPLATExecute(Sender: TObject);
var
  Excel, Sheet: Variant;
  i,matrizLinha,ARow: Integer;
  texto: String;
begin
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Close;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.SQL.Clear;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.SQL.Add(
  'SELECT tblExecutanteAPLAT.* FROM tblExecutanteAPLAT '+
  'ORDER BY NomeExecutante;');
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Open;
  FrmDataModule.ADOQueryImportarExecutanteAPLAT.Active:= true;
  try
    try
      ARow:= 7;
      try
        Excel := CreateOleObject('Excel.Application');
      except
        MessageBox(0,'Verifique se o Microsoft Office Excel esta instalado na sua máquina.',
        'Excel', MB_ICONERROR);
      end;
      // Abrir o arquivo
      try
        Excel.Workbooks.Open(edtEnderecoExecutanteAPLAT.Text);
      except
        MessageBox(0,'Arquivo não encontrado, verifique o endereço do arquivo!',
        'Excel', MB_ICONERROR);
      end;
      // Abrir a primeira planilha do arquivo
      Sheet := Excel.WorkSheets[1];
      // ============================================
      //Excluir todos os registros atuais
      // ============================================
      FrmPrincipal.deleteQueryRapido(FrmDataModule.ADOQueryImportarExecutanteAPLAT,
      'tblExecutanteAPLAT');
      // ============================================
      //Verificar a ultima Linha
      // ============================================
      matrizLinha:= (Sheet.Cells.SpecialCells(11).Row);
      // Inicializar ProgressBar
      FrmPrincipal.ProgressBarIncializa(matrizLinha,
      'Importando Arquivo Excel...');
      // ============================================
      // Preencher as planilhas
      // ============================================
      FrmDataModule.DataSourceImportarExecutanteAPLAT.Enabled := false;
      actInterromper.Enabled:= true;
      for i := ARow to matrizLinha do
      begin
        Application.ProcessMessages;
        if Interromper and (Application.MessageBox(PChar('Deseja realmente cancelar o processo?'),
        '.::ATENÇÃO::.',36) = 6) then
        begin
          Interromper := False;
          break;
        end
        else
          Interromper := False;

        try
          // Inserir novo registro na tabela de dados
          FrmDataModule.ADOQueryImportarExecutanteAPLAT.Insert;
          // Preencher as colunas do registro na tabela de dados
          texto:= Excel.Workbooks[1].Sheets[1].Cells[i,1];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('CamaLeito').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,1];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('NomeExecutante').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,2];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('codigoSAP').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,3];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('Funcao').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,4];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('DataEmbarque').AsDateTime :=
          StrToDate(Excel.Workbooks[1].Sheets[1].Cells[i,5]);
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('DataDesembarque').AsDateTime :=
          StrToDate(Excel.Workbooks[1].Sheets[1].Cells[i,6]);
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('Empresa').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,7];
          FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
          FieldByName('DiasEmbarcado').AsString := Excel.Workbooks[1].Sheets[1].Cells[i,8];

          FrmDataModule.ADOQueryImportarExecutanteAPLAT.Post;
        except
          FrmDataModule.ADOQueryImportarExecutanteAPLAT.Cancel;
        end;
        // Incremento ProgressBar
        FrmPrincipal.ProgressBarIncremento(1);
      end;
      FrmDataModule.ADOQueryImportarExecutanteAPLAT.First;
      FrmDataModule.DataSourceImportarExecutanteAPLAT.Enabled := true;
      actInterromper.Enabled:= false;
      // Atualizar ProgressBar
      FrmPrincipal.ProgressBarAtualizar;
    finally
      //Fechar Excel
      Excel.Application.DisplayAlerts := False; // Alle Rückfragen ausstellen
      Excel.Quit;
      Excel := Unassigned;
      Excel := 0;
    end;
  except
    MessageBox(0,'Ocorreu um erro durante a importação e a operação foi cancelada!',
    'Excel', MB_ICONERROR);
    FrmDataModule.ADOQueryImportarExecutanteAPLAT.First;
    FrmDataModule.DataSourceImportarExecutanteAPLAT.Enabled := true;
    // Atualizar ProgressBar
    FrmPrincipal.ProgressBarAtualizar;
    //Fechar Excel
    Excel.Application.DisplayAlerts := False; // Alle Rückfragen ausstellen
    Excel.Quit;
    Excel := Unassigned;
    Excel := 0;
  end;
end;

procedure TFrmImportacao.DBGridExecutanteAPLATDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridExecutanteAPLAT,ColunasAPLAT,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'txtTipoEtapaServico')then
  begin
    if (FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
    FieldByName('txtTipoEtapaServico').AsString = 'NÃO DEFINIDO')  then
    begin
      DBGridExecutanteAPLAT.Canvas.Brush.Color:= clRed;
      DBGridExecutanteAPLAT.Font.Color:= clBlack;
      DBGridExecutanteAPLAT.Canvas.FillRect(Rect);
      DBGridExecutanteAPLAT.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
  if (Column.Field.FieldName = 'txtFuncao')then
  begin
    if (FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
    FieldByName('txtFuncao').AsString = 'NÃO DEFINIDO')  then
    begin
      DBGridExecutanteAPLAT.Canvas.Brush.Color:= clRed;
      DBGridExecutanteAPLAT.Font.Color:= clBlack;
      DBGridExecutanteAPLAT.Canvas.FillRect(Rect);
      DBGridExecutanteAPLAT.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
  if (Column.Field.FieldName = 'txtEmpresa')then
  begin
    if (FrmDataModule.DataSourceImportarExecutanteAPLAT.DataSet.
    FieldByName('txtEmpresa').AsString = 'NÃO DEFINIDO')  then
    begin
      DBGridExecutanteAPLAT.Canvas.Brush.Color:= clRed;
      DBGridExecutanteAPLAT.Font.Color:= clBlack;
      DBGridExecutanteAPLAT.Canvas.FillRect(Rect);
      DBGridExecutanteAPLAT.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end;
end;

procedure TFrmImportacao.editEndereco(edt: TEdit);
begin
  OpenDialog1.Filter := 'Microsoft Excel |*.xlsx|Microsoft Excel 97-2003|*.xls|Dados XML|*.xml|Página da Web|*.htm;*.html';
  OpenDialog1.DefaultExt := '*.xlsx';
  //OpenDialog1.DefaultExt:= '*.xls';
  OpenDialog1.FileName:= edt.Text;
  if OpenDialog1.Execute then
    edt.Text:= OpenDialog1.FileName;
end;

procedure TFrmImportacao.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
  FrmImportacao:=nil;
end;

procedure TFrmImportacao.FormCreate(Sender: TObject);
begin
  FrmPrincipal.MDIChildCreated(self.Handle);
  PageControlImportacao.TabIndex:=0;
  Interromper:= false;
  tempoExcedido:= false;
  actInterromper.Enabled:= false;
  FrmDataModule.ADOQueryTipoEtapaServico.Active:= false;
  FrmDataModule.ADOQueryTipoEtapaServico.Active:= true;
  //=====================================
  PageControlImportacao.Pages[0].TabVisible:= true;
  //Incicialização
  FrmDataModule.setFilterDBGrid(DBGridExecutanteAPLAT);
  //Carregar endereços
  edtEnderecoExecutanteAPLAT.Text:= FrmPrincipal.registroEndereco('POB_APLAT');
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)) then
  begin
    DBGridExecutanteAPLAT.ReadOnly:= false;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= true;
    actAnalisaExecutanteAPLAT.Enabled:= true;
    actND_APLAT.Enabled:= true;
    actExcluirAPLAT_Executantes.Enabled:= true;
    actEnderecoExecutanteAPLAT.Enabled:= true;
  end
  else if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) then
  begin
    DBGridExecutanteAPLAT.ReadOnly:= false;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= true;
    actAnalisaExecutanteAPLAT.Enabled:= true;
    actND_APLAT.Enabled:= true;
    actExcluirAPLAT_Executantes.Enabled:= true;
    actEnderecoExecutanteAPLAT.Enabled:= true;
  end
  else if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO) then
  begin
    DBGridExecutanteAPLAT.ReadOnly:= true;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= true;
    actAnalisaExecutanteAPLAT.Enabled:= true;
    actND_APLAT.Enabled:= true;
    actExcluirAPLAT_Executantes.Enabled:= true;
    actEnderecoExecutanteAPLAT.Enabled:= true;
  end
  else if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT) then
  begin
    DBGridExecutanteAPLAT.ReadOnly:= true;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= true;
    actAnalisaExecutanteAPLAT.Enabled:= true;
    actND_APLAT.Enabled:= true;
    actExcluirAPLAT_Executantes.Enabled:= true;
    actEnderecoExecutanteAPLAT.Enabled:= true;
  end
  else if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILHOTELARIA) then
  begin
    DBGridExecutanteAPLAT.ReadOnly:= true;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= true;
    actAnalisaExecutanteAPLAT.Enabled:= true;
    actND_APLAT.Enabled:= true;
    actExcluirAPLAT_Executantes.Enabled:= true;
    actEnderecoExecutanteAPLAT.Enabled:= true;
  end
  else
  begin
    DBGridExecutanteAPLAT.ReadOnly:= true;
    //Executantes APLAT
    actImportarExecutanteAPLAT.Enabled:= false;
    actAnalisaExecutanteAPLAT.Enabled:= false;
    actND_APLAT.Enabled:= false;
    actExcluirAPLAT_Executantes.Enabled:= false;
    actEnderecoExecutanteAPLAT.Enabled:= false;
  end;
  actProcurar.Execute;
end;

procedure TFrmImportacao.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmImportacao.PanelTituloAjudaMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmImportacao.SpeedButton4Click(Sender: TObject);
begin
  PanelAjuda.Visible:= false;
  Splitter1.Visible:= false;
end;

procedure TFrmImportacao.StringGridNDFixedCellClick(Sender: TObject; ACol,
  ARow: Integer);
begin
  FrmPrincipal.clasifica(StringGridND,ACol,true);
  AutoFitGrade(StringGridND);
end;

procedure TFrmImportacao.StringGridNDSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
  var
    R: TRect;
begin
  if ((ACol = 1) AND(ARow > 0)) then
  begin
    R := StringGridND.CellRect(ACol, ARow);
    R.Left := R.Left + StringGridND.Left;
    R.Right := R.Right + StringGridND.Left;
    R.Top := R.Top + StringGridND.Top;
    R.Bottom := R.Bottom + StringGridND.Top;
    ComboBoxTipoEtapaServico.Left := R.Left + 1;
    ComboBoxTipoEtapaServico.Top := R.Top + 1;
    ComboBoxTipoEtapaServico.Width := (R.Right + 1) - R.Left;
    ComboBoxTipoEtapaServico.Height := (R.Bottom + 1) - R.Top;
    ComboBoxTipoEtapaServico.Visible := True;
    ComboBoxTipoEtapaServico.SetFocus;
  end
  else
    ComboBoxTipoEtapaServico.Visible := false;
  if ((ACol = 2) AND(ARow > 0)) then
  begin
    R := StringGridND.CellRect(ACol, ARow);
    R.Left := R.Left + StringGridND.Left;
    R.Right := R.Right + StringGridND.Left;
    R.Top := R.Top + StringGridND.Top;
    R.Bottom := R.Bottom + StringGridND.Top;
    ComboBoxFuncao.Left := R.Left + 1;
    ComboBoxFuncao.Top := R.Top + 1;
    ComboBoxFuncao.Width := (R.Right + 1) - R.Left;
    ComboBoxFuncao.Height := (R.Bottom + 1) - R.Top;
    ComboBoxFuncao.Visible := True;
    ComboBoxFuncao.SetFocus;
  end
  else
    ComboBoxFuncao.Visible := false;
  if ((ACol = 3) AND(ARow > 0)) then
  begin
    R := StringGridND.CellRect(ACol, ARow);
    R.Left := R.Left + StringGridND.Left;
    R.Right := R.Right + StringGridND.Left;
    R.Top := R.Top + StringGridND.Top;
    R.Bottom := R.Bottom + StringGridND.Top;
    ComboBoxEmpresa.Left := R.Left + 1;
    ComboBoxEmpresa.Top := R.Top + 1;
    ComboBoxEmpresa.Width := (R.Right + 1) - R.Left;
    ComboBoxEmpresa.Height := (R.Bottom + 1) - R.Top;
    ComboBoxEmpresa.Visible := True;
    ComboBoxEmpresa.SetFocus;
  end
  else
    ComboBoxEmpresa.Visible := false;



  CanSelect := True;
end;

procedure TFrmImportacao.Timer1Timer(Sender: TObject);
begin
  ShowMessage('Tempo excedido, tente novamente!');
end;

function TFrmImportacao.VerificaCodigoTipoEtapaServico(
  TipoEtapaServico: String): Integer;
begin
  FrmDataModule.ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico.Active := false;
  FrmDataModule.ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico.Parameters.Items[0].Value:=
  TipoEtapaServico;
  FrmDataModule.ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico.Active := true;
  try
   Result:= FrmDataModule.DataSourceConsultaCodigoTipoEtapaServico_TipoEtapaServico.
   DataSet.FieldByName('idTipoEtapaServico').AsInteger;
  except
    Result:= 0;
  end;
end;

function TFrmImportacao.AnalisaExecutanteAPLAT(CodigoSAP: String): TExecCadastro;
  var
    txtTipoEtapaServico,txtEmpresa,txtFuncao,NomeExecutante,
    CPF,Documento: String;
begin
  CPF:= FrmPrincipal.VerificaCPF(CodigoSAP);
  if ((CodigoSAP = '')or(CodigoSAP = '0')) then
    CodigoSAP:= '999999999';
  try
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(
    'SELECT tblExecutante.* '+
    'FROM tblExecutante '+
    'WHERE ((CodigoSAP like '+QuotedStr(CodigoSAP)+
    ')OR(CPF like '+QuotedStr(CodigoSAP)+
    ')OR(OutroDocumento like '+QuotedStr(CodigoSAP)+'));');
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Active:= true;
    txtTipoEtapaServico:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtTipoEtapaServico').AsString;
    txtEmpresa:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtEmpresa').AsString;
    txtFuncao:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtFuncao').AsString;
    NomeExecutante:= FrmPrincipal.TextoMaiusculo(FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('NomeExecutante').AsString);
    if CPF = '' then
      Documento:= FrmDataModule.DataSourceTemporarioDBConsulta1.
      DataSet.FieldByName('OutroDocumento').AsString
    else
      Documento:= CPF;
    //==========================================================================
    if txtTipoEtapaServico <> '' then
    begin
      Result.NomeExecutante:= NomeExecutante;
      Result.TipoEtapaServico:= txtTipoEtapaServico;
      Result.Funcao:= txtFuncao;
      Result.Empresa:= txtEmpresa;
      Result.Documento:= Documento;
    end
    else
    begin
      Result.NomeExecutante:= 'NÃO DEFINIDO';
      Result.TipoEtapaServico:= 'NÃO DEFINIDO';
      Result.Funcao:= 'NÃO DEFINIDO';
      Result.Empresa:= 'NÃO DEFINIDO';
      Result.Documento:= 'NÃO DEFINIDO';
    end;
  except
    Result.NomeExecutante:= 'NÃO DEFINIDO';
    Result.TipoEtapaServico:= 'NÃO DEFINIDO';
    Result.Funcao:= 'NÃO DEFINIDO';
    Result.Empresa:= 'NÃO DEFINIDO';
    Result.Documento:= 'NÃO DEFINIDO';
  end;
end;

procedure TFrmImportacao.ComboBoxEmpresaCloseUp(Sender: TObject);
begin
  StringGridND.Cells[3,StringGridND.Row]:= (ComboBoxEmpresa.Text);
  ComboBoxEmpresa.Visible := False;
end;

procedure TFrmImportacao.ComboBoxFuncaoCloseUp(Sender: TObject);
begin
  StringGridND.Cells[2,StringGridND.Row]:= (ComboBoxFuncao.Text);
  ComboBoxFuncao.Visible := False;
end;

procedure TFrmImportacao.ComboBoxTipoEtapaServicoCloseUp(Sender: TObject);
begin
  StringGridND.Cells[1,StringGridND.Row]:= (ComboBoxTipoEtapaServico.Text);
  ComboBoxTipoEtapaServico.Visible := False;
end;

procedure TFrmImportacao.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
