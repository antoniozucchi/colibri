
unit untPrincipal;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,
  System.ImageList, Vcl.ImgList, Vcl.ExtCtrls, Vcl.Menus, Vcl.StdCtrls,
  Vcl.ComCtrls, Vcl.Tabs, Vcl.DBCtrls, Vcl.ToolWin, Vcl.Mask, Vcl.CheckLst,ClipBrd,
  Vcl.Buttons, Vcl.Imaging.pngimage, SDL_ProgBar, Data.Win.ADODB,
  ShellAPI,ComOBJ,Math,Registry,SDL_rchart,DateUtils,FileCtrl,
  ActiveX, uAccessDBUtils, untDBGridFilter, Vcl.WinXPickers,
  System.Generics.Collections, System.Generics.Defaults, uZucchi;

type
  TFrmPrincipal = class(TForm)
    MainMenu1: TMainMenu;
    MenuArquivo: TMenuItem;
    Abrirbancodedados1: TMenuItem;
    SalvarBancoDadosComo1: TMenuItem;
    Fechar1: TMenuItem;
    Timer1: TTimer;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    MenuProgramacao: TMenuItem;
    Programaodiria1: TMenuItem;
    StatusBarPrincipal: TStatusBar;
    ImageList1: TImageList;
    MenuCadastro: TMenuItem;
    Plataforma1: TMenuItem;
    Executantes1: TMenuItem;
    Embarcacao1: TMenuItem;
    ProgressBarPrincipal: TProgBar;
    N1: TMenuItem;
    CadastroUsuario1: TMenuItem;
    ProgramacaoDiaria2: TMenuItem;
    Image2: TImage;
    MenuSistema: TMenuItem;
    Sobre1: TMenuItem;
    N2: TMenuItem;
    GerenciarSolicitaes1: TMenuItem;
    GerenciarTransportes1: TMenuItem;
    DateTimePickerMinima: TDateTimePicker;
    ServiosProgramados1: TMenuItem;
    ExecutantesProgramadosporPerodo1: TMenuItem;
    ControledeGeradores1: TMenuItem;
    AplatanexodePT1: TMenuItem;
    ActionManager1: TActionManager;
    actVerificaVersao: TAction;
    actAbrirDB: TAction;
    CompactarBancoDados1: TMenuItem;
    mdiChildrenTabs: TTabSet;
    MemoPrincipal: TMemo;
    actExcel: TAction;
    MaskEditData: TMaskEdit;
    StatusBar1: TStatusBar;
    ConexoLOCAL1: TMenuItem;
    actSalvatagemPlataforma: TAction;
    actConverterDB: TAction;
    Converterverso1: TMenuItem;
    MovimentacaoCarga1: TMenuItem;
    CondicaoMarEmbarcacao1: TMenuItem;
    N3: TMenuItem;
    Ajuda1: TMenuItem;
    actMatrizExecutanteCadastro: TAction;
    ResumoProgramao1: TMenuItem;
    SituaodosEquipamentoseAcessodasPlataformas1: TMenuItem;
    actMatrizForaOperacao: TAction;
    EmbarquePassageiros1: TMenuItem;
    ManifestodeEmbarque1: TMenuItem;
    APLAT1: TMenuItem;
    procedure SalvarBancoDadosComo1Click(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Programaodiria1Click(Sender: TObject);
    procedure mdiChildrenTabsChange(Sender: TObject; NewTab: Integer;
      var AllowChange: Boolean);
    procedure Fechar1Click(Sender: TObject);
    procedure rocarLogin1Click(Sender: TObject);
    procedure Plataforma1Click(Sender: TObject);
    procedure Executantes1Click(Sender: TObject);
    procedure Embarcacao1Click(Sender: TObject);
    procedure StatusBarPrincipalDrawPanel(StatusBar: TStatusBar;
      Panel: TStatusPanel; const Rect: TRect);
    procedure CadastroUsuario1Click(Sender: TObject);
    procedure ProgramacaoDiaria2Click(Sender: TObject);
    procedure Sobre1Click(Sender: TObject);
    procedure GerenciarSolicitaes1Click(Sender: TObject);
    procedure GerenciarTransportes1Click(Sender: TObject);
    procedure ServiosProgramados1Click(Sender: TObject);
    procedure ExecutantesProgramadosporPerodo1Click(Sender: TObject);
    procedure ControledeGeradores1Click(Sender: TObject);
    procedure AplatanexodePT1Click(Sender: TObject);
    procedure actVerificaVersaoExecute(Sender: TObject);
    procedure actAbrirDBExecute(Sender: TObject);
    procedure CompactarBancoDados1Click(Sender: TObject);
    procedure PanelTituloAjuda1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure DBComboBoxPreventivaKeyPress(Sender: TObject; var Key: Char);
    procedure ComboBoxOperadorKeyPress(Sender: TObject; var Key: Char);
    procedure FormDestroy(Sender: TObject);
    procedure ConexoLOCAL1Click(Sender: TObject);
    procedure actSalvatagemPlataformaExecute(Sender: TObject);
    procedure actConverterDBExecute(Sender: TObject);
    procedure MovimentacaoCarga1Click(Sender: TObject);
    procedure CondicaoMarEmbarcacao1Click(Sender: TObject);
    procedure Ajuda1Click(Sender: TObject);
    procedure actMatrizExecutanteCadastroExecute(Sender: TObject);
    procedure ResumoProgramao1Click(Sender: TObject);
    procedure SituaodosEquipamentoseAcessodasPlataformas1Click(Sender: TObject);
    procedure actMatrizForaOperacaoExecute(Sender: TObject);
    procedure ManifestodeEmbarque1Click(Sender: TObject);
    type
      TExecutanteProgramado = record
        ListaDestinos: TStrings;
        booleanProgramado: Boolean;
    end;
    type
      TTransbordoLeg = record
        IDProgramacaoExecutante: Integer;
        Origem, Destino: string;
        NomeExecutante: string;
        TipoEtapaServico: string;
        Funcao: string;
        Empresa: string;
        DistKm: Double;
        HasDist: Boolean;
        HoraInicioFicticia: TDateTime;
      end;


  private
    { Private declarations }
    function versaoAtualizada(versaoPRG,versaoDB: String): Integer;
    function versaoExe: String;
    procedure CopiarRegistrosRelacionados(AGravarQuery: TADOQuery;
      const AIdOrigem, AIdDestino: Integer; const ACampoLigacao,
      ASQLConsulta: string; const AExecutante: Boolean);
    procedure LimparCamposRTDaProgramacao(
      const ACodigoProgramacaoDiaria: Integer);
    procedure CriarTabelasProntidao;
    //procedure Descomprime(ArquivoZip: TFileName; DiretorioDestino: string);

  public
    
    var
    SelPlataforma: Boolean;
    ultimaSelecao: TPointFloat;
    logChave,logMaquina,logPerfil,filtroFO,filtroRO,
    MagicSQLQuery,MagicFonteDB,MagicFieldName,MagicIndexColuna,MagicTextoProcura: string;
    MouseDownSpot: TPoint;
    Capturing,booLOCAL,Interromper: Boolean;
    HintPadrao,txtBarraProgresso,enderecoColibriRegistro: String;

    MatrizSalvatagem,MatrizForaOperacao,
    MatrizExecutanteCadastro: array of array of String;
    const
    CoresAuto: array[0..14] of TColor =
    (clMaroon,clGreen,clOlive,clNavy,clPurple,clTeal,clRed,clLime,clYellow,clBlue,
    clFuchsia,clAqua,clBlack,clGray,clSilver);

    PERFILRT = 'Requisição de Transporte';
    PERFILADM = 'Administrador';
    PERFILPROGRAMACAO = 'Programação';
    PERFILHOTELARIA = 'Hotelaria';
    PERFILSUPERVISAO = 'Supervisão';
    { Public declarations }
    procedure Force_Reconnect;
    procedure ADOConnection_Reconnect(Conn: TADOConnection);
    //Registros Windows
    function FormatarCPF(CPF: String): String;
    procedure registroEscrever(keyString,Endereco: String);
    function registroEndereco(keyString: String): String;
    //Conexão DB
    function conectarBD(Endereco: String;sourceADOConnection: TADOConnection): String;
    procedure conectarBDDireto(Endereco: String;sourceADOConnection: TADOConnection);
    procedure MDIChildCreated(const childHandle : THandle);
    procedure MDIChildDestroyed(const childHandle : THandle);
    //DBGrid e StringGrid
    procedure ClassificaDBGrid(Grid: TDBGrid;sourceADOQuery: TADOQuery;TipoSORT: Integer);
    function CharAscDesc(const Ch: Char; const S: string): Boolean;

    procedure SetupGridPickListSQL(Conection: TADOConnection; const FieldName, sql: String;
      var Tabela: TDBGrid; DeletarRepetidos: Boolean);
    procedure SetupGridFilterPickListSQL(Conection: TADOConnection;
      const FieldName, sql: String; var Tabela: TFilterDBGrid;
      DeletarRepetidos: Boolean);

    procedure SetupGridPickListLista(FieldName: string; Lista: TStringList; Tabela: TDBGrid);
    procedure ConfigGridLayout(Grid: TStringGrid; ACol,ARow: Integer; aRect: TRect);
    procedure carregarCheckListBox(ADOQuery: TADOQuery; checkList: TCheckListBox;FildName:String);
    procedure carregarRadioGroup(ADOQuery: TADOQuery; RadioGroup: TRadioGroup; FildName,PrimeiroItem: String);
    procedure AutoSizeDBGrid(Grid: TDBGrid);
    procedure DeleteRow(Grid: TStringGrid; ARow: Integer);
    function InsertRow1(StrGrid: TStringGrid; Linha:integer):integer;
    procedure DeleteCol(Grid: TStringGrid; ACol: Integer);
    function Clasifica(Grid: TStringGrid; ACol: Integer;indicar: Boolean): String;
    procedure selecaoDBGrid(Grid: TDBGrid; selecao: boolean; Field: String);
    function palavraBusca(SQLString,FieldBusca,Operador,StringBusca,CondicionalVariaveis: String): String;
    function palavraBuscaAND(FieldBusca,PalavraBusca: String): String;
    function incialListaBusca(CampoString: String): TStringList;
    procedure LayoutPadrao(NomeArquivoMemo:string; GridPadrao: TStringGrid;Tabela: String);
    procedure addColuna(Grid: TDBGrid; Field,Caption,alinha: String;
    numColuna,Largura: Integer; booleanReadOnly: Boolean);

    procedure limparTitleGrid(Grid: TDBGrid);
    procedure AlterarTituloColuna(ColunasLayout:TStringGrid;FieldName,strTitulo: String);
    function indexLayoutFieldName(FieldName: String; ColunasLayout: TStringGrid): Integer;
    function indexLayoutCaption(Caption: String; ColunasLayout: TStringGrid): Integer;
    //Tratamento de Strings
    function DeleteChar(const Ch: Char; const S: string): string;
    //Excel
    procedure GeraTexto(DataSource: TDataSource; Separador: Char; nomeTabela: String);
    function RefToCell(const iCol: Integer; const iRow: Integer): String;
    procedure importarExcel(nomeTabela: String;
    ACol,ARow,TabSheet: Integer; Tabela: TDBGrid; ADOQuery: TADOQuery);
    //Tratamento de Query
    procedure deleteQuery(ADOQuery: TADOQuery; nomeTabela: String);
    procedure deleteQueryRapido(ADOQuery: TADOQuery; nomeTabela: String);
    procedure procuraQuerySimples(txtField,txtTabela: String;edtLocal: TEdit; sourceQuery: TADOQuery);
    //ProgressBar
    procedure ProgressBarAtualizar;
    procedure ProgressBarIncializa(MaxPosition: integer; Texto: String);
    procedure ProgressBarIncremento(Incremento: integer);
    procedure MenssagemStatus(Texto: String);
    //Configurações
    procedure SetupCheckListSQL1(FieldName,txtProcura: String;CheckList: TCheckListBox;
    ListaGroup: TStringList);
    procedure carregarComboBox(conection: TADOConnection; txtField,sql: String; comboBox: TComboBox);
    procedure selComboBox(combo: TComboBox; txt: String);
    procedure deleteRepetidosCombo(ComboBox: TComboBox);
    procedure deleteRepetidosCheckListBox(CheckListBox: TCheckListBox);
    procedure deleteRepetidosList(Lista: TStringList);
    procedure carregarLoginUsuario(Chave: String);
    function NomeMaquina: String;
    procedure GravarCanceladoAprovado(idProgramacaoDiaria: Integer);
    procedure compactarDB(EnderecoArquivo: String;booleanPerguntar,conexaoColibri: Boolean;
    SourceADOConnection: TADOConnection);
    function VerificaCPF(Texto: String): String;
    function SomenteNumero(Texto: String): String;
    function FormataCPF(Texto: String): String;
    //Funções
    function DistanciaPontos(X1,Y1,X2,Y2: Double): Double;
    function carregaDataMinima(Servico: Boolean): TDateTime;
    function SomaHoras(Hora1, Hora2 :String): String;
    function SubtrairHoras(Hora1,Hora2: String): String;
    Function DecimalHoras(cValor : Double): String;
    function LatidudeLongitude_Graus(cValor: Double): String;
    function LatLong_XY(Latitude,Longitude: Double): TPointFloat;
    function isNumeric(txt: String): Integer;
    function isData(txt: String): Boolean;
    Function usuarioWindows: string;
    function TextoMaiusculo(Texto: String): String;
    function dataFormatar(SourceData: String): TDateTime;
    function dataSAP(SourceData: String): String;
    function DataSAP_ToDate(const SourceData: string): TDateTime;
    function dataLimpa(SourceData: String): String;
    function CalcNumCanceladosAprovado(idProgramacaoDiaria,Tipo: Integer): Integer;
    function CalcNumExecutantes(idProgramacaoDiaria: Integer): Integer;
    procedure AvaliarProgramacaoExecutante(idProgramacaoExecutante,Fonte: Integer;
    StatusProgramacao,Motivo: String);
    procedure ProcuraQuery(SQLBase: String; sourceQuery: TADOQuery;StatusBar: TStatusBar);

    function verificaOperacao(StatusSistemaOP: String): String;
    //Layout Grid
    //PDF
    procedure ImageExcluir(NomeArquivo: String);
    function dadosColuna(txtCaption: String; ColunasLayout: TStringGrid): TStringList;
    procedure CarregaFiltrosProcura(ColunasLayout: TStringGrid);
    function DiasFinalUtil(DataInicio: TDateTime; DiasUteis: Integer):TDateTime;
    function DiasFinalCorridos(DataInicio: TDateTime; DiasCorridos: Integer):TDateTime;
    function corrigirData(strData: TDateTime): String;
    procedure ImportDataAccess(const AccessDb, TableName, ExcelFileName:String);
    procedure ImportExcelAcess(enderecoExcel,enderecoDB,Tabela: String;ListaColunas: TStringList;ADOQuery: TADOQuery);
    procedure DownLoadQuery(GravarQuery: TADOQuery;
    GravarDataSource: TDataSource; parametroREF, parametroGravar,campo1: Integer;
    parametroString, nomeTabela,SQLConsulta: String; booleanInsert,booleanExecutante: Boolean);
    function CountChecked(Grid: TStringGrid): Integer;
    procedure limparStringGrid(Grid: TStringGrid);
    function selecionarServico(GRID: TStringGrid; ServicoPT,TextoBreveOP,OrdemManutencao: String): Boolean;
    procedure selecaoPlataforma(RChartMapa: TRChart; noPlataforma: TPointFloat;StatusBarMapa: TStatusBar);
    procedure CotaReta(RChartMapa: TRChart; no1,no2: TPointFloat;StatusBarMapa: TStatusBar);
    procedure CopiarProgramacao(const ADataProgramacao: TDateTime;
      ASource: TDataSource);
    procedure ProgramarExecutantes(Origem: String;idProgramacaoDiaria: Integer;
    QueryExecutante,QueryProgramacao: TADOQuery; DataSourceExecutante,DataSourceProgramacao: TDataSource);
    function RT_Local(Plataforma: String): String;
    function RT_CentroCusto(Plataforma: String): String;
    function incluirServicoPadrao(TipoEtapaServico: String): String;
    
    function verificaCadastroExecutante(OrigemCadastro,NomeExecutante,Funcao: String): Boolean;
    function HoraParaMin(Hora: String): Integer;
    function YEquacaoReta(X,X1,Y1,X2,Y2: Double): Double;
    function incluirProgramacao(DataProgramacao,
    Destino, txtTipoEtapaServico, CriadoPor,Computador: String;DataCriacao: TDate): Integer;
    procedure selecaoGrid(Grid: TStringGrid; Status: TStatusBar; txt: String);

    function RetiraEnter(aText : string): string;
    function OrigemPlataformas: String;
    procedure configurarCheckListBox(CheckListBox: TCheckListBox;DBGrid: TDBGrid;Layout: TStringGrid;BooleanLimparColunas: Boolean);
    function buscarString(strBusca,strTexto: String): Boolean;
    procedure DesenharCalendario(Calendario: TStringGrid; PanelNome: TPanel;PrimeiroDiaMes: TDateTime);
    function PrimeiraLetraMaiscula(Str: string): string;
    procedure AbrirBancoDados(enderecoColibri,Chave: String;BooleanSplash: Boolean);
    procedure criarVariavelTempoExecucao(Field,Tabela,TipoField: String; sourceQuery: TADOQuery);
    procedure msgSplash(texto1,texto2: String);
    function ForaOperacao(Plataforma: String; IndexSistema: Integer): String;
  end;

var
  FrmPrincipal: TFrmPrincipal;


implementation
  uses untDataModule, untConsultaProgramacao,untFrmlogin,
  untEmbarcacao,untExecutante, untPlataforma, untFrmCadastroUsuario, untProgramacaoDiaria,
  untFrmSobre, untGerenciarSolicitacoes,
  untGerenciarEmbarcacoes, untConsultaServicosProgramados,
  untConsultaExecutantesProgramados, untControleGeradores,
  untAgendaIntervencao,untMovimentacaoCarga,untTelaAbertura,
  untSituacaoEquipamentoAcesso,
  untCondicaoEmbarcacao, untFrmTabela, untConexaoLOCAL, untAplatAnexos, untFrmManifestoEmbarque;

{$R *.dfm}

{ TFrmPrincipal }

procedure TFrmPrincipal.AbrirBancoDados(enderecoColibri,Chave: String;BooleanSplash: Boolean);
  var
    enderecoConsulta, enderecoRT: String;
begin
  try
    enderecoColibri:= conectarBD(enderecoColibri,FrmDataModule.ADOConnectionColibri);
    enderecoConsulta:= ExtractFilePath(enderecoColibri)+'dbConsulta.mdb';
    enderecoRT:= ExtractFilePath(enderecoColibri)+'dbRT.mdb';
    carregarLoginUsuario(Chave);
    //========================================================================
    if BooleanSplash then
    begin
      msgSplash('Conectando ao banco de dados dbColibri.colibri',
      ' no endereço: '+enderecoColibri);
      //------------------------------------
      try
        msgSplash('Conectando ao dbConsulta.mdb',
        ' no endereço: '+enderecoConsulta);
        FrmPrincipal.conectarBDDireto(enderecoConsulta,FrmDataModule.ADOConnectionConsulta);
      except
        MessageBox(0,PChar('Banco de dados de Consulta não encontrado ou corrompido: '+enderecoConsulta),
        'Colibri',MB_ICONEXCLAMATION);
      end;
      try
        msgSplash('Conectando ao dbRT.mdb',
        ' no endereço: '+enderecoRT);
        FrmPrincipal.conectarBDDireto(enderecoRT,FrmDataModule.ADOConnectionRT);
      except
        MessageBox(0,PChar('Banco de dados de RT não encontrado ou corrompido: '+enderecoRT),
        'Colibri',MB_ICONEXCLAMATION);
      end;
    end;
    //Configurações finais
    registroEscrever('Banco de dados', enderecoColibri);
    enderecoColibriRegistro:= enderecoColibri;
    actSalvatagemPlataforma.Execute;
  except
    MessageBox(0, 'Banco de dados não foi aberto e o programa será fechado!',
    'Colibri', MB_ICONERROR);
    Application.Terminate;
  end;
end;

procedure TFrmPrincipal.actAbrirDBExecute(Sender: TObject);
begin
  OpenDialog1.Filter:= 'Arquivo Colibri|*.colibri';
  OpenDialog1.DefaultExt:= '*.colibri';
  try
    OpenDialog1.FileName:= registroEndereco('Banco de dados');
  except
  end;
  if OpenDialog1.Execute then
  begin
    //Abrir banco de dados
    AbrirBancoDados(OpenDialog1.FileName,logChave,false);
    registroEscrever(OpenDialog1.FileName,'Banco de dados');
  end;
end;

procedure TFrmPrincipal.actConverterDBExecute(Sender: TObject);
  var
    versaoDB,versaoPRG: String;
begin
  try
    FrmDataModule.ADOQueryColibri.Active:= true;
    versaoPRG:= FrmPrincipal.VersaoExe;
    versaoDB:= FrmDataModule.DataSourceColibri.DataSet.FieldByName('versao').AsString;
    if versaoDB = '1.6.5.5' then
    begin
      //FrmPrincipal.CriarFieldDB('MarkX1','tblPainelDutos','YESNO',FrmDataModule.ADOQueryTemporarioDBOperacao);
      //FrmPrincipal.CriarFieldDB('MarkX2','tblPainelDutos','YESNO',FrmDataModule.ADOQueryTemporarioDBOperacao);
      CriarFieldDB('MotivoMudanca','tblProgramacaoExecutante',
      'VARCHAR(255)',FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('PalavraChaveMudanca','tblProgramacaoExecutante',
      'VARCHAR(255)',FrmDataModule.ADOQueryTemporarioDBColibri);
      versaoDB:= '1.6.5.6';
    end;
    if versaoDB = '1.6.5.6' then
    begin
      //Não existe alteração de banco de dados
      versaoDB:= '1.6.5.7';
    end;
    if versaoDB = '1.6.5.7' then
    begin
      CriarFieldDB('NumeroProrrogacao','tblPericiaMarinha','INTEGER',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('NumeroProrrogacao','tblPericiaMarinha','INTEGER',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      versaoDB:= '1.6.5.8';
    end;
    if versaoDB = '1.6.5.8' then
    begin
      CriarFieldDB('Descricao','tblMovimentacaoCarga','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      versaoDB:= '1.6.5.9';
    end;
    if versaoDB = '1.6.5.9' then
    begin
      CriarFieldDB('booleanIP19','tblProgramacaoDiaria','YESNO',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('AvaliadoPor','tblProgramacaoDiaria','VARCHAR(10)',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('DataAtualizacao','tblProgramacaoDiaria','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('ComputadorAtualizacao','tblProgramacaoDiaria','VARCHAR(20)',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('TextoBreveOM','tblProgramacaoServico','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('DataAtendimento','tblMovimentacaoCarga','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      AlterarNomeFieldDB('DataNecessidade','Data','tblMovimentacaoCarga',
      FrmDataModule.ADOConnectionConsulta);
      versaoDB:= '1.6.6.0';
    end;
    if versaoDB = '1.6.6.0' then
    begin
      CriarFieldDB('booleanFavorito','tblProgramacaoDiaria','YESNO',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      versaoDB:= '1.6.6.1';
    end;
    if versaoDB = '1.6.6.1' then
    begin
      CriarFieldDB('LogAcao','tblProgramacaoDiaria','Memo',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      versaoDB:= '1.6.6.2';
    end;
    if versaoDB = '1.6.6.2' then
    begin
      AlterarNomeFieldDB('MotivoProgramacao','MotivoCancelamento',
      'tblProgramacaoExecutante',FrmDataModule.ADOConnectionColibri);
      AlterarNomeFieldDB('PalavraChaveProgramacao','PalavraChaveCancelamento',
      'tblProgramacaoExecutante',FrmDataModule.ADOConnectionColibri);
      versaoDB:= '1.6.6.3';
    end;
    if versaoDB = '1.6.6.3' then
    begin
      versaoDB:= '1.6.6.4';
    end;
    if versaoDB = '1.6.6.4' then
    begin
      //Tabela Plataforma
      CriarFieldDB('Painel_Condicao','tblPlataforma','VARCHAR(25)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('NumFAM','tblPlataforma','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('booleanSelecao','tblPlataforma','YESNO',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('Escala','tblPlataforma','DOUBLE',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      versaoDB:= '1.6.6.5';
    end;
    if versaoDB = '1.6.6.5' then
    begin
      //Pericia de Marinha
      CriarFieldDB('RelatorioFlagState','tblPericiaMarinha','VARCHAR(25)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DataReferenciaFlagState','tblPericiaMarinha','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DataAtendimento','tblPericiaMarinha','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      //Reltorio de Marinha
      CriarFieldDB('CartaRelatorio','tblDocumentosMarinha','VARCHAR(25)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('CartaItem','tblDocumentosMarinha','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      //============================================================
      versaoDB:= '1.6.6.6';
    end;
    if versaoDB = '1.6.6.6' then
    begin
      //Planilha de executantes
      CriarFieldDB('booleanCESS','tblExecutante','YESNO',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DataValidadeCESS','tblExecutante','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      //============================================================
      versaoDB:= '1.6.6.7';
    end;
    if versaoDB = '1.6.6.7' then
    begin
      //Planilha Analise de PT
      //CriarFieldDB('Pressurizado','tblAnaliseServicoPT','VARCHAR(5)',
      //FrmDataModule.ADOQueryTemporarioDBOperacao);
      //============================================================
      versaoDB:= '1.6.6.8';
    end;
    if versaoDB = '1.6.6.8' then
    begin
      //Tabela de observações da programação
      CriarTableDB('tblProgramacaoNotas','idProgramacaoNotas',
      '[Notas] Memo,[DataProgramacao] DATETIME, [AvaliadoPor] VARCHAR(10),'+
      '[DataAtualizacao] DATETIME, [Computador] VARCHAR(20)',
      FrmDataModule.ADOConnectionColibri);
      //============================================================
      versaoDB:= '1.6.6.9';
    end;
    if versaoDB = '1.6.6.9' then
    begin
      //Tabela Colibri
      CriarFieldDB('UnidadePetrobras','tblColibri','INTEGER',
      FrmDataModule.ADOQueryTemporarioDBColibri);
      versaoDB:= '1.6.7.0';
    end;
    if versaoDB = '1.6.7.0' then
    begin
      //Tabela tblProgramacaoCalendario
      CriarTableDB('tblProgramacaoCalendario','idProgramacaoCalendario',
      '[DataProgramacao] DATETIME, [Registro] Memo',FrmDataModule.ADOConnectionColibri);
      versaoDB:= '1.6.7.1';
    end;
    if versaoDB = '1.6.7.1' then
    begin
      CriarTableDB('tblFiltroSistemas','idFiltroSitemas',
      '[Item] INTEGER, [Descricao] VARCHAR(255), [TituloColuna] VARCHAR(255), [PalavraBusca] VARCHAR(255),[Condicional] VARCHAR(255)',
      FrmDataModule.ADOConnectionConsulta);
      //Corrigir colunas Balsas
      ExcluirFieldDB('Campo','tblBalsas',
      FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('BP','tblBalsas',
      FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Resposavel','tblBalsas',
      FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Condicao','tblBalsas',
      FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('AnoFabricacao','tblBalsas',
      FrmDataModule.ADOConnectionConsulta);
      versaoDB:= '1.6.7.2';
    end;
    if versaoDB = '1.6.7.2' then
    begin
      //Criar Tabela tblClasseBCVT
      CriarTableDB('tblClasseBCVT','idClasseBCVT',
      '[Classe] VARCHAR(100) ',FrmDataModule.ADOConnectionConsulta);
      versaoDB:= '1.6.7.3';
    end;
    if versaoDB = '1.6.7.3' then
    begin
      //Consulta
      CriarTableDB('tblRTIObservacao','idRTIObservacao',
      '[NumItem] VARCHAR(4), [Medida] VARCHAR(4),[Observacao] VARCHAR(255),[StatusProgramacao] VARCHAR(15),[NotaManutencao] VARCHAR(25)',
      FrmDataModule.ADOConnectionConsulta);
      //Consulta
      CriarFieldDB('Observacao','tblBalsas','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      versaoDB:= '1.6.7.5';
    end;
    if versaoDB = '1.6.7.5' then
    begin
      versaoDB:= '1.6.7.6';
    end;
    if versaoDB = '1.6.7.7' then
    begin
      //Colibri
      CriarTableDB('tblProgramacaoRascunho','idProgramacaoRascunho',
      '[DataProgramacao] DATETIME,[Registro] VARCHAR(255), [Plataforma] VARCHAR(255)',
      FrmDataModule.ADOConnectionColibri);
      //Consulta
      CriarTableDB('tblPalavraChave','idPalavraChave',
      '[booleanSelecao] YESNO,[PalavraChave] VARCHAR(255)',
      FrmDataModule.ADOConnectionColibri);
      versaoDB:= '1.6.7.8';
    end;
    if versaoDB = '1.6.7.8' then
    begin
      //Excluir variaveis tabela Plataforma
      ExcluirFieldDB('Painel_X','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Painel_Y','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Painel_Cor','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Painel_Tipo','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Painel_Boolean','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Link_FEng','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Link_FProc','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Link_MCE','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Painel_Condicao','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('NumFAM','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('Escala','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('ElementoPEP','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      versaoDB:= '1.6.7.9';
    end;
    if versaoDB = '1.6.7.9' then
    begin
      versaoDB:= '1.6.8.0';
    end;
    if versaoDB = '1.6.8.0' then
    begin
      //dbConsulta
      ExcluirTableDB('tblBalsas',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblCentroTrabalho',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblClasseBCVT',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblComentarioForaOperacao',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblComentarioICPI',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblComentarioICPM',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblComentarioRTI',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblControleCarta',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblDocumentosMarinha',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblFotosMarinha',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblPericiaMarinha',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblResumoIndicador',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      ExcluirTableDB('tblFiltroSistemas',FrmDataModule.ADOQueryTemporarioDBConsulta1);

      ExcluirFieldDB('LimiteInferior','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      ExcluirFieldDB('LimiteSuperior','tblPlataforma',FrmDataModule.ADOConnectionConsulta);
      CriarFieldDB('SituacaoGD','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoBCI','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoLinhaBCI','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoAgua','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoBalsa','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoAcesso','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoDegraus','tblPlataforma','VARCHAR(5)',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoNotas','tblPlataforma','MEMO',FrmDataModule.ADOQueryTemporarioDBConsulta1);

      CriarFieldDB('CapPrincipal','tblPlataforma','DOUBLE',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('CapAuxiliar','tblPlataforma','DOUBLE',FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DataRealizacaoDegraus','tblPlataforma','DATETIME',FrmDataModule.ADOQueryTemporarioDBConsulta1);

      versaoDB:= '1.6.9.0';
    end;
    if ((versaoDB = '1.6.9.0') OR (versaoDB = '1.7.0.0')OR (versaoDB = '1.7.0.1')) then
    begin
     //tabela de Programação de Executantes
      CriarFieldDB('RT_Descricao','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('OutroDocumento','tblProgramacaoExecutante','VARCHAR(15)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Modal','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Tipo','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Classe','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_HoraIda','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_HoraVolta','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Mensagem','tblProgramacaoExecutante','VARCHAR(100)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Status','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Erro','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('DataVolta','tblProgramacaoExecutante','DATETIME',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('booleanRecolhimento','tblProgramacaoExecutante','YESNO',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RecolherPara','tblProgramacaoExecutante','VARCHAR(20)',
          FrmDataModule.ADOQueryTemporarioDBColibri);

      ExcluirFieldDB('tblProgramacaoExecutante','RT_HoraInicio',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','booleanBateVolta',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','booleanBateVolta',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_CentroCusto',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_DiagramaRede',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_OperRede',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_Descricao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_Cobertura',
          FrmDataModule.ADOConnectionColibri);


      //Tabela Cadastro de Executantes
      CriarFieldDB('CentroCusto','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DiagramaRede','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('OperRede','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('ElementoPEP','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);

      //Tabela Plataforma
      CriarFieldDB('CentroCusto','tblPlataforma','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DiagramaRede','tblPlataforma','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('OperRede','tblPlataforma','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('ElementoPEP','tblPlataforma','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('NomeSAP','tblPlataforma','VARCHAR(15)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('booleanNaoCriarRT','tblPlataforma','YESNO',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('booleanProntidao','tblPlataforma','YESNO',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('booleanTurno2','tblPlataforma','YESNO',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('RT_Modal','tblPlataforma','VARCHAR(2)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);


      CriarFieldDB('RT_Modal','tblPlataforma','VARCHAR(2)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);



      ExcluirFieldDB('tblColibri','booleanCriarRT',
          FrmDataModule.ADOConnectionColibri);


      //Tabela Programacao de RT
      //ExcluirTableDB('tblProgramacaoRT',FrmDataModule.ADOQueryTemporarioDBColibri);
      //ExcluirTableDB('tblProgramacaoRT',FrmDataModule.ADOQueryTemporarioRT);

      CriarTableDB(
        'tblProgramacaoRT',
        'idProgramacaoRT',
        '[idProgramacaoExecutante] INTEGER, '+

        '[DataIda] DATETIME, '+
        '[DataVolta] DATETIME, '+
        '[RT_HoraIda] VARCHAR(5), '+
        '[RT_HoraVolta] VARCHAR(5), '+

        '[CodigoSAP] VARCHAR(10), '+
        '[TipoDoc] VARCHAR(2), '+
        '[Documento] VARCHAR(30), '+
        '[Passageiro] VARCHAR(80), '+

        '[Origem] VARCHAR(20), '+
        '[txtDestino] VARCHAR(20), '+
        '[RecolherPara] VARCHAR(20), '+

        '[RT_Tipo] VARCHAR(3), '+
        '[RT_Descricao] VARCHAR(80), '+
        '[RT_Requisitante] VARCHAR(10), '+
        '[RT_TelContato] VARCHAR(20), '+
        '[RT_PessoaContato] VARCHAR(40), '+
        '[RT_CentroPlan] VARCHAR(4), '+
        '[RT_GrpPlan] VARCHAR(3), '+
        '[RT_Modal] VARCHAR(5), '+
        '[RT_Classe] VARCHAR(5), '+

        '[RT_CentroCusto] VARCHAR(20), '+
        '[RT_DiagramaRede] VARCHAR(20), '+
        '[RT_OperRede] VARCHAR(4), '+
        '[RT_ElementoPEP] VARCHAR(20), '+

        '[RT_Cobertura] YESNO, '+
        '[RT_Cancelada] YESNO, '+
        '[booleanRecolhimento] YESNO, '+
        '[RT_ImportadaSAP] YESNO, '+

        '[RT] VARCHAR(15), '+
        '[RT_Mensagem] VARCHAR(255), '+
        '[RT_Status] VARCHAR(40), '+
        '[RT_Erro] VARCHAR(100), '+

        '[RT_StatusSAP] VARCHAR(2), '+
        '[RT_StatusSAPDesc] VARCHAR(40), '+

        '[ChavePassageiro] VARCHAR(50), '+
        '[ChaveIda] VARCHAR(255), '+
        '[ChaveVolta] VARCHAR(255), '+
        '[ChaveCompleta] VARCHAR(255)',

        FrmDataModule.ADOConnectionRT
      );

      CriarTableDB(
        'tblConfigRT',
        'idConfigRT',
        '[RT_Tipo] VARCHAR(2), '+
        '[RT_Requisitante] VARCHAR(8), '+
        '[RT_PessoaContato] VARCHAR(20), '+
        '[RT_TelContato] VARCHAR(20), '+
        '[RT_CentroPlan] VARCHAR(4), '+
        '[RT_GrpPlan] VARCHAR(3), '+
        '[HoraInicio_Turno1] VARCHAR(5), '+
        '[HoraVolta_Turno1] VARCHAR(5), '+
        '[HoraInicio_Turno2] VARCHAR(5), '+
        '[HoraVolta_Turno2] VARCHAR(5), '+
        '[RT_Salvar] YESNO, '+
        '[RT_HoraCobertura] DATETIME',
        FrmDataModule.ADOConnectionRT
      );
      CriarFieldDB('RT_Salvar','tblConfigRT','YESNO',
          FrmDataModule.ADOQueryTemporarioRT);

      //Excluir
      ExcluirFieldDB('tblColibri','RT_Requisitante',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_PessoaContato',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_TelContato',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_CentroPlan',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_GrpPlan',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_Modal',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_HoraInicio',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_HoraVolta',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblColibri','RT_HoraCobertura',
          FrmDataModule.ADOConnectionColibri);


      // ======================================================
      // ESPELHO DAS RTS EXISTENTES NO SAP
      // ======================================================
      CriarTableDB(
        'tblRTSapImport',
        'idRTSapImport',
        '[DataImportacao] DATETIME, '+

        '[OrigemConsulta] VARCHAR(15), '+      // PONTUAL / PERIODO
        '[PeriodoInicio] DATETIME, '+
        '[PeriodoFim] DATETIME, '+

        '[QMNUM] VARCHAR(15), '+               // número da RT
        '[QMART] VARCHAR(3), '+                // R3 / R7
        '[IWERK] VARCHAR(4), '+                // centro
        '[INGRP] VARCHAR(3), '+                // grupo planejador

        '[Origem] VARCHAR(20), '+
        '[txtDestino] VARCHAR(20), '+
        '[DataViagem] DATETIME, '+
        '[HoraViagem] VARCHAR(5), '+

        '[PERNR] VARCHAR(10), '+               // usado no R3
        '[TipoDoc] VARCHAR(2), '+              // C / P
        '[Documento] VARCHAR(30), '+
        '[Passageiro] VARCHAR(80), '+
        '[QMTXT] VARCHAR(80), '+

        '[RT_Modal] VARCHAR(5), '+
        '[RT_Classe] VARCHAR(5), '+

        '[StatusItem] VARCHAR(2), '+           // ex.: 09
        '[StatusDescricao] VARCHAR(40), '+     // ex.: Cancelado
        '[RT_Cancelada] YESNO, '+

        '[ChavePassageiro] VARCHAR(50), '+
        '[ChaveIda] VARCHAR(255), '+
        '[ChaveVolta] VARCHAR(255), '+
        '[ChaveCompleta] VARCHAR(255), '+

        '[idProgramacaoRT] INTEGER, '+         // quando reconciliado
        '[idProgramacaoExecutante] INTEGER, '+ // quando reconciliado

        '[ImportadoColibri] YESNO, '+          // True quando vinculado ao processo do Colibri
        '[Observacao] VARCHAR(255)',

        FrmDataModule.ADOConnectionRT
      );

      // ======================================================
      // RESULTADO DA RECONCILIAÇÃO
      // ======================================================
      CriarTableDB(
        'tblReconRT',
        'idReconRT',
        '[DataRecon] DATETIME, '+

        '[idProgramacaoExecutante] INTEGER, '+
        '[idProgramacaoRT] INTEGER, '+
        '[idRTSapImport] INTEGER, '+

        '[ChavePassageiro] VARCHAR(50), '+
        '[ChaveIda] VARCHAR(255), '+
        '[ChaveVolta] VARCHAR(255), '+
        '[ChaveCompleta] VARCHAR(255), '+

        '[RT] VARCHAR(15), '+

        '[TipoRT] VARCHAR(3), '+
        '[Origem] VARCHAR(20), '+
        '[txtDestino] VARCHAR(20), '+
        '[RecolherPara] VARCHAR(20), '+
        '[DataIda] DATETIME, '+
        '[RT_HoraIda] VARCHAR(5), '+
        '[DataVolta] DATETIME, '+
        '[RT_HoraVolta] VARCHAR(5), '+

        '[CodigoSAP] VARCHAR(10), '+
        '[TipoDoc] VARCHAR(2), '+
        '[Documento] VARCHAR(30), '+
        '[Passageiro] VARCHAR(80), '+

        '[RT_Modal] VARCHAR(5), '+
        '[RT_Classe] VARCHAR(5), '+
        '[RT_CentroPlan] VARCHAR(4), '+
        '[RT_GrpPlan] VARCHAR(3), '+

        '[TipoAcao] VARCHAR(15), '+            // CRIAR / VINCULAR / CANCELAR / MANTER / ANALISAR
        '[Motivo] VARCHAR(100), '+
        '[StatusRecon] VARCHAR(20), '+         // PENDENTE / EXECUTADO / ERRO

        '[RT_StatusSAP] VARCHAR(2), '+
        '[RT_StatusSAPDesc] VARCHAR(40), '+

        '[BooleanRecolhimento] YESNO, '+
        '[Divergencia] YESNO, '+
        '[Observacao] VARCHAR(255)',

        FrmDataModule.ADOConnectionRT
      );

      // ======================================================
      // REGRAS DE RECOLHIMENTO DE RT
      // ======================================================
      CriarTableDB(
        'tblRTRegraRecolhimento',
        'idRegraRecolhimento',
        '[Ativa] YESNO, '+
        '[Prioridade] INTEGER, '+
        '[Descricao] VARCHAR(100), '+

        '[Origem] VARCHAR(20), '+
        '[txtDestino] VARCHAR(20), '+
        '[DestinoNaoCriarRT] VARCHAR(3), '+     // '', 'SIM', 'NAO'

        '[RecolherParaTipo] VARCHAR(10), '+     // ORIGEM / DESTINO / FIXO
        '[RecolherParaValor] VARCHAR(20), '+

        '[Observacao] VARCHAR(255)',

        FrmDataModule.ADOConnectionRT
      );


      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusSAPDescricao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusSAPDescricao',
          FrmDataModule.ADOConnectionRT);

      ExcluirFieldDB('RT_StatusProcessamento','tblProgramacaoExecutante',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('RT_StatusProcessamento','tblProgramacaoRT',
          FrmDataModule.ADOConnectionRT);
      //--------------------------------------------
      CriarFieldDB('RT_StatusAvaliacao', 'tblProgramacaoExecutante', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_StatusProcessamento', 'tblProgramacaoExecutante', 'VARCHAR(50)',
        FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_StatusSAPCodigo', 'tblProgramacaoExecutante', 'VARCHAR(4)',
        FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_StatusSAPDescricao', 'tblProgramacaoExecutante', 'VARCHAR(50)',
        FrmDataModule.ADOQueryTemporarioDBColibri);
      //--------------------------------------------
      CriarFieldDB('RT_StatusAvaliacao', 'tblProgramacaoRT', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_StatusProcessamento', 'tblProgramacaoRT', 'VARCHAR(50)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_StatusSAPCodigo', 'tblProgramacaoRT', 'VARCHAR(4)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_StatusSAPDescricao', 'tblProgramacaoRT', 'VARCHAR(50)',
        FrmDataModule.ADOQueryTemporarioRT);
      //--------------------------------------------
      CriarFieldDB('RT_Orfa', 'tblProgramacaoRT', 'BIT',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_PendenteCancelamento', 'tblProgramacaoRT', 'BIT',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_MotivoConciliacao', 'tblProgramacaoRT', 'VARCHAR(255)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_DataUltimaConciliacao', 'tblProgramacaoRT', 'DATETIME',
        FrmDataModule.ADOQueryTemporarioRT);
      //--------------------------------------------
      CriarFieldDB('RT_Orfa', 'tblRTSapImport', 'BIT',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_PendenteCancelamento', 'tblRTSapImport', 'BIT',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_MotivoConciliacao', 'tblRTSapImport', 'VARCHAR(255)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_DataUltimaConciliacao', 'tblRTSapImport', 'DATETIME',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('Observacao', 'tblRTSapImport', 'VARCHAR(255)',
        FrmDataModule.ADOQueryTemporarioRT);

      //---------------------------------------------
      // ======================================================
      // CONTROLE DE TRANSBORDO - EXECUTANTE
      // ======================================================
      CriarFieldDB('RT_Transbordo', 'tblProgramacaoExecutante', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_TransbordoAereo', 'tblProgramacaoExecutante', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_GrupoTransbordo', 'tblProgramacaoExecutante', 'VARCHAR(80)',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_BaseRetornoTransbordo', 'tblProgramacaoExecutante', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_PrimeiraOrigemTransbordo', 'tblProgramacaoExecutante', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_PrimeiroDestinoTransbordo', 'tblProgramacaoExecutante', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('RT_SeqTransbordo', 'tblProgramacaoExecutante', 'INTEGER',
        FrmDataModule.ADOQueryTemporarioDBColibri);

      // ======================================================
      // CONTROLE DE TRANSBORDO - TABELA RT
      // ======================================================
      CriarFieldDB('RT_Transbordo', 'tblProgramacaoRT', 'YESNO',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_TransbordoAereo', 'tblProgramacaoRT', 'YESNO',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_GrupoTransbordo', 'tblProgramacaoRT', 'VARCHAR(80)',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_BaseRetornoTransbordo', 'tblProgramacaoRT', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_PrimeiraOrigemTransbordo', 'tblProgramacaoRT', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_PrimeiroDestinoTransbordo', 'tblProgramacaoRT', 'VARCHAR(20)',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('RT_SeqTransbordo', 'tblProgramacaoRT', 'INTEGER',
        FrmDataModule.ADOQueryTemporarioRT);

      //=============================================
      ExcluirFieldDB('tblProgramacaoExecutante','TipoEmbarque',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','MotivoNaoExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','PalavraChaveNaoExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','AvaliadoPorExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','DataAvaliacaoExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','ComputadorExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','StatusExecucao',
          FrmDataModule.ADOConnectionColibri);


      ExcluirFieldDB('tblProgramacaoDiaria','booleanIP19',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioSaldo',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioDemora',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioNOTA',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','PorcentoExecucao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioSaida',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioPT',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoDiaria','HorarioEncerramento',
          FrmDataModule.ADOConnectionColibri);

      versaoDB:= '1.7.0.2';
    end;
    if (versaoDB = '1.7.0.2') then
    begin
      CriarFieldDB('SituacaoSurfer', 'tblPlataforma', 'VARCHAR(3)',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoSOV', 'tblPlataforma', 'VARCHAR(3)',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('SituacaoAqua', 'tblPlataforma', 'VARCHAR(5)',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);

      CriarFieldDB('HoraSaidaOrigem', 'tblPlataforma', 'VARCHAR(5)',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);

      ExcluirFieldDB('tblPlataforma','SituacaoAcesso',
          FrmDataModule.ADOConnectionConsulta);

      CriarFieldDB('Distribuicao', 'tblEmbarcacao', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('UsaBridgeMesmoGrupo', 'tblEmbarcacao', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('OrigemBridge', 'tblEmbarcacao', 'VARCHAR(25)',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);


      versaoDB:= '1.7.0.3';
    end;
    if (versaoDB = '1.7.0.3') then
    begin
      ExcluirFieldDB('tblProgramacaoRT','RT_Erro',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusProcessamento',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusAvaliacao',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusSAPCodigo',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusSAPDescricao',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusSAP',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_StatusSAPDesc',
          FrmDataModule.ADOConnectionRT);
      ExcluirFieldDB('tblProgramacaoRT','RT_Mensagem',
          FrmDataModule.ADOConnectionRT);
      CriarFieldDB('RT_Mensagem','tblProgramacaoRT','VARCHAR(100)',
          FrmDataModule.ADOQueryTemporarioRT);



      CriarFieldDB('RT_Status','tblProgramacaoRT','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('RT_Status','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);

      ExcluirFieldDB('tblProgramacaoExecutante','RT_Erro',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusProcessamento',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusAvaliacao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusSAPCodigo',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusSAPDescricao',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusSAP',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_StatusSAPDesc',
          FrmDataModule.ADOConnectionColibri);
      ExcluirFieldDB('tblProgramacaoExecutante','RT_Mensagem',
          FrmDataModule.ADOConnectionColibri);
      CriarFieldDB('RT_Mensagem','tblProgramacaoExecutante','VARCHAR(100)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      versaoDB:= '1.7.0.4';
    end;
    if (versaoDB = '1.7.0.4') then
    begin
      // -----------------------------------------------------------------------
      // Distribuição logística automática via solver
      // -----------------------------------------------------------------------

      // tblEmbarcacao (dbConsulta) — suporte ao solver
      CriarFieldDB('Distribuicao',       'tblEmbarcacao', 'YESNO',       FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('TipoEmbarcacao',     'tblEmbarcacao', 'VARCHAR(20)', FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('VelocidadeNos',      'tblEmbarcacao', 'DOUBLE',      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('CapacidadeMaxSolver','tblEmbarcacao', 'INTEGER',     FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('HorarioLimiteSolar', 'tblEmbarcacao', 'VARCHAR(5)',  FrmDataModule.ADOQueryTemporarioDBConsulta1);

      // tblPlataforma (dbConsulta) — suporte ao solver
      CriarFieldDB('CodigoNormSolver',       'tblPlataforma', 'VARCHAR(10)', FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('GrupoFisico',            'tblPlataforma', 'VARCHAR(30)', FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('GrupoHorario',           'tblPlataforma', 'VARCHAR(5)',  FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('PrioridadeDistribuicao', 'tblPlataforma', 'INTEGER',     FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('ExcluirDistribuicaoAuto','tblPlataforma', 'YESNO',       FrmDataModule.ADOQueryTemporarioDBConsulta1);
      // Classifica o papel logístico do nó na rede de distribuição:
      //   'TERMINAL'   = nó terrestre de apoio (ex: TMIB)
      //   'HUB'        = plataforma offshore que concentra/distribui PAX (ex: PCM-09)
      //   '' ou vazio  = plataforma offshore regular
      CriarFieldDB('TipoNoLogistico',        'tblPlataforma', 'VARCHAR(20)', FrmDataModule.ADOQueryTemporarioDBConsulta1);

      // tblRoteamento (bdColibri) — rastreabilidade da geração automática
      CriarFieldDB('OrigemDistribuicao',      'tblRoteamento', 'VARCHAR(10)', FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('GeradaPorSolver',         'tblRoteamento', 'YESNO',       FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('IdOperacaoDistribuicao',  'tblRoteamento', 'INTEGER',     FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('GrupoHorarioRota',        'tblRoteamento', 'VARCHAR(5)',  FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('DistanciaTotalNM',        'tblRoteamento', 'DOUBLE',      FrmDataModule.ADOQueryTemporarioDBColibri);

      // tblOperacaoDistribuicao (bdColibri) — registro de cada execução do solver
      CriarTableDB('tblOperacaoDistribuicao', 'idOperacaoDistribuicao',
        '[DataOperacao] DATETIME, ' +
        '[Perfil] VARCHAR(20), ' +
        '[Versao] INTEGER, ' +
        '[DataHoraExecucao] DATETIME, ' +
        '[UsuarioExecucao] VARCHAR(50), ' +
        '[StatusExecucao] VARCHAR(20), ' +
        '[MensagemStatus] MEMO, ' +
        '[TrocaTurma] YESNO, ' +
        '[RendidosM9] INTEGER, ' +
        '[ArquivoInput] VARCHAR(255), ' +
        '[ArquivoOutput] VARCHAR(255)',
        FrmDataModule.ADOConnectionColibri);

      // tblDistribuicaoRota (bdColibri) — uma rota por embarcação por operação
      CriarTableDB('tblDistribuicaoRota', 'idDistribuicaoRota',
        '[idOperacaoDistribuicao] INTEGER, ' +
        '[idRoteamento] INTEGER, ' +
        '[NomeEmbarcacao] VARCHAR(50), ' +
        '[HoraPartida] VARCHAR(5), ' +
        '[SequenciaSolver] VARCHAR(255), ' +
        '[DistanciaNM] DOUBLE, ' +
        '[UsaHubM9] YESNO, ' +
        '[PaxTMIB] INTEGER, ' +
        '[PaxM9] INTEGER, ' +
        '[RotaFixa] YESNO',
        FrmDataModule.ADOConnectionColibri);

      // tblDistribuicaoPaxAlocado (bdColibri) — vínculo executante ↔ rota do solver
      CriarTableDB('tblDistribuicaoPaxAlocado', 'idDistribuicaoPaxAlocado',
        '[idDistribuicaoRota] INTEGER, ' +
        '[idProgramacaoExecutante] INTEGER, ' +
        '[OrigemSolver] VARCHAR(10), ' +
        '[PosicaoNaSequencia] INTEGER, ' +
        '[DataHoraChegadaEstimada] DATETIME',
        FrmDataModule.ADOConnectionColibri);

      // tblDistribuicaoExclusao (bdColibri) — executantes excluídos da distribuição e motivo
      CriarTableDB('tblDistribuicaoExclusao', 'idDistribuicaoExclusao',
        '[idOperacaoDistribuicao] INTEGER, ' +
        '[idProgramacaoExecutante] INTEGER, ' +
        '[MotivoExclusao] VARCHAR(50), ' +
        '[DescricaoMotivo] VARCHAR(255)',
        FrmDataModule.ADOConnectionColibri);

      // tblDistribuicaoConfig (bdColibri) — parâmetros configuráveis do solver
      CriarTableDB('tblDistribuicaoConfig', 'idDistribuicaoConfig',
        '[ChaveConfig] VARCHAR(50), ' +
        '[ValorConfig] VARCHAR(255), ' +
        '[TipoValor] VARCHAR(10), ' +
        '[Descricao] VARCHAR(255)',
        FrmDataModule.ADOConnectionColibri);

      // Valores padrão da configuração do solver
      try
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''DEFAULT_SPEED_KN'', ''14.0'', ''FLOAT'', ' +
          '''Velocidade padrao das embarcacoes em nos quando nao configurada individualmente'')');
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''AQUA_APPROACH_TIME'', ''25'', ''INT'', ' +
          '''Minutos extras de aproximacao por parada para embarcacao do tipo AQUA'')');
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''PRIORITY1_PRECEDENCE_PENALTY_NM'', ''250.0'', ''FLOAT'', ' +
          '''Penalidade em NM por ordenar destino Prioridade 1 apos destinos sem prioridade'')');
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''PASTA_REDE_RAIZ'', '''', ''TEXT'', ' +
          '''Pasta raiz de rede para salvar arquivos de entrada e saida da distribuicao'')');
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''PYTHON_EXE_PATH'', ''python.exe'', ''TEXT'', ' +
          '''Caminho do executavel Python usado para rodar o solver.py'')');
        FrmDataModule.ADOConnectionColibri.Execute(
          'INSERT INTO tblDistribuicaoConfig (ChaveConfig, ValorConfig, TipoValor, Descricao) VALUES ' +
          '(''SOLVER_WORK_PATH'', '''', ''TEXT'', ' +
          '''Pasta de trabalho para arquivos temporarios do solver (input/output)'')');
      except
        // Silencia erros de insert duplicado em re-execução da conversão
      end;

      versaoDB := '1.7.0.5';
    end;
    if (versaoDB = '1.7.0.5') then
    begin
      // -----------------------------------------------------------------------
      // Campos booleanos de distribuição em tblPlataforma (dbConsulta)
      // -----------------------------------------------------------------------

      // Hub principal: identifica a plataforma que concentra PAX de outras origens
      // (ex: PCM-09). Apenas um hub pode estar ativo por vez (regra no BeforePost).
      CriarFieldDB('booleanHubPrincipal', 'tblPlataforma', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);

      // Gangway Aqua: plataforma acessível por embarcações do tipo AQUA Helix
      CriarFieldDB('booleanGangwayAqua', 'tblPlataforma', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);

      // Gangway SOV: plataforma acessível por embarcações SOV/Bridge
      CriarFieldDB('booleanGangwaySOV', 'tblPlataforma', 'YESNO',
        FrmDataModule.ADOQueryTemporarioDBConsulta1);

      // -----------------------------------------------------------------------
      // tblParObrigatorio (dbConsulta) — pares de plataformas que obrigatoriamente
      // devem ser atendidas pela mesma embarcação em uma distribuição.
      // Ex: PCM-02 + PCM-03 (M2+M3, muito próximas)
      //     PCM-06 + PCB-01 (M6+B1, muito próximas)
      // -----------------------------------------------------------------------
      CriarTableDB('tblParObrigatorio', 'idParObrigatorio',
        '[PlataformaA] VARCHAR(50), ' +
        '[PlataformaB] VARCHAR(50), ' +
        '[Ativo] YESNO',
        FrmDataModule.ADOConnectionConsulta);

      // Pares padrão baseados na operação offshore atual
      try
        FrmDataModule.ADOConnectionConsulta.Execute(
          'INSERT INTO tblParObrigatorio (PlataformaA, PlataformaB, Ativo) VALUES ' +
          '(''PCM-02'', ''PCM-03'', True)');
        FrmDataModule.ADOConnectionConsulta.Execute(
          'INSERT INTO tblParObrigatorio (PlataformaA, PlataformaB, Ativo) VALUES ' +
          '(''PCM-06'', ''PCB-01'', True)');
      except
        // Silencia erros de insert duplicado em re-execução
      end;

      versaoDB := '1.7.0.6';
    end;
    if (versaoDB = '1.7.0.6') then
    begin
      CriarTableDB(
        'tblManifestoEmbarque',
        'idManifestoEmbarque',
        '[NomeExecutante] VARCHAR(255), '+
        '[CodigoSAP] VARCHAR(10), '+
        '[CPF] VARCHAR(14), '+
        '[OutroDocumento] VARCHAR(25), '+
        '[Origem] VARCHAR(20), '+
        '[Destino] VARCHAR(20), '+
        '[CentroCusto] VARCHAR(20), '+
        '[DiagramaRede] VARCHAR(20), '+
        '[OperRede] VARCHAR(4), '+
        '[ElementoPEP] VARCHAR(20), '+
        '[TipoEmbarque] VARCHAR(20), '+
        '[DataEmbarque] DATETIME',
        FrmDataModule.ADOConnectionRT
      );
      CriarFieldDB('txtTipoEtapaServico','tblManifestoEmbarque','VARCHAR(100)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('txtEmpresa','tblManifestoEmbarque','VARCHAR(150)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('Servico','tblManifestoEmbarque','VARCHAR(255)',
        FrmDataModule.ADOQueryTemporarioRT);

      CriarFieldDB('txtFuncao','tblManifestoEmbarque','VARCHAR(50)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('EmbarqueDesembarque','tblManifestoEmbarque','VARCHAR(15)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('HorarioSaida','tblManifestoEmbarque','VARCHAR(10)',
        FrmDataModule.ADOQueryTemporarioRT);
      CriarFieldDB('TransporteTerrestre','tblManifestoEmbarque','YESNO',
        FrmDataModule.ADOQueryTemporarioRT);



      versaoDB := '1.7.0.7';
    end;
    if (versaoDB = '1.7.0.7') then
    begin
      CriarTabelasProntidao;


      versaoDB := '1.7.0.7';
    end;

    FrmDataModule.ADOQueryColibri.Edit;
    FrmDataModule.DataSourceColibri.DataSet.FieldByName('versao').AsString:= versaoPRG;
    FrmDataModule.ADOQueryColibri.Post;
    MessageBox(0,'Conversão realizada com sucesso!','Colibri',MB_ICONEXCLAMATION);
  except
    MessageBox(0,'Ocorreu algum erro e a operação foi cancelada!','Colibri',MB_ICONERROR);
  end;
end;

procedure TFrmPrincipal.CriarTabelasProntidao;
begin
  // ============================================================
  // 1. CAMPANHAS DE PRONTIDÃO
  // ============================================================

  CriarTableDB(
    'tblProntidaoCampanha',
    'idCampanha',
    '[NomeCampanha] VARCHAR(150), ' +
    '[PlataformaBase] VARCHAR(50), ' +
    '[DataInicioCampanha] DATETIME, ' +
    '[DataFimCampanha] DATETIME, ' +
    '[StatusCampanha] VARCHAR(50), ' +
    '[Observacao] MEMO, ' +
    '[DataCadastro] DATETIME, ' +
    '[UsuarioCadastro] VARCHAR(100), ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  CriarTableDB(
    'tblProntidaoPremissaCampanha',
    'idPremissa',
    '[idCampanha] LONG, ' +
    '[DataInicioVigencia] DATETIME, ' +
    '[DataFimVigencia] DATETIME, ' +
    '[DescricaoPremissa] VARCHAR(150), ' +
    '[QtdeTurnosPrevistos] LONG, ' +
    '[POBMaximoPorTurno] LONG, ' +
    '[POBDiretoReferencia] LONG, ' +
    '[POBIndiretoReferencia] LONG, ' +
    '[POBTotalReferencia] LONG, ' +
    '[HEPPorTurno] DOUBLE, ' +
    '[IntervaloRefeicaoHoras] DOUBLE, ' +
    '[FontePremissa] VARCHAR(100), ' +
    '[Observacao] MEMO, ' +
    '[DataCadastro] DATETIME, ' +
    '[UsuarioCadastro] VARCHAR(100), ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 2. IMPORTAÇÃO DO WHATSAPP
  // ============================================================

  CriarTableDB(
    'tblWAImportacao',
    'idImportacao',
    '[idCampanha] LONG, ' +
    '[ArquivoOrigem] VARCHAR(255), ' +
    '[CaminhoArquivo] VARCHAR(255), ' +
    '[GrupoWhatsApp] VARCHAR(150), ' +
    '[DataImportacao] DATETIME, ' +
    '[UsuarioImportacao] VARCHAR(100), ' +
    '[HashArquivo] VARCHAR(100), ' +
    '[QtdMensagensImportadas] LONG, ' +
    '[QtdMensagensOperacionais] LONG, ' +
    '[StatusImportacao] VARCHAR(50), ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  CriarTableDB(
    'tblWAMensagem',
    'idMensagem',
    '[idImportacao] LONG, ' +
    '[DataHoraMensagem] DATETIME, ' +
    '[Remetente] VARCHAR(150), ' +
    '[TextoOriginal] MEMO, ' +
    '[TextoNormalizado] MEMO, ' +
    '[TipoMensagem] VARCHAR(50), ' +
    '[Processada] YESNO, ' +
    '[Ignorada] YESNO, ' +
    '[HashMensagem] VARCHAR(100), ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 3. PRODUTIVIDADE POR TURNO
  // Uma linha por campanha + data + turno + origem + plataforma serviço
  // ============================================================

  CriarTableDB(
    'tblProntidaoProdutividadeTurno',
    'idProdutividadeTurno',
    '[idCampanha] LONG, ' +
    '[idPremissa] LONG, ' +
    '[idImportacao] LONG, ' +
    '[idMensagemInicio] LONG, ' +
    '[idMensagemFechamento] LONG, ' +

    '[DataReferencia] DATETIME, ' +
    '[Turno] VARCHAR(30), ' +                         // DIURNO / NOTURNO
    '[PlataformaBase] TEXT(50), ' +                // plataforma da campanha/grupo
    '[PlataformaServico] VARCHAR(50), ' +             // onde o serviço ocorreu

    '[OrigemTipo] VARCHAR(50), ' +                    // Plataforma / Embarcacao / Terminal / Outro
    '[OrigemCodigo] VARCHAR(100), ' +                 // PCM-01 / SOV / TMIB etc.
    '[OrigemDescricao] VARCHAR(150), ' +

    '[EmbarcacaoTipo] VARCHAR(50), ' +                // SOV / SURFER / AQUA HELIX
    '[EmbarcacaoNome] VARCHAR(100), ' +               // SURF 1930 / P5 / NORWIND GALE etc.
    '[TipoOperacao] VARCHAR(50), ' +                  // Prontidao / Apoio / Deslocamento

    '[HoraChegadaPlataforma] DATETIME, ' +
    '[HoraInicioAtividade] DATETIME, ' +
    '[HoraEncerramentoAtividade] DATETIME, ' +
    '[HoraSaidaPlataforma] DATETIME, ' +

    '[PermanenciaHoras] DOUBLE, ' +
    '[HoraEfetivaProjetada] DOUBLE, ' +            // HEP
    '[HoraEfetivaRealizada] DOUBLE, ' +            // HER
    '[MediaHoraEfetivaRealizada] DOUBLE, ' +       // MHER, se desejar gravar
    '[PerdasPNP] DOUBLE, ' +                       // paradas não programadas
    '[PerdasPOBReduzido] DOUBLE, ' +               // PPR
    '[HoraUtilReal] DOUBLE, ' +                    // HUR

    '[QtdPOBTotalInformado] LONG, ' +
    '[QtdPOBDiretoInformado] LONG, ' +
    '[QtdPOBIndiretoInformado] LONG, ' +
    '[DiferencaPOBDireto] LONG, ' +
    '[DeficitPOBDireto] LONG, ' +

    '[StatusExtracao] VARCHAR(50), ' +                // Pendente / Extraido / Parcial / Erro
    '[StatusRevisao] VARCHAR(50), ' +                 // Pendente / Revisado / Consolidado
    '[ConfiancaExtracao] DOUBLE, ' +
    '[Revisado] YESNO, ' +
    '[TextoFonteResumo] MEMO, ' +
    '[ObservacaoRevisao] MEMO, ' +
    '[DataCadastro] DATETIME, ' +
    '[UsuarioCadastro] VARCHAR(100)',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 4. EQUIPE INFORMADA NO WHATSAPP POR TURNO
  // ============================================================

  CriarTableDB(
    'tblProntidaoEquipeTurno',
    'idEquipeTurno',
    '[idProdutividadeTurno] LONG, ' +
    '[idFuncaoProdutividade] LONG, ' +
    '[FuncaoOriginalWhatsApp] VARCHAR(150), ' +
    '[FuncaoNormalizada] VARCHAR(150), ' +
    '[TipoEtapaServico] VARCHAR(150), ' +
    '[QuantidadeInformada] LONG, ' +
    '[EhIndireto] YESNO, ' +
    '[ContaComoPOBDireto] YESNO, ' +
    '[OrigemInformacao] VARCHAR(50), ' +              // WhatsApp / Manual / Colibri
    '[TextoOriginalLinha] MEMO, ' +
    '[ConfiancaExtracao] DOUBLE, ' +
    '[Revisado] YESNO, ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 5. DESLOCAMENTO DE EFETIVO PARA OUTRAS PLATAFORMAS
  // ============================================================

  CriarTableDB(
    'tblProntidaoDeslocamentoEfetivo',
    'idDeslocamento',
    '[idProdutividadeTurnoOrigem] LONG, ' +
    '[idCampanha] LONG, ' +
    '[DataReferencia] DATETIME, ' +
    '[Turno] VARCHAR(30), ' +

    '[PlataformaBase] VARCHAR(50), ' +
    '[PlataformaDestino] VARCHAR(50), ' +

    '[OrigemDeslocamentoTipo] VARCHAR(50), ' +
    '[OrigemDeslocamentoCodigo] VARCHAR(100), ' +

    '[EmbarcacaoTipo] VARCHAR(50), ' +
    '[EmbarcacaoNome] VARCHAR(100), ' +

    '[HoraSaidaOrigem] DATETIME, ' +
    '[HoraChegadaDestino] DATETIME, ' +
    '[HoraSaidaDestino] DATETIME, ' +
    '[HoraRetornoOrigem] DATETIME, ' +

    '[MotivoDeslocamento] MEMO, ' +
    '[TextoFonte] MEMO, ' +
    '[Revisado] YESNO, ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  CriarTableDB(
    'tblProntidaoDeslocamentoEquipe',
    'idDeslocamentoEquipe',
    '[idDeslocamento] LONG, ' +
    '[idFuncaoProdutividade] LONG, ' +
    '[FuncaoOriginalWhatsApp] VARCHAR(150), ' +
    '[FuncaoNormalizada] VARCHAR(150), ' +
    '[TipoEtapaServico] VARCHAR(150), ' +
    '[QuantidadeDeslocada] LONG, ' +
    '[TextoOriginalLinha] MEMO, ' +
    '[ConfiancaExtracao] DOUBLE, ' +
    '[Revisado] YESNO, ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 6. DESVIOS, TEMPOS MORTOS E TEMPOS INTRA-EFETIVA
  // ============================================================

  CriarTableDB(
    'tblProntidaoDesvioTurno',
    'idDesvioTurno',
    '[idProdutividadeTurno] LONG, ' +
    '[idTipoDesvio] LONG, ' +
    '[TipoDesvio] VARCHAR(100), ' +
    '[CategoriaDesvio] VARCHAR(100), ' +
    '[DescricaoDesvio] MEMO, ' +
    '[TempoMortoHoras] DOUBLE, ' +
    '[TempoIntraEfetivaHoras] DOUBLE, ' +
    '[AfetaHoraEfetiva] YESNO, ' +
    '[AfetaEficiencia] YESNO, ' +
    '[TextoFonte] MEMO, ' +
    '[ConfiancaExtracao] DOUBLE, ' +
    '[Revisado] YESNO, ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

  CriarTableDB(
    'tblTipoDesvioProdutividade',
    'idTipoDesvio',
    '[TipoDesvio] VARCHAR(100), ' +
    '[CategoriaDesvio] VARCHAR(100), ' +
    '[AfetaHoraEfetivaPadrao] YESNO, ' +
    '[AfetaEficienciaPadrao] YESNO, ' +
    '[PalavrasChave] MEMO, ' +
    '[Observacao] MEMO, ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 7. CADASTRO DE FUNÇÕES / ETAPAS DE SERVIÇO
  // ============================================================

  CriarTableDB(
    'tblFuncaoProdutividade',
    'idFuncaoProdutividade',
    '[FuncaoNormalizada] VARCHAR(150), ' +
    '[TipoEtapaServico] VARCHAR(150), ' +
    '[EhIndireto] YESNO, ' +
    '[ContaComoPOBDireto] YESNO, ' +
    '[ContaNoPOBTotal] YESNO, ' +
    '[Observacao] MEMO, ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  CriarTableDB(
    'tblAliasFuncaoProdutividade',
    'idAliasFuncao',
    '[TextoAlias] VARCHAR(150), ' +
    '[TextoAliasNormalizado] VARCHAR(150), ' +
    '[idFuncaoProdutividade] LONG, ' +
    '[Prioridade] LONG, ' +
    '[Observacao] MEMO, ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 8. DE/PARA COM A PROGRAMAÇÃO DO COLIBRI
  // Segundo momento: comparação com tblProgramacaoExecutante
  // ============================================================

  CriarTableDB(
    'tblDeParaFuncaoColibriProdutividade',
    'idDePara',
    '[txtFuncaoColibri] VARCHAR(150), ' +
    '[txtFuncaoColibriNormalizada] VARCHAR(150), ' +
    '[TipoEtapaServico] VARCHAR(150), ' +
    '[idFuncaoProdutividade] LONG, ' +
    '[Observacao] MEMO, ' +
    '[Ativo] YESNO',
    FrmDataModule.ADOConnectionProntidao
  );

  // ============================================================
  // 9. CONFRONTO CONSOLIDADO ENTRE COLIBRI E WHATSAPP
  // Pode ser tabela materializada para auditoria histórica.
  // Também pode ser gerada inicialmente via query.
  // ============================================================

  CriarTableDB(
    'tblProntidaoConfrontoEquipe',
    'idConfrontoEquipe',
    '[idProdutividadeTurno] LONG, ' +
    '[idCampanha] LONG, ' +
    '[DataReferencia] DATETIME, ' +
    '[Turno] VARCHAR(30), ' +
    '[PlataformaBase] VARCHAR(50), ' +
    '[PlataformaServico] VARCHAR(50), ' +
    '[TipoEtapaServico] VARCHAR(150), ' +
    '[idFuncaoProdutividade] LONG, ' +
    '[FuncaoNormalizada] VARCHAR(150), ' +
    '[QtdProgramadaColibri] LONG, ' +
    '[QtdInformadaWhatsApp] LONG, ' +
    '[QtdDeslocada] LONG, ' +
    '[QtdOficialConsolidada] LONG, ' +
    '[DiferencaProgramadoInformado] LONG, ' +
    '[StatusConfronto] VARCHAR(50), ' +               // OK / Divergente / Extra / Ausente / Deslocado
    '[FonteOficial] VARCHAR(50), ' +                  // Colibri / WhatsApp / Revisao Manual
    '[Revisado] YESNO, ' +
    '[Observacao] MEMO',
    FrmDataModule.ADOConnectionProntidao
  );

end;

function TFrmPrincipal.FormatarCPF(CPF: String): String;
begin
  CPF := SomenteNumero(CPF);

  if Length(CPF) = 11 then
    Result := Copy(CPF, 1, 3) + '.' +
              Copy(CPF, 4, 3) + '.' +
              Copy(CPF, 7, 3) + '-' +
              Copy(CPF, 10, 2)
  else
    Result := CPF;
end;

procedure TFrmPrincipal.actMatrizExecutanteCadastroExecute(Sender: TObject);
  var
    NumRegistros,i: Integer;
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblExecutante.* FROM tblExecutante '+
  'WHERE ((Ativo = True)) ORDER BY NomeExecutante,txtTipoEtapaServico;';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  NumRegistros:= FrmDataModule.ADOQueryTemporarioDBConsulta1.RecordCount;
  //=========================================================================
  MatrizExecutanteCadastro:= nil;
  SetLength(MatrizExecutanteCadastro, 15);
  for i := 0 to High(MatrizExecutanteCadastro) do
    SetLength(MatrizExecutanteCadastro[i], NumRegistros);

  for i := 0 to High(MatrizExecutanteCadastro[0]) do
  begin
    MatrizExecutanteCadastro[0,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('Turma').AsString;
    MatrizExecutanteCadastro[1,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('NomeExecutante').AsString;
    MatrizExecutanteCadastro[2,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtTipoEtapaServico').AsString;
    MatrizExecutanteCadastro[3,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtFuncao').AsString;
    MatrizExecutanteCadastro[4,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('txtEmpresa').AsString;
    MatrizExecutanteCadastro[5,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('CodigoSAP').AsString;
    MatrizExecutanteCadastro[6,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('OrigemCadastro').AsString;
    MatrizExecutanteCadastro[7,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('Chave').AsString;
    MatrizExecutanteCadastro[8,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('CPF').AsString;
    MatrizExecutanteCadastro[9,i]:= UPPERCASE(FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('RequisitantePT').AsString);
    MatrizExecutanteCadastro[10,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('CentroCusto').AsString;
    MatrizExecutanteCadastro[11,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('DiagramaRede').AsString;
    MatrizExecutanteCadastro[12,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('OperRede').AsString;
    MatrizExecutanteCadastro[13,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('ElementoPEP').AsString;
    MatrizExecutanteCadastro[14,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('OutroDocumento').AsString;


    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
end;

procedure TFrmPrincipal.actMatrizForaOperacaoExecute(Sender: TObject);
  var
    NumRegistros,i: Integer;
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblPlataforma.* FROM tblPlataforma '+
  'WHERE (BooleanPlataforma = True) ORDER BY Plataforma;';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  NumRegistros:= FrmDataModule.ADOQueryTemporarioDBConsulta1.RecordCount;
  //=========================================================================
  MatrizForaOperacao:= nil;
  SetLength(MatrizForaOperacao, 14);
  for i := 0 to High(MatrizForaOperacao) do
    SetLength(MatrizForaOperacao[i], NumRegistros);

  for i := 0 to High(MatrizForaOperacao[0]) do
  begin
    MatrizForaOperacao[0,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('Plataforma').AsString;
    MatrizForaOperacao[1,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoGD').AsString;
    MatrizForaOperacao[2,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('CapPrincipal').AsString;
    MatrizForaOperacao[3,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('CapAuxiliar').AsString;
    MatrizForaOperacao[4,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoBCI').AsString;
    MatrizForaOperacao[5,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoLinhaBCI').AsString;
    MatrizForaOperacao[6,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoAgua').AsString;
    MatrizForaOperacao[7,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoBalsa').AsString;


    MatrizForaOperacao[9,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoDegraus').AsString;
    MatrizForaOperacao[10,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('DataRealizacaoDegraus').AsString;
    MatrizForaOperacao[11,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoSurfer').AsString;
    MatrizForaOperacao[12,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoSOV').AsString;
    MatrizForaOperacao[13,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoAqua').AsString;

    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
end;

procedure TFrmPrincipal.actSalvatagemPlataformaExecute(Sender: TObject);
  var
    NumRegistros,i: Integer;
begin
  FrmDataModule.ADOQueryConsultaPlataforma.Active:= false;
  FrmDataModule.ADOQueryConsultaPlataforma.Active:= true;
  NumRegistros:= FrmDataModule.ADOQueryConsultaPlataforma.RecordCount;
  MatrizSalvatagem:= nil;
  SetLength(MatrizSalvatagem, 2);
  for i := 0 to High(MatrizSalvatagem) do
    SetLength(MatrizSalvatagem[i], NumRegistros);

  for i := 0 to High(MatrizSalvatagem[0]) do
  begin
    MatrizSalvatagem[0,i]:= FrmDataModule.DataSourceConsultaPlataforma.DataSet.FieldByName('Plataforma').AsString;
    MatrizSalvatagem[1,i]:= FrmDataModule.DataSourceConsultaPlataforma.DataSet.FieldByName('Salvatagem').AsString;
    FrmDataModule.ADOQueryConsultaPlataforma.Next;
  end;
end;

procedure TFrmPrincipal.actVerificaVersaoExecute(Sender: TObject);
    var
    testeVersao: Integer;
    versaoPRG,versaoDB,enderecoInstaladorEXE: String;
begin
  //Verificar a versão
  try
    FrmDataModule.ADOQueryColibri.Active:= true;
    versaoPRG:= VersaoExe;
    versaoDB:= FrmDataModule.DataSourceColibri.DataSet.FieldByName('versao').AsString;
    enderecoInstaladorEXE:= FrmDataModule.DataSourceColibri.DataSet.FieldByName('enderecoInstalador').AsString;
    testeVersao:= versaoAtualizada(versaoPRG,versaoDB);
    //criarVariavelTempoExecucao('TAG_DUTO','tblPainelDutos','VARCHAR(255)',FrmDataModule.ADOQuerySIGMGeral);
    case testeVersao of
      0: //Nada
      begin

      end;
      1://intVersaoPRG < intVersaoDB
      begin
        if Application.MessageBox(PChar(
        'A versão do banco de dados é SUPERIOR à do programa instalado. Deseja instalar a nova versão?'+#13+
        '   * Versão do banco de dados: '+versaoDB+#13+
        '   * Versão do programa instalado: '+versaoPRG),
        '.::ATENÇÃO::.',36) = 6 then
        begin
          if System.SysUtils.DirectoryExists(enderecoInstaladorEXE) then
          begin
            FrmTelaAbertura.Close;
            enderecoInstaladorEXE:= ExtractFilePath(Application.ExeName)+'Instalador';
            if POS('\OneDrive - PETROBRAS\',enderecoInstaladorEXE)>0 then
              StringReplace('\OneDrive - PETROBRAS\',enderecoInstaladorEXE,'\',[rfReplaceAll]);

            enderecoInstaladorEXE:=  enderecoInstaladorEXE+'\v'+versaoDB+'.exe';
            ShellExecute(0,'open',PChar(enderecoInstaladorEXE),'','',SW_SHOWNORMAL);
            Application.Terminate;
          end
          else
          begin
            ShowMessage('Arquivo instalador não encontrado no endereço: '+
            enderecoInstaladorEXE+#13+'O programa será fechado!');
            Application.Terminate;
          end;
        end
        else
        begin
          if Application.MessageBox(PChar('Deseja abrir outro banco de dados?'),
          '.::ATENÇÃO::.',36) = 6 then
            actAbrirDB.Execute
          else
            Application.Terminate;
        end;
      end;
      2://intVersaoPRG > intVersaoDB
      begin
        if Application.MessageBox(PChar('Deseja converter para nova versão?'+
          #13+'Versão do programa: '+versaoPRG+
          #13+'Versão do banco de dados: '+versaoDB),
        '.::ATENÇÃO::.',36) = 6 then
          actConverterDB.Execute;
      end
      else
      begin
        MessageBox(0, 'Este banco de dados não é válido e o programado será fechado!',
        'Colibri', MB_ICONERROR);
        Application.Terminate;
      end;
    end;
  except
    MessageBox(0, 'Este banco de dados não é válido, abra outro!',
    'Colibri', MB_ICONERROR);
    actAbrirDB.Execute;
    //Application.Terminate;
  end;
end;

procedure TFrmPrincipal.addColuna(Grid: TDBGrid; Field, Caption, alinha: String;
  numColuna,Largura: Integer; booleanReadOnly: Boolean);
begin
  Grid.columns.add;
  Grid.columns[numColuna].FieldName:= Field;
  Grid.columns[numColuna].Title.Caption:= Caption;
  if alinha = 'CENTER' then
    Grid.columns[numColuna].Alignment:= taCenter
  else if alinha = 'LEFT' then
    Grid.columns[numColuna].Alignment:= taLeftJustify
  else if alinha = 'RIGHT' then
    Grid.columns[numColuna].Alignment:= taRightJustify;

  Grid.columns[numColuna].Title.Alignment := taCenter;
  Grid.columns[numColuna].ReadOnly:= booleanReadOnly;
  Grid.columns[numColuna].Width:= Largura;
end;

procedure TFrmPrincipal.ADOConnection_Reconnect(Conn: TADOConnection);
var
  CS: string;
begin
  if Conn = nil then Exit;
  CS := Conn.ConnectionString;
  Conn.Connected := False;
  Conn.Close;
  Conn.ConnectionString := CS;
  Conn.Open;
  Conn.Connected := True;
end;

procedure TFrmPrincipal.Ajuda1Click(Sender: TObject);
var
  wideChars   : array[0..200] of WideChar;
  myString    : String;
begin
  myString:= ExtractFilePath(Application.ExeName)+'Ajuda\Colibri.chm';
  StringToWideChar(myString, wideChars, 200);
  ShellExecute(0,nil,wideChars ,nil, nil, SW_SHOWMAXIMIZED);
end;

procedure TFrmPrincipal.AlterarTituloColuna(ColunasLayout: TStringGrid; FieldName,
  strTitulo: String);
  var
    i: Integer;
begin
  for I := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[0,i] = FieldName then
    begin
      ColunasLayout.Cells[1,i]:= strTitulo;
    end;
  end;
end;

procedure TFrmPrincipal.AutoSizeDBGrid(Grid: TDBGrid);
const
  DEFBORDER = 10;
var
  temp, n: Integer;
  lmax: array [0..30] of Integer;
begin
  try
    with Grid do
    begin
      Canvas.Font := Font;
      for n := 0 to Columns.Count - 1 do
        //if columns[n].visible then
        lmax[n] := Canvas.TextWidth(Fields[n].FieldName) + DEFBORDER;
      grid.DataSource.DataSet.First;
      while not grid.DataSource.DataSet.EOF do
      begin
        for n := 0 to Columns.Count - 1 do
        begin
          //if columns[n].visible then begin
          temp := Canvas.TextWidth(trim(Columns[n].Field.DisplayText)) + DEFBORDER;
          if temp > lmax[n] then lmax[n] := temp;
          //end; { if }
        end; {for}
        grid.DataSource.DataSet.Next;
      end; { while }
      grid.DataSource.DataSet.First;
      for n := 0 to Columns.Count - 1 do
        if lmax[n] > 0 then
          Columns[n].Width := lmax[n];
    end; { With }
  except
  end;
end;

procedure TFrmPrincipal.AvaliarProgramacaoExecutante(
  idProgramacaoExecutante,Fonte: Integer; StatusProgramacao,
  Motivo: String);
begin
  //Procurar Programação do Executante
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
  idProgramacaoExecutante;
  FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= true;
  //Gravar
  try
    case Fonte of
      0://Cancelamento ou Mudança (Gerenciamento de Solicitações)
      begin
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('StatusProgramacao').AsString:= StatusProgramacao;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('MotivoProgramacao').AsString:= Motivo;
        //=====================================================
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('AvaliadoPorProgramacao').AsString:= logChave;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('ComputadorProgramacao').AsString:= logMaquina;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('DataAvaliacaoProgramacao').AsDateTime:= now;
        //=====================================================
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
      end;
    end;
  except
    ShowMessage('Erro');
  end;
end;

function TFrmPrincipal.buscarString(strBusca, strTexto: String): Boolean;
  var
    i,j: Integer;
    ListaBusca: TStringList;
begin
  ListaBusca:=TStringList.Create;
  ListaBusca.delimiter:= ';';
  ListaBusca.StrictDelimiter:=true;
  ListaBusca.Clear;
  ListaBusca.DelimitedText:= strBusca;
  Result:= false;
  if strBusca = '' then
    Result:= true
  else
  begin
    for j := 0 to ListaBusca.Count-1 do
    begin
      if ListaBusca[j] = '' then
      begin
        Result:= true;
        break
      end
      else
      begin
        i:= Pos(UpperCase(ListaBusca[j]),UpperCase(strTexto));
        if i<>0 then
        begin
          Result:= true;
          break
        end
        else
          Result:= false;
      end;
    end;
  end;
end;

procedure TFrmPrincipal.CadastroUsuario1Click(Sender: TObject);
begin
  if not Assigned(FrmCadastroUsuario) then
    FrmCadastroUsuario:= TFrmCadastroUsuario.Create(Application)
  else
    FrmCadastroUsuario.Show;
end;

function TFrmPrincipal.carregaDataMinima(Servico: Boolean): TDateTime;
begin
  if Servico then
  begin
    if ((logPerfil = PERFILADM)OR
    (logPerfil = PERFILPROGRAMACAO)OR
    (logPerfil = PERFILSUPERVISAO)) then
      DateTimePickerMinima.DateTime:= EncodeDate(2012, 01, 01)
    else
      DateTimePickerMinima.DateTime:= now;//IncDay(now,1);
  end
  else
  begin
    if (logPerfil = PERFILADM) then
      DateTimePickerMinima.DateTime:= EncodeDate(2012, 01, 01)
    else
      DateTimePickerMinima.DateTime:= now;//IncDay(now,1);
  end;

  DateTimePickerMinima.Time:= 0;
  Result:= DateTimePickerMinima.DateTime;
end;

procedure TFrmPrincipal.CarregaFiltrosProcura(ColunasLayout: TStringGrid);

begin

end;

procedure TFrmPrincipal.carregarComboBox(Conection: TADOConnection;
  txtField,sql: String; comboBox: TComboBox);
var
  Query: TADOQuery;
begin
  Query := TADOQuery.Create(self);
  try
    Query.Connection := Conection;
    Query.Close;
    Query.SQL.Clear;
    Query.SQL.Add(sql);
    Query.Open;
    Query.Active:= true;
    Query.First;
    comboBox.Items.Clear;
    comboBox.Text:= Query.FieldByName(txtField).AsString;
    while not (Query.Eof) do
    begin
      comboBox.Items.Add(Query.FieldByName(txtField).AsString);
      Query.Next;
    end;
  except
    showmessage('Erro ao carregar o combobox do campo: '+txtField);
  end;
end;

procedure TFrmPrincipal.carregarLoginUsuario(Chave: String);
  var
    logar: Boolean;
begin
  logChave:= Chave;
  logMaquina:= NomeMaquina;
  //============CARREGAR DADOS DO USUARIO CADASTRADO================
  try
    logar:= true;
    FrmDataModule.ADOQueryUsuario.Active:=false;
    FrmDataModule.ADOQueryUsuario.Active:=true;
  except
    MessageBox(0, 'Não foi possível abrir o banco de dados!' + #13 + #10 +
    'Selecione um banco de dados válido!',
    'Colibri', MB_ICONERROR);
    actAbrirDB.Execute;
    logar:= false;
  end;
  if logar then
  begin
    //Filtrar para procurar usuario na tabela de cadastro interna
    FrmDataModule.ADOQueryUsuario.Filtered:= false;
    FrmDataModule.ADOQueryUsuario.Filter:= 'Chave like '+QuotedStr(logChave);
    FrmDataModule.ADOQueryUsuario.Filtered:=true;
    //Carregar variaveis de acesso
    if FrmDataModule.DataSourceUsuario.DataSet.FieldByName('Chave').AsString<>'' then
      logPerfil:= FrmDataModule.DataSourceUsuario.DataSet.FieldByName('Perfil').AsString
    else
      logPerfil:= 'Consulta';
    //================CONFIGURAR O STATUSBAR======================
    StatusBarPrincipal.Panels[0].Text:= logChave;
    StatusBarPrincipal.Panels[1].Text:= logMaquina;
    StatusBarPrincipal.Panels[2].Text:= DateToStr(now);
    StatusBarPrincipal.Panels[3].Text:= logPerfil;
    AutoFitStatusBar(StatusBarPrincipal);
    //=========================================================
    //===========CARREGAR O PERFIL DE USUARIO==================
    //=========================================================
    if logPerfil = PERFILADM then //Administrador
    begin
      CadastroUsuario1.Enabled:= true;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else if logPerfil = PERFILSUPERVISAO then
    begin
      CadastroUsuario1.Enabled:= true;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else if logPerfil = PERFILPROGRAMACAO then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else if logPerfil = PERFILHOTELARIA then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else if logPerfil = PERFILRT then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= false;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= false;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else //if logPerfil = 'Consulta' then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= false;
      SalvarBancoDadosComo1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= false;
      CompactarBancoDados1.Enabled:= true;
      Converterverso1.Enabled:= true;
    end;
  end;
end;

procedure TFrmPrincipal.Force_Reconnect;
begin
  ADOConnection_Reconnect(FrmDataModule.ADOConnectionColibri);
  ADOConnection_Reconnect(FrmDataModule.ADOConnectionConsulta);
end;

procedure TFrmPrincipal.carregarRadioGroup(ADOQuery: TADOQuery;
  RadioGroup: TRadioGroup; FildName,PrimeiroItem: String);
begin
  ADOQuery.Active:= false;
  ADOQuery.Active:= true;
  RadioGroup.Items.Clear;
  if PrimeiroItem<>'' then
    RadioGroup.Items.Add(PrimeiroItem);
  while not ADOQuery.Eof do
  begin
    RadioGroup.Items.Add(ADOQuery.FieldByName(FildName).AsString);
    ADOQuery.Next;
  end;
end;

procedure TFrmPrincipal.carregarCheckListBox
(ADOQuery: TADOQuery; checkList: TCheckListBox;FildName:String);
begin
  ADOQuery.Active:= false;
  ADOQuery.Active:= true;
  checkList.Items.Clear;
  while not ADOQuery.Eof do
  begin
    checkList.Items.Add(ADOQuery.FieldByName(FildName).AsString);
    ADOQuery.Next;
  end;
end;

function TFrmPrincipal.Clasifica(Grid: TStringGrid; ACol: Integer;indicar: Boolean): String;
  var
    Lista,ListaDesc : TStringlist;
    p, i: integer;
    linha,strTitulo:string;
    booAsc,booDesc: Boolean;
Begin
  Lista := TStringlist.Create;
  ListaDesc := TStringlist.Create;
  Lista.Clear;
  ListaDesc.Clear;
  booAsc:= CharAscDesc(char(9650),Grid.Cells[ACol,0]);
  booDesc:= CharAscDesc(char(9660),Grid.Cells[ACol,0]);
  if indicar then
  begin
    for i := 0 to Grid.ColCount-1  do
    begin
      strTitulo:= DeleteChar(char(9650),Grid.Cells[i,0]);//Asc
      strTitulo:= DeleteChar(char(9660),strTitulo);//Desc
      strTitulo:= DeleteChar(char(8722),strTitulo);
      Grid.Cells[i,0]:= strTitulo;
    end;
  end;
  for i := 1 to Grid.RowCount-1  do
  begin
    if trim(Grid.Rows[i].text)<>'' then
    begin
      Lista.Append(Grid.Cells[ACol,i]+'//limite//'+Grid.Rows[i].Text);
      ListaDesc.Append(Grid.Cells[ACol,i]+'//limite//'+Grid.Rows[i].Text);
    end;
  end;
  Lista.Sort;
  for i := 1 to Grid.RowCount-1  do
  begin
    Grid.Rows[i].Clear;
  end;
  for i := 0 to Lista.Count-1  do
  begin
    linha := Lista.Strings[i];
    p := pos('//limite//',linha)+10;
    linha := copy(linha,p,length(linha));
    Grid.Rows[i+1].Text := linha;
  end;
  if booAsc then
  begin
    for i := 0 to Lista.Count-1  do
    begin
      ListaDesc[i]:= Lista[Lista.Count-1-i];
    end;
    for i := 0 to ListaDesc.Count-1  do
    begin
      linha := ListaDesc.Strings[i];
      p := pos('//limite//',linha)+10;
      linha := copy(linha,p,length(linha));
      Grid.Rows[i+1].Text := linha;
    end;
  end;
  Lista.Destroy;
  ListaDesc.Destroy;
  if indicar then
  begin
    if booAsc then
    begin
      Grid.Cells[ACol,0]:= Grid.Cells[ACol,0] +char(8722)+char(9660);
    end
    else if booDesc then
    begin
      Grid.Cells[ACol,0]:= Grid.Cells[ACol,0] +char(8722)+char(9650);
    end
    else
    begin
      Grid.Cells[ACol,0]:= Grid.Cells[ACol,0] +char(8722)+char(9650);
    end;
  end;
end;

procedure TFrmPrincipal.ClassificaDBGrid(Grid: TDBGrid;sourceADOQuery: TADOQuery; TipoSORT: Integer);
  var
    i,IndexColumn: Integer;
    strTitulo: String;
begin
  Try
    IndexColumn:= StrToInt(MagicIndexColuna);
    //LIMPAR O DBGRID
    for i:=0 to Grid.FieldCount-1 do
    begin
      Grid.Columns.Items[i].Title.Font.Color:=ClBlack;//Cor da Fonte
      Grid.Columns.Items[i].Title.Font.Style := [];//Cot da Fonte
      Grid.Columns.Items[i].Title.Color:=clBtnFace;//Cor do Fundo do Titulo Normal
      Grid.Columns.Items[i].Font.color:= ClBlack;
      strTitulo:= DeleteChar(char(9650),Grid.Columns.Items[i].Title.Caption);
      strTitulo:= DeleteChar(char(9660),strTitulo);
      strTitulo:= DeleteChar(char(8722),strTitulo);
      Grid.Columns.Items[i].Title.Caption:= strTitulo;
    end;
    case TipoSORT of
      0:
      begin
        Grid.Columns.Items[IndexColumn].Title.Color := clSilver;//Cor do Fundo do Titulo Selecionado
        sourceADOQuery.Sort := MagicFieldName+ ' ASC';
        Grid.Columns.Items[IndexColumn].Title.Caption:=
        Grid.Columns.Items[IndexColumn].Title.Caption +char(8722)+char(9650);
      end;
      1:
      begin
        Grid.Columns.Items[IndexColumn].Title.Color := clSilver;//Cor do Fundo do Titulo Selecionado
        sourceADOQuery.Sort := MagicFieldName+ ' DESC';
        Grid.Columns.Items[IndexColumn].Title.Caption:=
        Grid.Columns.Items[IndexColumn].Title.Caption +char(8722)+char(9660);
      end;
    end;
  Except
    ShowMessage('Não foi possível organizar');
  End;
end;

procedure TFrmPrincipal.ComboBoxOperadorKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key:= #0;
end;

procedure TFrmPrincipal.CompactarBancoDados1Click(Sender: TObject);
  var
    caminhoReg: String;
begin
  if Application.MessageBox(PChar(
  'Deseja realmente compactar e reparar todos os bancos de dados?'),
  '.::ATENÇÃO::.',36) = 6 then
  begin
    //Colibri
    caminhoReg := registroEndereco('Banco de dados');
    compactarDB(caminhoReg,false,false,FrmDataModule.ADOConnectionColibri);
    //Consulta
    caminhoReg:= ExtractFilePath(caminhoReg)+'dbConsulta.mdb';
    compactarDB(caminhoReg,false,false,FrmDataModule.ADOConnectionConsulta);
    //RT
    caminhoReg:= ExtractFilePath(caminhoReg)+'dbRT.mdb';
    compactarDB(caminhoReg,false,false,FrmDataModule.ADOConnectionRT);
  end;
end;

procedure TFrmPrincipal.compactarDB(EnderecoArquivo: String;booleanPerguntar,conexaoColibri: Boolean;
SourceADOConnection: TADOConnection);
var
  dao: OleVariant;
  Caminho_Copia,Caminho_Temp: String;
  buttonSelected: Integer;
begin
  // Compactar e Reparar Banco de dados
  try
    if booleanPerguntar then
    begin
      if Application.MessageBox(PChar(
      'Deseja realmente compactar o banco de dados?'
      + #13 + 'Atenção: Este processo pode demorar e dados '+
      'alterados por outros usuários neste período podem ser perdidos!'),
      '.::ATENÇÃO::.',36) = 6 then
        buttonSelected := mrYes
      else
        buttonSelected := mrNo;
    end
    else
      buttonSelected := mrYes;
    //=====================================================================
    if buttonSelected = mrYes then
    begin
      Caminho_Copia := ExtractFilePath(Application.ExeName) + 'dbCopia.mdb';
      Caminho_Temp := ExtractFilePath(Application.ExeName) + 'dbTemporario.mdb';
      DeleteFile(Caminho_Copia);
      DeleteFile(Caminho_Temp);
      if CopyFile(PChar(EnderecoArquivo), PChar(Caminho_Copia), false) then
      begin
        FrmPrincipal.ProgressBarIncializa(29, 'Compactando banco de dados...');
        //Fechar Conexão
        SourceADOConnection.Connected := false;
        FrmPrincipal.ProgressBarIncremento(2);
        SourceADOConnection.Close;
        FrmPrincipal.ProgressBarIncremento(3);
        //Criar banco
        dao := CreateOleObject('DAO.DBEngine.36');
        FrmPrincipal.ProgressBarIncremento(5);
        //Compactar banco
        dao.CompactDatabase(Caminho_Copia, Caminho_Temp);
        FrmPrincipal.ProgressBarIncremento(5);
        DeleteFile(Caminho_Copia);
        FrmPrincipal.ProgressBarIncremento(5);
        CopyFile(PChar(Caminho_Temp), PChar(EnderecoArquivo), false);
        DeleteFile(Caminho_Temp);
        //renameFile(Caminho_Temp,caminhoReg);
        FrmPrincipal.ProgressBarIncremento(3);
        // Menssagem
        if booleanPerguntar then
          MessageBox(0, 'Banco de dados compactado com sucesso!', 'Colibri',
          MB_ICONINFORMATION);
      end;
    end;
  finally
    if booleanPerguntar then
    begin
      try
        if conexaoColibri then
          conectarBD(EnderecoArquivo,SourceADOConnection)
        else
          conectarBDDireto(EnderecoArquivo,SourceADOConnection);

        FrmPrincipal.ProgressBarAtualizar;
      except
        MessageBox(0, 'Não foi possível conectar com o banco de dados e o programa será finalizado!',
        'Colibri', MB_ICONERROR);
        Application.Terminate;
      end;
    end
    else
    begin
      SourceADOConnection.Open;
      SourceADOConnection.Connected := true;
      FrmPrincipal.ProgressBarAtualizar;
    end;
  end;
end;

procedure TFrmPrincipal.CondicaoMarEmbarcacao1Click(Sender: TObject);
begin
  if not Assigned(FrmCondicaoEmbarcacao) then
    FrmCondicaoEmbarcacao:= TFrmCondicaoEmbarcacao.Create(Application)
  else
    FrmCondicaoEmbarcacao.Show;
end;

function TFrmPrincipal.conectarBD(Endereco: String;
  sourceADOConnection: TADOConnection): String;
begin
  try
    OpenDialog1.Filter:= 'Arquivo Colibri|*.colibri';
    OpenDialog1.DefaultExt:= '*.colibri';
    //Carregar banco de dados
    sourceADOConnection.Connected:=false;
    sourceADOConnection.Close;
    sourceADOConnection.ConnectionString :=
    'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+
    Endereco +
    ';Mode=Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
    'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking '+
    'Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet '+
    'OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:'+
    'Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact '+
    'Without Replica Repair=False;Jet OLEDB:SFP=False;';
    sourceADOConnection.Open;
    sourceADOConnection.Connected:=true;
    FrmPrincipal.Caption:= Endereco;//ExtractFileName(caminhoReg);
    actVerificaVersao.Execute;
    Result:= Endereco;
  except
    if Application.MessageBox(PChar(
    'Não foi possível conectar ao banco de dados!' + #13 +
    'Possibilidades: '
    +#13+#10+
    '1- O banco de dados está corrompido!'
    +#13+#10+
    '2- O endereço não existe: '+Endereco
    +#13+#10+#13+
    '* Para realizar uma tentativa de recuperar o banco de dados tecle "SIM"'
    +#13+#10+
    '* Para abrir outro banco de dados tecle "NÃO"'),
    '.::ATENÇÃO::.',36) = 6 then
    begin
      compactarDB(Endereco,false,true,FrmDataModule.ADOConnectionColibri);
    end
    else
    begin
      OpenDialog1.FileName:= Endereco;
      if OpenDialog1.Execute then
      begin
        //Carregar banco de dados
        sourceADOConnection.Connected:=false;
        sourceADOConnection.Close;
        Endereco := OpenDialog1.FileName;
        Result:= Endereco;
        sourceADOConnection.ConnectionString :=
        'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+
        Endereco +
        ';Mode=Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
        'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking '+
        'Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet '+
        'OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:'+
        'Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact '+
        'Without Replica Repair=False;Jet OLEDB:SFP=False;';
        try
          sourceADOConnection.Open;
          sourceADOConnection.Connected:=true;
          FrmPrincipal.Caption:= 'Colibri: '+ Endereco;// ExtractFileName(OpenDialog1.FileName);
          registroEscrever('Banco de dados', Endereco);
          actVerificaVersao.Execute;
        except
          MessageBox(0, 'Não foi possível conectar com o banco de dados e o programa será finalizado!',
          'Colibri', MB_ICONERROR);
          Application.Terminate;
        end;
      end;
    end;
  end;
end;

procedure TFrmPrincipal.conectarBDDireto(Endereco: String;
  sourceADOConnection: TADOConnection);
begin
  //Carregar banco de dados
  sourceADOConnection.Connected:=false;
  sourceADOConnection.Close;
  sourceADOConnection.ConnectionString :=
  'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+
  Endereco +
  ';Mode=Share Deny None;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
  'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking '+
  'Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet '+
  'OLEDB:New Database Password="";Jet OLEDB:Create System Database=False;Jet OLEDB:'+
  'Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;Jet OLEDB:Compact '+
  'Without Replica Repair=False;Jet OLEDB:SFP=False;';
  sourceADOConnection.Open;
  sourceADOConnection.Connected:=true;
end;

procedure TFrmPrincipal.ConexoLOCAL1Click(Sender: TObject);
begin
  if not Assigned(FrmConexaoLOCAL) then
    FrmConexaoLOCAL:= TFrmConexaoLOCAL.Create(Application)
  else
    FrmConexaoLOCAL.Show;
end;

procedure TFrmPrincipal.ConfigGridLayout(Grid: TStringGrid; ACol, ARow: Integer;
  aRect: TRect);
begin
  with Grid do
  begin
    If (ARow = 0)  then
    begin
      Canvas.Font.Color:= clBlack;
      Canvas.Font.Style:= [fsBold];
      Canvas.Brush.Color:= clGradientActiveCaption;
      aRect.Left:= aRect.Left-6;
      Canvas.FillRect(aRect);
      Canvas.TextOut((aRect.Right-Canvas.TextWidth(Cells[ACol, ARow])) div 2,aRect.Top,
      Cells[ACol, ARow]);
    end
    else if ((ARow = Row)AND(ACol = 1)) then
    begin
      Canvas.Brush.Color:= $004080FF;
      aRect.Left:= aRect.Left-4;
      Canvas.FillRect(aRect);
      Canvas.TextOut(aRect.Left+5, aRect.Top,
      Cells[ACol, ARow]);
    end;
  end;
end;

procedure TFrmPrincipal.configurarCheckListBox(CheckListBox: TCheckListBox;
  DBGrid: TDBGrid; Layout: TStringGrid; BooleanLimparColunas: Boolean);
  var
    FieldName,TituloColuna,Condicional,Busca,Descricao,SQLBase: String;
    i,j: Integer;
begin
  //Sistemas
  if BooleanLimparColunas then
  begin
    limparTitleGrid(DBGrid);
    ClearFilterValues(Layout);
  end;
  //=======================================================
  for i := 0 to CheckListBox.Items.Count-1 do
  begin
    if CheckListBox.Checked[i] then
    begin
      Descricao:= CheckListBox.Items.Strings[i];
      SQLBase:= 'SELECT tblFiltroSistemas.* FROM tblFiltroSistemas '+
      'WHERE (Descricao LIKE '+QuotedStr(Descricao)+');';
      FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
      FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
      FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
      FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
      TituloColuna:= FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('TituloColuna').AsString;
      for j := 0 to DBGrid.Columns.Count-1 do
      begin
        if (DBGrid.Columns[j].Title.Caption = TituloColuna) then
        begin
          FieldName:= DBGrid.Columns[j].FieldName;
          Busca:= FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('PalavraBusca').AsString;
          Condicional:= FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('Condicional').AsString;
          SetFilterValue(FieldName,Busca,Condicional,Layout,4,true);
        end;
      end;
    end;
  end;
end;

procedure TFrmPrincipal.ControledeGeradores1Click(Sender: TObject);
begin
  if not Assigned(FrmControleGeradores) then
    FrmControleGeradores:= TFrmControleGeradores.Create(Application)
  else
    FrmControleGeradores.Show;
end;

procedure TFrmPrincipal.AplatanexodePT1Click(Sender: TObject);
begin
  if not Assigned(FrmAplatAnexos) then
    FrmAplatAnexos := TFrmAplatAnexos.Create(Application)
  else
    FrmAplatAnexos.Show;
end;

{procedure TFrmPrincipal.CopiarProgramacao(DataProgramacao: String; DataSource: TDataSource);
  var
    CodigoProgramacaoDiariaREF,CodigoProgramacaoDiariaGravar: Integer;
    LogAcao,txtTipoEtapaServico,txtDestino,NumExecutantes: String;
begin
  FrmDataModule.ADOQueryInserirProgramacao.Active:= false;
  FrmDataModule.ADOQueryInserirProgramacao.Active:= true;
  FrmDataModule.ADOQueryInserirProgramacao.Insert;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('DataProgramacao').AsString:=
  DataProgramacao;
  txtDestino:= DataSource.DataSet.FieldByName('txtDestino').AsString;
  txtTipoEtapaServico:= DataSource.DataSet.FieldByName('txtTipoEtapaServico').AsString;
  NumExecutantes:= DataSource.DataSet.FieldByName('NumExecutantes').AsString;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('txtDestino').AsString:=
  txtDestino;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('txtTipoEtapaServico').AsString:=
  txtTipoEtapaServico;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('NumExecutantes').AsString:=
  NumExecutantes;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('NumCancelados').AsInteger:= 0;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('NumAprovados').AsString:=
  NumExecutantes;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('CriadoPor').AsString:=
  FrmPrincipal.logChave;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('ComputadorCriacao').AsString:=
  FrmPrincipal.logMaquina;
  FrmDataModule.DataSourceInserirProgramacao.DataSet.FieldByName('DataCriacao').AsDateTime:= now;
  //===============================================================================
  MemoPrincipal.Text:= FrmDataModule.DataSourceInserirProgramacao.DataSet.
  FieldByName('LogAcao').AsString;
  LogAcao:= 'Programação: '+DataProgramacao+#9+txtDestino+#9+txtTipoEtapaServico+#9+
  '['+logChave+'; '+DateToStr(now)+'; '+logMaquina+'; Criação]';
  MemoPrincipal.Lines.Add(LogAcao);
  FrmDataModule.DataSourceInserirProgramacao.DataSet.
  FieldByName('LogAcao').AsString:= MemoPrincipal.Text;
  FrmDataModule.ADOQueryInserirProgramacao.Post;
  //===============================================================================
  CodigoProgramacaoDiariaREF:= DataSource.DataSet.FieldByName('idProgramacaoDiaria').AsInteger;
  CodigoProgramacaoDiariaGravar:= FrmDataModule.ADOQueryInserirProgramacao.
  FieldByName('idProgramacaoDiaria').AsInteger;
  //===============================================================================
  FrmPrincipal.DownLoadQuery(FrmDataModule.ADOQueryInserirExecutante1,
  FrmDataModule.DataSourceInserirExecutante,CodigoProgramacaoDiariaREF,
  CodigoProgramacaoDiariaGravar,1,'CodigoProgramacaoDiaria','Executantes',
  'SELECT tblProgramacaoExecutante.* FROM tblProgramacaoExecutante '+
  'WHERE (CodigoProgramacaoDiaria=@CodigoProgramacaoDiaria) '+
  'ORDER BY NomeExecutante;',true,true);
  //===============================================================================
  FrmPrincipal.DownLoadQuery(FrmDataModule.ADOQueryInserirServico,
  FrmDataModule.DataSourceInserirServico,CodigoProgramacaoDiariaREF,
  CodigoProgramacaoDiariaGravar,1,'CodigoProgramacaoDiaria','Serviços',
  'SELECT tblProgramacaoServico.* FROM tblProgramacaoServico '+
  'WHERE (CodigoProgramacaoDiaria=@Servico);',true,false);
  //===============================================================================
end;}

procedure TFrmPrincipal.CopiarProgramacao(const ADataProgramacao: TDateTime;
  ASource: TDataSource);
var
  CodigoProgramacaoDiariaREF: Integer;
  CodigoProgramacaoDiariaGravar: Integer;
  LogAcao: string;
  Destino: string;
  TipoEtapaServico: string;
  NumExecutantes: Integer;
  DS: TDataSet;
begin
  if (ASource = nil) or (ASource.DataSet = nil) or (not ASource.DataSet.Active) then
    raise Exception.Create('Fonte de dados da programação inválida.');

  DS := ASource.DataSet;

  Destino := DS.FieldByName('txtDestino').AsString;
  TipoEtapaServico := DS.FieldByName('txtTipoEtapaServico').AsString;
  NumExecutantes := DS.FieldByName('NumExecutantes').AsInteger;
  CodigoProgramacaoDiariaREF := DS.FieldByName('idProgramacaoDiaria').AsInteger;

  FrmDataModule.ADOConnectionColibri.BeginTrans;
  try
    with FrmDataModule.ADOQueryInserirProgramacao do
    begin
      Close;
      Open;
      Insert;

      FieldByName('DataProgramacao').AsDateTime := Trunc(ADataProgramacao);
      FieldByName('txtDestino').AsString := Destino;
      FieldByName('txtTipoEtapaServico').AsString := TipoEtapaServico;
      FieldByName('NumExecutantes').AsInteger := NumExecutantes;
      FieldByName('NumCancelados').AsInteger := 0;
      FieldByName('NumAprovados').AsInteger := NumExecutantes;
      FieldByName('CriadoPor').AsString := FrmPrincipal.logChave;
      FieldByName('ComputadorCriacao').AsString := FrmPrincipal.logMaquina;
      FieldByName('DataCriacao').AsDateTime := Now;

      LogAcao :=
        'Programação: ' + FormatDateTime('dd/mm/yyyy', ADataProgramacao) + #9 +
        Destino + #9 +
        TipoEtapaServico + #9 +
        '[' + logChave + '; ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) +
        '; ' + logMaquina + '; Criação]';

      FieldByName('LogAcao').AsString := LogAcao;

      Post;

      CodigoProgramacaoDiariaGravar := FieldByName('idProgramacaoDiaria').AsInteger;
    end;

    CopiarRegistrosRelacionados(
      FrmDataModule.ADOQueryInserirExecutante1,
      CodigoProgramacaoDiariaREF,
      CodigoProgramacaoDiariaGravar,
      'CodigoProgramacaoDiaria',
      'SELECT * FROM tblProgramacaoExecutante ' +
      'WHERE CodigoProgramacaoDiaria = :CodigoProgramacaoDiaria ' +
      'ORDER BY NomeExecutante',
      True
    );

    CopiarRegistrosRelacionados(
      FrmDataModule.ADOQueryInserirServico,
      CodigoProgramacaoDiariaREF,
      CodigoProgramacaoDiariaGravar,
      'CodigoProgramacaoDiaria',
      'SELECT * FROM tblProgramacaoServico ' +
      'WHERE CodigoProgramacaoDiaria = :CodigoProgramacaoDiaria',
      False
    );

    LimparCamposRTDaProgramacao(CodigoProgramacaoDiariaGravar);

    FrmDataModule.ADOConnectionColibri.CommitTrans;
  except
    on E: Exception do
    begin
      FrmDataModule.ADOConnectionColibri.RollbackTrans;
      raise Exception.Create('Erro ao copiar programação: ' + E.Message);
    end;
  end;
end;

procedure TFrmPrincipal.CopiarRegistrosRelacionados(
  AGravarQuery: TADOQuery;
  const AIdOrigem, AIdDestino: Integer;
  const ACampoLigacao: string;
  const ASQLConsulta: string;
  const AExecutante: Boolean);
var
  QConsulta: TADOQuery;
  I: Integer;
  NomeCampo: string;
begin
  QConsulta := TADOQuery.Create(nil);
  try
    QConsulta.Connection := FrmDataModule.ADOConnectionColibri;
    QConsulta.Close;
    QConsulta.SQL.Text := ASQLConsulta;
    QConsulta.Parameters.ParamByName('CodigoProgramacaoDiaria').Value := AIdOrigem;
    QConsulta.Open;

    AGravarQuery.Close;
    AGravarQuery.Open;

    while not QConsulta.Eof do
    begin
      AGravarQuery.Insert;

      for I := 0 to QConsulta.FieldCount - 1 do
      begin
        NomeCampo := QConsulta.Fields[I].FieldName;

        if SameText(NomeCampo, 'idProgramacaoExecutante') or
           SameText(NomeCampo, 'idProgramacaoServico') or
           SameText(NomeCampo, 'id') then
          Continue;

        if AGravarQuery.FindField(NomeCampo) <> nil then
          AGravarQuery.FieldByName(NomeCampo).Value :=
            QConsulta.FieldByName(NomeCampo).Value;
      end;

      if AGravarQuery.FindField(ACampoLigacao) <> nil then
        AGravarQuery.FieldByName(ACampoLigacao).AsInteger := AIdDestino;

      if AExecutante then
      begin
        if AGravarQuery.FindField('StatusProgramacao') <> nil then
          AGravarQuery.FieldByName('StatusProgramacao').AsString := 'Aprovado';

        if AGravarQuery.FindField('MotivoProgramacao') <> nil then
          AGravarQuery.FieldByName('MotivoProgramacao').AsString := '';

        if AGravarQuery.FindField('InseridoProgramacaoTransporte') <> nil then
          AGravarQuery.FieldByName('InseridoProgramacaoTransporte').AsBoolean := False;
      end;

      AGravarQuery.Post;
      QConsulta.Next;
    end;

  finally
    QConsulta.Free;
  end;
end;

procedure TFrmPrincipal.LimparCamposRTDaProgramacao(
  const ACodigoProgramacaoDiaria: Integer);
var
  Q: TADOQuery;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;

    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante ' +
      'SET ' +
      '  RT = NULL, ' +
      '  RT_Transbordo = 0, ' +
      '  RT_TransbordoAereo = 0, ' +
      '  booleanRecolhimento = 0, ' +
      '  RT_HoraIda = NULL, ' +
      '  RT_HoraVolta = NULL, ' +
      '  RT_GrupoTransbordo = NULL, ' +
      '  RT_PrimeiraOrigemTransbordo = NULL, ' +
      '  RT_PrimeiroDestinoTransbordo = NULL, ' +
      '  RT_SeqTransbordo = NULL, ' +
      '  RT_Status = NULL ' +
      'WHERE CodigoProgramacaoDiaria = :pCodigoProgramacaoDiaria';

    Q.Parameters.ParamByName('pCodigoProgramacaoDiaria').DataType := ftInteger;
    Q.Parameters.ParamByName('pCodigoProgramacaoDiaria').Value := ACodigoProgramacaoDiaria;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function TFrmPrincipal.corrigirData(strData: TDateTime): String;
  var
    dataCorrigida: String;
begin
  dataCorrigida:= (FormatDateTime('dd/mm/yyyy',strData));
  MaskEditData.Text:= dataCorrigida;
  Result:= MaskEditData.Text;
end;

procedure TFrmPrincipal.CotaReta(RChartMapa: TRChart; no1,no2: TPointFloat;StatusBarMapa: TStatusBar);
  var
    Dimensao: String;
    ptTexto: TPointFloat;
begin
  with RChartMapa do
  begin
    //Confiuração das linhas
    TransparentItems:= false;
    FillColor:= clNavy;
    PenStyle:= psSolid;
    DataColor:= clNavy;
    //TEXTO
    Dimensao:= FormatFloat('0.00 km',
    FrmPrincipal.DistanciaPontos(No1.X,No1.Y,No2.X,No2.Y));
    ptTexto.X := (No1.X + No2.X)/2;//Coordenada X
    ptTexto.Y := (No1.Y + No2.Y)/2;//Coordenada Y
    Arrow(No1.X,No1.Y,No2.X,No2.Y,7);
    Arrow(No2.X,No2.Y,No1.X,No1.Y,7);
    TextBkStyle:= tbClear;
    TextAlignment:= taLeftJustify;
    TextBkColor:= ChartColor;
    DataColor:= clNavy;
    Text(ptTexto.X,ptTexto.Y,10,Dimensao);
    Repaint;
    StatusBarMapa.Panels[3].Text:= Dimensao;
    AutoFitStatusBar(StatusBarMapa);
  end;
end;

function TFrmPrincipal.CountChecked(Grid: TStringGrid): Integer;
  var
    i,contar: Integer;
begin
  contar:= 0;
  for i := 1 to Grid.RowCount-1 do
  begin
    if Grid.Cells[0,i] = ' ' then
      contar:= contar+1;
  end;
  Result:= contar;
end;

procedure TFrmPrincipal.criarVariavelTempoExecucao(Field, Tabela,
  TipoField: String; sourceQuery: TADOQuery);
  var
    Variavel: String;
begin
  try
    sourceQuery.Active:= true;
    Variavel:= sourceQuery.FieldByName(Field).AsString;
  except
    CriarFieldDB(Field,Tabela,TipoField,sourceQuery);
  end;
end;

function TFrmPrincipal.dadosColuna(txtCaption: String; ColunasLayout: TStringGrid): TStringList;
  var
    i: Integer;
begin
  Result:=TStringList.Create;
  for i := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[1,i] = txtCaption then
    begin
      Result.Add(ColunasLayout.Cells[0,i]);//FieldName
      Result.Add(ColunasLayout.Cells[1,i]);//Titulo da coluna
      Result.Add(ColunasLayout.Cells[2,i]);//Alinhamento
      Result.Add(ColunasLayout.Cells[3,i]);//ReadOnly
      Result.Add(ColunasLayout.Cells[6,i]);//Largura coluna
    end;
  end;
end;

function TFrmPrincipal.dataFormatar(SourceData: String): TDateTime;
  var
    charArray: array [0..10] of char;
    i: Integer;
    ano,mes,dia: String;
begin
  try
    Result:= StrToDateTime(SourceData);
  except
    if SourceData <> '' then
    begin
      for i := 0 to Length(SourceData) do
        charArray[i]:= SourceData[i];

      ano:= charArray[1]+charArray[2]+charArray[3]+charArray[4];
      mes:= charArray[5]+charArray[6];
      dia:= charArray[7]+charArray[8];
      Result:= EncodeDate(StrToInt(ano),StrToInt(mes),StrToInt(dia));
    end
    else
    begin
      Result:= EncodeDate(1900,1,1);
    end;
  end;
end;

function TFrmPrincipal.dataLimpa(SourceData: String): String;
  var
    charArray: array [0..6] of char;
    i: Integer;
    ano,mes,dia,strData: String;
begin
  Result:= '';
  strData:= FrmPrincipal.SomenteNumero(SourceData);
  if strData <> '' then
  begin
    for i := 0 to Length(strData) do
      charArray[i]:= strData[i];

    dia:= charArray[1]+charArray[2];
    mes:= charArray[3]+charArray[4];
    ano:= charArray[5]+charArray[6];

    Result:= dia+'/'+mes+'/'+ano;
  end;
end;

function TFrmPrincipal.dataSAP(SourceData: String): String;
begin
  Result := '';

  if Trim(SourceData) = '' then
    Exit;

  // DD/MM/YYYY → DD.MM.YYYY
  Result :=
    Copy(SourceData, 1, 2) + '.' +
    Copy(SourceData, 4, 2) + '.' +
    Copy(SourceData, 7, 4);
end;

function TFrmPrincipal.DataSAP_ToDate(const SourceData: string): TDateTime;
var
  Dia, Mes, Ano: Word;
begin
  if Trim(SourceData) = '' then
    raise Exception.Create('Data SAP vazia');

  // Ex: 15/02/2026
  Dia := StrToInt(Copy(SourceData, 1, 2));
  Mes := StrToInt(Copy(SourceData, 4, 2));
  Ano := StrToInt(Copy(SourceData, 7, 4));

  Result := EncodeDate(Ano, Mes, Dia);
end;

procedure TFrmPrincipal.DBComboBoxPreventivaKeyPress(Sender: TObject; var Key: Char);
begin
key:= #0;
end;

function TFrmPrincipal.DecimalHoras(cValor: Double): String;
  var
    MMinteiro,HHinteiro: Integer;
    HHfracional,mm: Double;
begin
  //Converter a Hora Decimal em HH:MM:SS
  //- Multiplicar a parte fraccional por 60 -> a parte inteira são os minutos
  //- Multiplicar a parte fraccional do resultado anterior por 60 -> a parte inteira são os segundos
  //Ex: 13.595
  //- 0.595*60 = 35.7 -> 35 minutos
  //- 0.7* 60 = 42 -> 42 segundos
  //resultado = 13:35:42
  HHinteiro:= trunc(cValor);
  HHfracional:= cValor-HHinteiro;
  mm:= HHfracional*60;
  MMinteiro:= trunc(mm);
  //MMfracional:= mm-HHinteiro;
  //ss:= MMfracional*60;
  Result:= FormatFloat('00',HHinteiro)+
  ':'+FormatFloat('00',MMinteiro);
end;

function TFrmPrincipal.DeleteChar(const Ch: Char; const S: string): string;
var
  Posicao: integer;
begin
  Result := S;
  Posicao := Pos(Ch, Result);
  while Posicao > 0 do
  begin
    Delete(Result, Posicao, 1);
    Posicao := Pos(Ch, Result);
  end;
end;

procedure TFrmPrincipal.DeleteCol(Grid: TStringGrid; ACol: Integer);
  var
    i: Integer;
begin
  for i := ACol to Grid.ColCount-1 do
    Grid.Cols[i].Assign(Grid.Cols[i+1]);

  Grid.ColCount:= Grid.ColCount-1;
end;

procedure TFrmPrincipal.deleteQuery(ADOQuery: TADOQuery; nomeTabela: String);
begin
  FrmPrincipal.ProgressBarIncializa(ADOQuery.RecordCount,
  'Excluindo registros: '+nomeTabela+'...');
  ADOQuery.Active := true;
  ADOQuery.First;
  while not(ADOQuery.Eof) do
  begin
    try
      ADOQuery.Delete;
    except
      ADOQuery.Next;
    end;
    //Incremento ProgressBar
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  ADOQuery.Active := false;
  ADOQuery.Active := true;
  //Atualizar ProgressBar
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmPrincipal.deleteQueryRapido(ADOQuery: TADOQuery; nomeTabela: String);
  var
    SQLBase: String;
begin
  try
    SQLBase:= ADOQuery.SQL.Text;
    ADOQuery.SQL.Clear;
    ADOQuery.SQL.Add('DELETE * FROM '+nomeTabela);
    ADOQuery.ExecSQL;
    ADOQuery.Close;
    ADOQuery.SQL.Clear;
    ADOQuery.SQL.Add(SQLBase);
    ADOQuery.Open;
  except
  end;
end;

procedure TFrmPrincipal.deleteRepetidosCheckListBox(CheckListBox: TCheckListBox);
  var
    i: Integer;
    ListaDelete: TStringList;
begin
  ListaDelete:=  TStringList.Create;
  for I := CheckListBox.Items.count - 1 downto 1 do
    ListaDelete.Add(CheckListBox.Items[i]);

  deleteRepetidosList(ListaDelete);
  CheckListBox.Items.Clear;
  for I := 0 to ListaDelete.Count-1 do
    CheckListBox.Items.Add(ListaDelete[i]);
end;

procedure TFrmPrincipal.deleteRepetidosCombo(ComboBox: TComboBox);
  var
    i: Integer;
    ListaDelete: TStringList;
begin
  ListaDelete:=  TStringList.Create;
  for I := ComboBox.Items.count - 1 downto 1 do
    ListaDelete.Add(ComboBox.Items[i]);

  deleteRepetidosList(ListaDelete);
  ComboBox.Items.Clear;
  for I := 0 to ListaDelete.Count-1 do
    ComboBox.Items.Add(ListaDelete[i]);
end;

procedure TFrmPrincipal.deleteRepetidosList(Lista: TStringList);
  var
    i: Integer;
begin
  Lista.Sort;
  for I := Lista.count - 1 downto 1 do
  begin
    if UPPERCASE(Lista[I-1]) = UPPERCASE(Lista[I]) then
       Lista.Delete(I);
  end;
end;

procedure TFrmPrincipal.DeleteRow(Grid: TStringGrid; ARow: Integer);
  var
    i: Integer;
begin
  if Grid.RowCount > 1 then
  begin
    for i := ARow to Grid.RowCount-1 do
      Grid.Rows[i].Assign(Grid.Rows[i+1]);

    Grid.RowCount:= Grid.RowCount-1;
  end;
end;
{
procedure TFrmPrincipal.Descomprime(ArquivoZip: TFileName;
  DiretorioDestino: string);
  var
    archiver : TZipForge;
begin
  // Create an instance of the TZipForge class
  archiver := TZipForge.Create(nil);
  try
  with archiver do
  begin
    // The name of the ZIP file to unzip
    FileName := ArquivoZip;
    // Open an existing archive
    OpenArchive(fmOpenRead);
    // Set base (default) directory for all archive operations
    BaseDir := DiretorioDestino;
    // Extract all files from the archive to C:\Temp folder
    ExtractFiles('*.*');
    CloseArchive();
  end;
  except
  on E: Exception do
    begin
      Writeln('Exception: ', E.Message);
      // Wait for the key to be pressed
      Readln;
    end;
  end;
end;  }

procedure TFrmPrincipal.DesenharCalendario(Calendario: TStringGrid;
  PanelNome: TPanel;PrimeiroDiaMes: TDateTime);
var
  Linha, Coluna,DiaDaSemana: Integer;
  DiaPrimeiroMes: TDateTime;
  i : Integer;
  RecCalendar: TGridRect;
begin
  PanelNome.Caption := PrimeiraLetraMaiscula(FormatDateTime('mmmm yyyy', PrimeiroDiaMes));
  DiaPrimeiroMes := PrimeiroDiaMes - 1;
  DiaDaSemana:= DayOfWeek(DiaPrimeiroMes);
  if DiaDaSemana = 1 then
  begin
    i := 0;
    with Calendario do
    begin
      for Linha := 1 to Pred(RowCount) do
      begin
        for Coluna := 0 to Pred(ColCount) do
        begin
          Cells[Coluna, Linha]  := DateToStr(DiaPrimeiroMes + i);
          if DiaPrimeiroMes + i = Date then
          begin
            RecCalendar.Left    := Coluna;
            RecCalendar.Right   := Coluna;
            RecCalendar.Top     := Linha;
            RecCalendar.Bottom  := Linha;
            // Seleciona o dia corrente
            Selection := RecCalendar;
          end;
          Inc(i);
        end;
      end;
    end;
  end
  else
  begin
    while DiaDaSemana <> 1 do
    begin
      DiaPrimeiroMes := DiaPrimeiroMes - 1;
      i := 0;
      with Calendario do
      begin
        for Linha := 1 to Pred(RowCount) do
        begin
          for Coluna := 0 to Pred(ColCount) do
          begin
            Cells[Coluna, Linha]  := DateToStr(DiaPrimeiroMes + i);
            if DiaPrimeiroMes + i = Date then
            begin
              RecCalendar.Left    := Coluna;
              RecCalendar.Right   := Coluna;
              RecCalendar.Top     := Linha;
              RecCalendar.Bottom  := Linha;
              // Seleciona o dia corrente
              Selection := RecCalendar;
            end;
            Inc(i);
          end;
        end;
      end;
      DiaDaSemana:= DayOfWeek(DiaPrimeiroMes);
    end;
  end;
end;

function TFrmPrincipal.DiasFinalCorridos(DataInicio: TDateTime;
  DiasCorridos: Integer): TDateTime;
begin
  Result:= DataInicio+DiasCorridos;
end;

function TFrmPrincipal.DiasFinalUtil(DataInicio: TDateTime;
  DiasUteis: Integer): TDateTime;
  var
    DaysWeek: Integer;
    data: TDateTime;
begin
  DaysWeek:= DayOfWeek(DataInicio)-1;
  data:=  DataInicio+DiasUteis+((DiasUteis-1+DaysWeek)div 5)*2;
  Result:= data;
end;

function TFrmPrincipal.DistanciaPontos(X1, Y1, X2, Y2: Double): Double;
begin
  Result:= SQRT(POWER(X2-X1,2)+POWER(Y2-Y1,2));
end;

procedure TFrmPrincipal.DownLoadQuery(GravarQuery: TADOQuery;
  GravarDataSource: TDataSource; parametroREF, parametroGravar, campo1: Integer;
  parametroString, nomeTabela,SQLConsulta: String; booleanInsert,booleanExecutante: Boolean);
var
  ConsultaQuery: TADOQuery;
  ConsultaDataSource: TDataSource;
  i: Integer;
begin
  try
    begin
      GravarQuery.Active := false;
      GravarQuery.Active := true;
      // ==========================================================
      ConsultaQuery := TADOQuery.Create(nil);
      ConsultaDataSource := TDataSource.Create(nil);
      ConsultaQuery.Connection := FrmDataModule.ADOConnectionColibri;
      ConsultaDataSource.DataSet := ConsultaQuery;
      ConsultaQuery.sql.Clear;
      ConsultaQuery.sql.Text := SQLConsulta;
      // Tabela de Consulta
      ConsultaQuery.Active := false;
      if (parametroREF <> 0) then
        ConsultaQuery.Parameters.Items[0].Value := parametroREF;
      ConsultaQuery.Active := true;
      // ===========================================================
      // Copiar tabela de carga de referência
      ConsultaQuery.First;
      while not(ConsultaQuery.Eof) do
      begin
        if booleanInsert then
          GravarQuery.Insert
        else
          GravarQuery.Edit;
        for i := campo1 to ConsultaQuery.FieldCount - 1 do
        begin
          GravarDataSource.DataSet.Fields[i].Value :=
            ConsultaDataSource.DataSet.Fields[i].Value;
        end;
        if (parametroGravar <> 0) then
          GravarDataSource.DataSet.FieldByName(parametroString).AsInteger :=
            parametroGravar;

        if booleanExecutante then
        begin
          GravarDataSource.DataSet.FieldByName('StatusProgramacao').AsString:= 'Aprovado';
          GravarDataSource.DataSet.FieldByName('MotivoProgramacao').AsString:= '';
          GravarDataSource.DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean:= false;
        end;

        GravarQuery.Post;
        ConsultaQuery.Next;
        if booleanInsert = false then
          GravarQuery.Next;
      end;
      // GravarDataSource.Enabled:= true;
      GravarQuery.Active := false;
      GravarQuery.Active := true;
      ConsultaQuery.Active := false;
      ConsultaQuery.Active := true;
    end;
  except
    ShowMessage('Ocorreu um erro na cópia da tabela "'+ nomeTabela +
    '" e a operação foi cancelada!');
    GravarQuery.Active := false;
    GravarQuery.Active := true;
    //Atualizar ProgressBar
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

procedure TFrmPrincipal.Embarcacao1Click(Sender: TObject);
begin
  if not Assigned(FrmEmbarcacao) then
    FrmEmbarcacao:= TFrmEmbarcacao.Create(Application)
  else
    FrmEmbarcacao.Show;
end;

procedure TExcluirTableDB(NomeTabela: String;
  sourceQuery: TADOQuery);
  var
    SQLBase: String;
begin
  try
    SQLBase:= 'DROP TABLE '+NomeTabela+';';
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(SQLBase);
    sourceQuery.ExecSQL;
  except
  end;
end;

procedure TFrmPrincipal.Executantes1Click(Sender: TObject);
begin
  if not Assigned(FrmExecutante) then
    FrmExecutante:= TFrmExecutante.Create(Application)
  else
    FrmExecutante.Show;
end;

procedure TFrmPrincipal.ExecutantesProgramadosporPerodo1Click(Sender: TObject);
begin
  if not Assigned(FrmConsultaExecutantesProgramados) then
    FrmConsultaExecutantesProgramados:= TFrmConsultaExecutantesProgramados.Create(Application)
  else
    FrmConsultaExecutantesProgramados.Show;
end;

procedure TFrmPrincipal.Fechar1Click(Sender: TObject);
begin
  Application.Terminate;
end;

function TFrmPrincipal.ForaOperacao(Plataforma: String; IndexSistema: Integer): String;
  var
    i: Integer;
begin
  Result:= '';
  for i := 0 to High(MatrizForaOperacao[0]) do
  begin
    if Plataforma = MatrizForaOperacao[0,i] then
    begin
      Result:= MatrizForaOperacao[IndexSistema,i];
      break;
    end;
  end;
  {
  MatrizForaOperacao[0,i] Plataforma;
  MatrizForaOperacao[1,i] 'SituacaoGD;
  MatrizForaOperacao[2,i] CapPrincipal;
  MatrizForaOperacao[3,i] CapAuxiliar;
  MatrizForaOperacao[4,i] SituacaoBCI;
  MatrizForaOperacao[5,i] SituacaoLinhaBCI;
  MatrizForaOperacao[6,i] SituacaoAgua;
  MatrizForaOperacao[7,i] SituacaoBalsa;

  MatrizForaOperacao[9,i] SituacaoDegraus;
  MatrizForaOperacao[10,i] DataRealizacaoDegraus;
  MatrizForaOperacao[11,i] SituacaoSurfer;
  MatrizForaOperacao[12,i] SituacaoSOV;
  MatrizForaOperacao[13,i] SituacaoAqua;}
end;

function TFrmPrincipal.FormataCPF(Texto: String): String;
  Var vTam, xx : Integer;
      vDoc : String;
begin
  vTam := Length(Texto);
  vDoc := '';
  For xx := 1 To vTam Do
  begin
  vDoc := vDoc + Copy(Texto,xx,1);
    If vTam = 11 Then
    begin
      If (xx in [3,6]) Then vDoc := vDoc + '.';
      If xx = 9 Then vDoc := vDoc + '-';
    end;
  end;
  Result := vDoc;
end;

procedure TFrmPrincipal.FormCreate(Sender: TObject);
begin
  FrmTelaAbertura.ProgressBarTelaAbertura.Value:= 0;
  booLOCAL:= false;
  filtroFO:= 'WHERE (((TextoBreveNOTA LIKE "FO:%"))AND((FimAvaria IS NULL)))';
  if ParamCount > 0 then
    enderecoColibriRegistro:= ParamStr(1)
  else
    enderecoColibriRegistro:= registroEndereco('Banco de dados');
  //=====================================================

  AbrirBancoDados(enderecoColibriRegistro,UpperCase(usuarioWindows),true);
  HintPadrao:= 'Entre com o texto de procura e clique em Filtrar'+#13+
  '* Utilize ";" para separar as palavras em condição "OU" e "&" para condição "E";';
  ProgressBarPrincipal.Value:= 0;
  StatusBarPrincipal.Panels[6].Style := psOwnerDraw;
  ProgressBarPrincipal.Parent := StatusBarPrincipal;
end;

procedure TFrmPrincipal.FormDestroy(Sender: TObject);
begin
  FrmPrincipal:=nil;
end;

procedure TFrmPrincipal.GeraTexto(DataSource: TDataSource; Separador: Char; nomeTabela: String);
  var intC1, intC2: Integer;
      strLinha: string;
begin
  SaveDialog1.Filter := 'Texto|*.txt';
  SaveDialog1.DefaultExt := '*.txt';
  SaveDialog1.FileName:= nomeTabela;
  if SaveDialog1.Execute then
  begin
    MemoPrincipal.Lines.Clear;
    strLinha := '';
    DataSource.Enabled:= false;
    //Cria a primeira linha com os nomes das colunas
    for intC1 := 0 to DataSource.DataSet.Fields.Count - 1 do
      if intC1 < DataSource.DataSet.Fields.Count - 1 then
        strLinha := strLinha + DataSource.DataSet.Fields[intC1].DisplayLabel + Separador
      else
        strLinha := strLinha + DataSource.DataSet.Fields[intC1].DisplayLabel;

    strLinha:= RetiraEnter(strLinha);
    MemoPrincipal.Lines.Add(strLinha);
    strLinha := '';

    FrmPrincipal.ProgressBarIncializa(DataSource.DataSet.RecordCount,'Gerando arquivo *txt...');
    for intC1 := 1 to DataSource.DataSet.RecordCount do
      begin
        for intC2 := 0 to DataSource.DataSet.Fields.Count - 1 do
          if intC2 < DataSource.DataSet.Fields.Count - 1 then
            strLinha := strLinha + DataSource.DataSet.Fields[intC2].AsString + Separador
          else
            strLinha := strLinha + DataSource.DataSet.Fields[intC2].AsString;

        strLinha:= RetiraEnter(strLinha);
        MemoPrincipal.Lines.Add(strLinha);
        strLinha := '';
        DataSource.DataSet.Next;
        FrmPrincipal.ProgressBarIncremento(1);
      end;

    MemoPrincipal.Lines.SaveToFile(SaveDialog1.FileName);
    DataSource.DataSet.First;
    DataSource.Enabled:= true;
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

procedure TFrmPrincipal.GerenciarSolicitaes1Click(Sender: TObject);
begin
  if not Assigned(FrmGerenciarSolicitacoes) then
    FrmGerenciarSolicitacoes:= TFrmGerenciarSolicitacoes.Create(Application)
  else
    FrmGerenciarSolicitacoes.Show;
end;

procedure TFrmPrincipal.GerenciarTransportes1Click(Sender: TObject);
begin
  if not Assigned(FrmGerenciarEmbarcacoes) then
    FrmGerenciarEmbarcacoes:= TFrmGerenciarEmbarcacoes.Create(Application)
  else
    FrmGerenciarEmbarcacoes.Show;
end;

procedure TFrmPrincipal.GravarCanceladoAprovado(idProgramacaoDiaria: Integer);
  var
    NumCancelados,NumAprovados,NumExecutantes: Integer;
begin
  FrmDataModule.ADOQueryNumAprovado.Active:= false;
  FrmDataModule.ADOQueryNumAprovado.Parameters.Items[0].Value:=idProgramacaoDiaria;
  FrmDataModule.ADOQueryNumAprovado.Active:= true;
  NumAprovados:= FrmDataModule.ADOQueryNumAprovado.RecordCount;
  //================================================
  FrmDataModule.ADOQueryNumCancelado.Active:= false;
  FrmDataModule.ADOQueryNumCancelado.Parameters.Items[0].Value:=idProgramacaoDiaria;
  FrmDataModule.ADOQueryNumCancelado.Active:= true;
  NumCancelados:= FrmDataModule.ADOQueryNumCancelado.RecordCount;
  //================================================
  FrmDataModule.ADOQueryNumExecutantes.Active:= false;
  FrmDataModule.ADOQueryNumExecutantes.Parameters.Items[0].Value:= idProgramacaoDiaria;
  FrmDataModule.ADOQueryNumExecutantes.Active:= true;
  NumExecutantes:= FrmDataModule.ADOQueryNumExecutantes.RecordCount;
  //Gravar N° de Cancelados e Aprovados
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= false;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Parameters.Items[0].Value:= idProgramacaoDiaria;
  FrmDataModule.ADOQueryConsultaProgramacao_ID.Active:= true;
  //================================================
  try
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumExecutantes').AsInteger:= NumExecutantes;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumAprovados').AsInteger:= NumAprovados;
    FrmDataModule.DataSourceConsultaProgramacao_ID.DataSet.
    FieldByName('NumCancelados').AsInteger:= NumCancelados;
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Post;
  except
    FrmDataModule.ADOQueryConsultaProgramacao_ID.Cancel;
  end;
end;

function TFrmPrincipal.InsertRow1(StrGrid: TStringGrid;Linha: integer): integer;
  Var
    Rown,Coln: Integer;
begin
  strgrid.RowCount:=strgrid.RowCount+1;
  if Linha<1 then
    Linha:=1
  else if Linha>strgrid.RowCount-1 then
    Linha:=strgrid.RowCount-1;
  if strgrid.RowCount-1 <> Linha then
  begin
    for Rown:=strgrid.RowCount-1 downto Linha do
      for Coln:=0 to strgrid.ColCount-1 do
         strgrid.Cells[Coln,Rown]:=strgrid.Cells[Coln,Rown-1];
  end;
  result:=Linha;
end;

function TFrmPrincipal.HoraParaMin(Hora: String): Integer;
begin
  try
    Result:= StrToInt(Copy(Hora,1,2))*60+StrToInt(Copy(Hora,4,2));
  except
    Result:= 0;
  end;
end;

procedure TFrmPrincipal.ImageExcluir(NomeArquivo: String);
var
  CaminhoArquivo: String;
begin
  CaminhoArquivo := FrmDataModule.DataSourceColibri.DataSet.FieldByName
  ('enderecoMarinha').AsString+'\Fotos\'+NomeArquivo;
  if FileExists(CaminhoArquivo) then
  begin
    //Excluir o arquivo
    DeleteFile(CaminhoArquivo);
    //Mensagem de exclusão
    MessageBox(0,'Arquivo excluído com sucesso!','Excluir arquivo',
    MB_ICONINFORMATION);
  end
  else
  begin
    MessageBox(0,'O arquivo não existe ou já foi eliminado.',
    'Excluir arquivo',MB_ICONERROR)
  end;
end;

procedure TFrmPrincipal.importarExcel(nomeTabela: String;
    ACol,ARow,TabSheet: Integer; Tabela: TDBGrid; ADOQuery: TADOQuery);
var
  Excel, Sheet: Variant;
  i, j, matrizLinha, matrizColuna,DataCol: Integer;
  FieldColuna: String;
begin
  OpenDialog1.Filter := 'Microsoft Excel |*.xlsx|Microsoft Excel 97-2003|*.xls';
  OpenDialog1.DefaultExt := '*.xlsx';

  if OpenDialog1.Execute then
  begin
    try
      Excel := CreateOleObject('Excel.Application');
      // Abrir o arquivo
      Excel.Workbooks.Open(OpenDialog1.FileName);
      // Abrir a primeira planilha do arquivo
      Sheet := Excel.WorkSheets[1];
      // ============================================
      //Excluir todos os registros atuais
      if Application.MessageBox(PChar(
      'Deseja excluir os registros "FILTRADOS" antes da importação?'),'.::ATENÇÃO::.',36) = 6 then
        deleteQuery(ADOQuery,nomeTabela);
      //Verificar a ultima Linha
      matrizLinha:= (Sheet.Cells.SpecialCells(11).Row);
      //Verificar a ultima Coluna
      matrizColuna:= Tabela.Columns.Count;
      Tabela.DataSource.Enabled := false;
      // Inicializar ProgressBar
      FrmPrincipal.ProgressBarIncializa(matrizLinha,
      'Importar Excel: '+nomeTabela+'...');
      // ============================================
      for i := ARow to matrizLinha do
      begin
        DataCol:= 0;
        // Inserir novo registro na tabela de dados
        Tabela.DataSource.DataSet.Insert;
        // Preencher as colunas do registro na tabela de dados
        for j := ACol to matrizColuna do
        begin
          FieldColuna := Tabela.Columns.Items[DataCol].FieldName;
          try
            Tabela.DataSource.DataSet.FieldByName(FieldColuna).AsString :=
              Excel.Workbooks[1].Sheets[TabSheet].Cells[i,j];
          except
            Tabela.DataSource.DataSet.FieldByName(FieldColuna).AsString := '';
          end;
          DataCol:= DataCol + 1;
        end;
        try
          Tabela.DataSource.DataSet.Post;
        except
          Tabela.DataSource.DataSet.Cancel;
        end;
        // Incremento ProgressBar
        FrmPrincipal.ProgressBarIncremento(1);
      end;
      Tabela.DataSource.DataSet.First;
      Tabela.DataSource.Enabled := true;
      // Atualizar ProgressBar
      FrmPrincipal.ProgressBarAtualizar;
      //Fechar Excel
      Excel.Application.DisplayAlerts := False; // Alle Rückfragen ausstellen
      Excel.Quit;
      Excel := Unassigned;
      Excel := 0;
    except
      MessageBox(0, 'Erro de conexão com o Excel' + #13 + #10 +
      'Verifique se o Microsoft Office Excel esta instalado na sua máquina.',
      'Excel', MB_ICONERROR);
      // ==========================================================================
      Excel.Workbooks.Close;
      Tabela.DataSource.DataSet.First;
      Tabela.DataSource.Enabled := true;
      // Atualizar ProgressBar
      FrmPrincipal.ProgressBarAtualizar;
      //Fechar Excel
      Excel.Application.DisplayAlerts := False; // Alle Rückfragen ausstellen
      Excel.Quit;
      Excel := Unassigned;
      Excel := 0;
    end;
  end;
end;

procedure TFrmPrincipal.ImportDataAccess(const AccessDb, TableName,
  ExcelFileName: String);
  Const
    acQuitSaveAll             = $00000001;
    acImport                  = $00000000;
    acSpreadsheetTypeExcel9   = $00000008;
    acSpreadsheetTypeExcel12  = $00000009;
  var
    LAccess : OleVariant;
begin
   //create the COM Object
   LAccess := CreateOleObject('Access.Application');
   //open the access database
   try
      LAccess.OpenCurrentDatabase(AccessDb,true);
   except
      //if the access database doesn't exist use the NewCurrentDatabase method instead.
      LAccess.NewCurrentDatabase(AccessDb,true);
   end;
   //import the data
   LAccess.DoCmd.TransferSpreadsheet(acImport, acSpreadsheetTypeExcel9,
   TableName, ExcelFileName, true);
   LAccess.CloseCurrentDatabase;
   LAccess.Quit(1);
end;

procedure TFrmPrincipal.ImportExcelAcess(enderecoExcel,enderecoDB,Tabela: String;
ListaColunas: TStringList;ADOQuery: TADOQuery);
  var
    Excel: Variant;
    i: Integer;
    primeiraLinha: String;
begin
  try
    FrmPrincipal.ProgressBarIncializa(30,'Importando planilha do Excel...');
    FrmPrincipal.deleteQueryRapido(ADOQuery,Tabela);
    FrmPrincipal.ProgressBarIncremento(2);
    //Conexão com a planilha do PI
    Excel := CreateOleObject('Excel.Application');
    FrmPrincipal.ProgressBarIncremento(4);
    // Abrir o arquivo
    Excel.Workbooks.Open(enderecoExcel);
    FrmPrincipal.ProgressBarIncremento(2);
    //Excluir primeiras linhas em branco
    primeiraLinha:= Excel.Workbooks[1].Sheets[1].Cells[1,2];
    while primeiraLinha = '' do
    begin
      Excel.WorkBooks[1].Sheets[1].Rows.Rows[1].delete;
      primeiraLinha:= Excel.Workbooks[1].Sheets[1].Cells[1,1];
    end;
    //Alterar titulo das colunas com nome das variaveis da tabela do banco de dados
    for I := 0 to ListaColunas.Count-1 do
      Excel.Workbooks[1].Sheets[1].Cells[1,i+1]:= ListaColunas[i];
    //Salvar planilha Excel
    Excel.ActiveWorkbook.Save;
    FrmPrincipal.ProgressBarIncremento(2);
    try
      Excel.Application.DisplayAlerts := False;
      Excel.Quit;
      Excel := Unassigned;
      Excel := 0;
    except
    end;
    try
      CoInitialize(nil);
      try
        FrmPrincipal.ImportDataAccess(enderecoDB,Tabela,enderecoExcel);
        Writeln('Done');
      finally
        CoUninitialize;
      end;
    except
      on E:EOleException do
        Writeln(Format('EOleException %s %x', [E.Message,E.ErrorCode]));
      on E:Exception do
        Writeln(E.Classname, ':', E.Message);
    end;
    Writeln('Pressione enter para sair!');
    Readln;
  except
  end;
  FrmPrincipal.ProgressBarAtualizar;
end;

function TFrmPrincipal.incialListaBusca(CampoString: String): TStringList;
begin
  Result:= TStringList.Create;
  //Definicoes gerais
  Result.delimiter:= ';';
  Result.StrictDelimiter:= true;
  Result.DelimitedText:= CampoString;
end;

function TFrmPrincipal.incluirProgramacao(DataProgramacao, Destino,
  txtTipoEtapaServico, CriadoPor, Computador: String;
  DataCriacao: TDate): Integer;
  var
    LogAcao: String;
begin
  try
    FrmDataModule.ADOQueryInserirProgramacao.Active:= true;
    FrmDataModule.ADOQueryInserirProgramacao.Insert;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('DataProgramacao').AsString:= DataProgramacao;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('txtDestino').AsString:= Destino;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('txtTipoEtapaServico').AsString:= txtTipoEtapaServico;
    //=========================================
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('NumExecutantes').AsInteger:= 1;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('NumCancelados').AsInteger:= 0;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('NumAprovados').AsInteger:= 1;
    //=========================================
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('CriadoPor').AsString:= CriadoPor;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('DataCriacao').AsDateTime:= DataCriacao;
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('ComputadorCriacao').AsString:= Computador;
    //=========================================
    MemoPrincipal.Text:= FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('LogAcao').AsString;
    LogAcao:= 'Programação: '+DataProgramacao+#9+Destino+#9+txtTipoEtapaServico+#9+
    '['+CriadoPor+'; '+DateToStr(DataCriacao)+'; '+Computador+'; Criação]';
    MemoPrincipal.Lines.Add(LogAcao);
    FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('LogAcao').AsString:= MemoPrincipal.Text;
    //==========================================
    FrmDataModule.ADOQueryInserirProgramacao.Post;
    Result:= FrmDataModule.DataSourceInserirProgramacao.DataSet.
    FieldByName('idProgramacaoDiaria').AsInteger;
  except
    Result:= 0;
  end;
end;

function TFrmPrincipal.incluirServicoPadrao(TipoEtapaServico: String): String;
  var
    SQLBase: String;
begin
  try
    SQLBase:=
    'SELECT tblTipoEtapaServico.* FROM tblTipoEtapaServico '+
    'WHERE ((TipoEtapaServico like '+QuotedStr('%'+TipoEtapaServico+'%')+'))';
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
    Result:= FrmDataModule.DataSourceTemporarioDBConsulta1.DataSet.FieldByName('ServicoPadrao').AsString;
  except
  end;
end;

function TFrmPrincipal.indexLayoutCaption(Caption: String;
  ColunasLayout: TStringGrid): Integer;
var
  I: Integer;
begin
  Result:= 0;
  for I := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[1,i] = Caption then
    begin
      Result:= i;
      break;
    end;
  end;
end;

function TFrmPrincipal.indexLayoutFieldName(FieldName: String;
  ColunasLayout: TStringGrid): Integer;
var
  I: Integer;
begin
  Result:= 0;
  for I := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[0,i] = FieldName then
    begin
      Result:= i;
      break;
    end;
  end;
end;

function TFrmPrincipal.isData(txt: String): Boolean;
begin
  Result := true;
  Try
    StrToDate(txt);
  Except
    Result := false;
  End;
end;

function TFrmPrincipal.isNumeric(txt: String): Integer;
begin
  Try
    StrToFloat(txt);//Float
    Result := 2;
  Except
    try
      StrToDate(txt);//Data
      Result := 3;
    Except
      Result := 0;//Texto
    end;
  End;
end;

function TFrmPrincipal.LatidudeLongitude_Graus(cValor: Double): String;
  var
    MMinteiro,HHinteiro: Integer;
    MMfracional,HHfracional,mm,ss: Double;
begin
  //Converter a Hora Decimal em HH:MM:SS
  //- Multiplicar a parte fraccional por 60 -> a parte inteira são os minutos
  //- Multiplicar a parte fraccional do resultado anterior por 60 -> a parte inteira são os segundos
  //Ex: 13.595
  //- 0.595*60 = 35.7 -> 35 minutos
  //- 0.7* 60 = 42 -> 42 segundos
  //resultado = 13:35:42
  HHinteiro:= trunc(cValor);
  HHfracional:= cValor-HHinteiro;
  mm:= HHfracional*60;
  MMinteiro:= trunc(mm);
  MMfracional:= mm-HHinteiro;
  ss:= MMfracional*60;
  Result:= FormatFloat('00',HHinteiro)+
  ':'+FormatFloat('00',MMinteiro)+':'+FormatFloat('00.00',ss);
end;

function TFrmPrincipal.LatLong_XY(Latitude, Longitude: Double): TPointFloat;
begin
  {-5402.2744	  -45
  -3423.9126	  -30
  -1705.8616	  -15
  0,0000		    0
  1678.1782	    15
  3420.6079	    30
  5374.5909	    45
  7613.2634	    60
  10476.6817	  75}
  Result.X:= 111.12*Longitude;
  if ((Latitude < -45) OR (Latitude > 75)) then
    Result.Y:= 0
  else
  begin
    //60 - 75
    if ((Latitude >60)AND(Latitude <=75)) then
      Result.Y:= 7613.2634+((Latitude-60)*((10476.6817-7613.2634)/15))
    //45 - 60
    else if ((Latitude > 45)AND(Latitude <=60)) then
      Result.Y:= 5374.5909+((Latitude-45)*((7613.2634-5374.5909)/15))
    //30 - 45
    else if ((Latitude > 30)AND(Latitude <=45)) then
      Result.Y:= 3420.6079+((Latitude-30)*((5374.5909-3420.6079)/15))
    //15 - 30
    else if ((Latitude > 15)AND(Latitude <=30)) then
      Result.Y:= 1678.1782+((Latitude-15)*((3420.6079-1678.1782)/15))
    //0 - 15
    else if ((Latitude > 0)AND(Latitude <=15)) then
      Result.Y:= Latitude*((1678.1782)/15)
    //-15 - 0
    else if ((Latitude > -15)AND(Latitude <=0)) then
      Result.Y:= Latitude*(1705.8616/15)
    //-30 - 15
    else if ((Latitude > -30)AND(Latitude <=-15)) then
      Result.Y:= -1705.8616+((Latitude+15)*((3423.9126-1705.8616)/15))
    //-30 - 15
    else if ((Latitude >= -45)AND(Latitude <=-30)) then
      Result.Y:= -3423.9126+((Latitude+30)*((5402.2744-3423.9126)/15));
  end;
end;

procedure TFrmPrincipal.LayoutPadrao(NomeArquivoMemo:string; GridPadrao: TStringGrid; Tabela: String);
  var
    i: Integer;
    FText: TextFile;
    LocalArquivo: String;
begin
  MemoPrincipal.Lines.Clear;
  for i := 0 to GridPadrao.RowCount-1 do
    MemoPrincipal.Lines.Add(Tabela+GridPadrao.Cells[1,i]);

  LocalArquivo:= ExtractFilePath(Application.ExeName)+'Layout\'+NomeArquivoMemo;
  AssignFile(FText,LocalArquivo);
  if not FileExists(LocalArquivo) then
  begin
    Rewrite(FText);
    Append(FText);
    MemoPrincipal.Lines.SaveToFile(LocalArquivo);
  end
  else
    MemoPrincipal.Lines.SaveToFile(LocalArquivo);
end;

procedure TFrmPrincipal.limparStringGrid(Grid: TStringGrid);
  var
    i: Integer;
begin
  for I := 0 to Grid.RowCount-1 do
    Grid.Rows[i].Clear;
end;

procedure TFrmPrincipal.limparTitleGrid(Grid: TDBGrid);
  var
    i: Integer;
    txtString: String;
begin
  //Limpar simbolo de classificação ASC DESC
  for i:=0 to Grid.FieldCount-1 do
  begin
    Grid.Columns.Items[i].Title.Font.Color:=ClBlack;//Cor da Fonte
    Grid.Columns.Items[i].Title.Font.Style := [];//Cot da Fonte
    Grid.Columns.Items[i].Title.Color:=clBtnFace;//Cor do Fundo do Titulo Normal
    Grid.Columns.Items[i].Font.color:= ClBlack;
    txtString:= FrmPrincipal.DeleteChar(char(9650),Grid.Columns.Items[i].Title.Caption);
    txtString:= FrmPrincipal.DeleteChar(char(9660),txtString);
    txtString:= FrmPrincipal.DeleteChar(char(8722),txtString);
    Grid.Columns.Items[i].Title.Caption:= txtString;
  end;
end;

procedure TFrmPrincipal.ManifestodeEmbarque1Click(Sender: TObject);
begin
  if not Assigned(FrmManifestoEmbarque) then
    FrmManifestoEmbarque:= TFrmManifestoEmbarque.Create(Application)
  else
    FrmManifestoEmbarque.Show;
end;

procedure TFrmPrincipal.MDIChildCreated(const childHandle: THandle);
begin
  mdiChildrenTabs.Tabs.AddObject(TForm(FindControl(childHandle)).Caption, TObject(childHandle));
  mdiChildrenTabs.TabIndex := -1 + mdiChildrenTabs.Tabs.Count;
end;

procedure TFrmPrincipal.MDIChildDestroyed(const childHandle: THandle);
var
  idx: Integer;
begin
  idx := mdiChildrenTabs.Tabs.IndexOfObject(TObject(childHandle));
  mdiChildrenTabs.Tabs.Delete(idx);
end;

procedure TFrmPrincipal.mdiChildrenTabsChange(Sender: TObject; NewTab: Integer;
  var AllowChange: Boolean);
var
  k,cHandle,aba: Integer;
begin
  cHandle := Integer(mdiChildrenTabs.Tabs.Objects[NewTab]);
  if mdiChildrenTabs.Tag = -1 then
    Exit;
  //==================================
  for k := 0 to MDIChildCount - 1 do
  begin
    aba:= MDIChildren[k].Handle;
    if (aba = cHandle) then
    begin
      MDIChildren[k].Show;
      Break;
    end;
  end;
end;

procedure TFrmPrincipal.MenssagemStatus(Texto: String);
begin
  StatusBarPrincipal.Panels[4].Text:= Texto;
  StatusBarPrincipal.Panels[4].Width:= StatusBarPrincipal.Canvas.TextWidth
  (StatusBarPrincipal.Panels[4].Text)+20;
end;

procedure TFrmPrincipal.MovimentacaoCarga1Click(Sender: TObject);
begin
  if not Assigned(FrmMovimentacaoCarga) then
    FrmMovimentacaoCarga:= TFrmMovimentacaoCarga.Create(Application)
  else
    FrmMovimentacaoCarga.Show;
end;

procedure TFrmPrincipal.msgSplash(texto1,texto2: String);
begin
  //if not Assigned(FrmTelaAbertura) then
    FrmTelaAbertura.MemoMensagem.Lines.Add(texto1);
    FrmTelaAbertura.MemoMensagem.Lines.Add(texto2);
    FrmTelaAbertura.MemoMensagem.Lines.Add(
  '--------------------------------------------------------------------------------------------------------');
    FrmTelaAbertura.ProgressBarTelaAbertura.Value:= FrmTelaAbertura.ProgressBarTelaAbertura.Value+1;
    FrmTelaAbertura.MemoMensagem.Repaint;
    //Sleep(1000);
end;

function TFrmPrincipal.NomeMaquina: String;
var
  lpBuffer: PChar;
  nSize: DWord;
const
  Buff_Size = MAX_COMPUTERNAME_LENGTH + 1;
begin
  try
    nSize := Buff_Size;
    lpBuffer := StrAlloc(Buff_Size);
    GetComputerName(lpBuffer, nSize);
    Result := String(lpBuffer);
    StrDispose(lpBuffer);
  except
    Result := '';
  end;
end;

function TFrmPrincipal.OrigemPlataformas: String;
  var
    SQLBase,Plataformas,Origem: String;
begin
  SQLBase:=
  'SELECT tblPlataforma.* FROM tblPlataforma '+
  'WHERE(booleanOrigem = True);';
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  Origem:= '';
  Plataformas:= '';
  while not FrmDataModule.ADOQueryTemporarioDBConsulta1.Eof do
  begin
    Plataformas:=  FrmDataModule.ADOQueryTemporarioDBConsulta1.FieldByName('Plataforma').AsString;
    if Origem <> '' then
      Origem:= Origem+';'+Plataformas
    else
      Origem:= Plataformas;
    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
  Result:= Origem;
end;

function TFrmPrincipal.CalcNumCanceladosAprovado(idProgramacaoDiaria,Tipo: Integer): Integer;
begin
  case Tipo of
    0:
    begin
      FrmDataModule.ADOQueryNumCancelado.Active:= false;
      FrmDataModule.ADOQueryNumCancelado.Parameters.Items[0].Value:=
      idProgramacaoDiaria;
      FrmDataModule.ADOQueryNumCancelado.Active:= true;
      Result:= FrmDataModule.ADOQueryNumCancelado.RecordCount;
    end;
    1:
    begin
      FrmDataModule.ADOQueryNumAprovado.Active:= false;
      FrmDataModule.ADOQueryNumAprovado.Parameters.Items[0].Value:=
      idProgramacaoDiaria;
      FrmDataModule.ADOQueryNumAprovado.Active:= true;
      Result:= FrmDataModule.ADOQueryNumAprovado.RecordCount;
    end
    else
      Result:= 0;
  end;
end;

function TFrmPrincipal.CalcNumExecutantes(
  idProgramacaoDiaria: Integer): Integer;
begin
  FrmDataModule.ADOQueryNumExecutantes.Active:= false;
  FrmDataModule.ADOQueryNumExecutantes.Parameters.Items[0].Value:= idProgramacaoDiaria;
  FrmDataModule.ADOQueryNumExecutantes.Active:= true;
  Result:= FrmDataModule.ADOQueryNumExecutantes.RecordCount;
end;

function TFrmPrincipal.palavraBusca(SQLString,FieldBusca,Operador,StringBusca,CondicionalVariaveis: String): String;
  var
    ListaFiltroOU,ListaFiltroAND: TStringList;
    SQLStringAND,SQLStringFINAL,Sinal,tudo: String;
    i,j: Integer;
begin
  if StringBusca <> '' then
  begin
    if ((Operador) = 'Contem')or(Operador = '') then
    begin
      Sinal:= ' LIKE ';
      tudo:= '%';
    end
    else if (Operador) = 'Diferente' then
    begin
      Sinal:= ' NOT LIKE ';
      tudo:= '%';
    end
    else if (Operador) = 'Exato' then
    begin
      Sinal:= ' LIKE ';
      tudo:= '';
    end;
    //Gerar lista de condições OU
    ListaFiltroOU:= TStringList.Create;
    ListaFiltroOU.delimiter:= ';';
    ListaFiltroOU.StrictDelimiter:= true;
    ListaFiltroOU.DelimitedText:= StringBusca;
    SQLStringFINAL:='';
    //======================================================
    if ListaFiltroOU.Count > 0 then
    begin
      begin
        for i := 0 to ListaFiltroOU.Count-1 do
        begin
          if i = 0 then
          begin
            //Gerar lista de condições AND
            ListaFiltroAND:= TStringList.Create;
            ListaFiltroAND.delimiter:= '&';
            ListaFiltroAND.StrictDelimiter:= true;
            ListaFiltroAND.DelimitedText:= ListaFiltroOU[i];
            SQLStringAND:= '';
            //======================================================
            if ListaFiltroAND.Count > 1 then
            begin
              for j := 0 to ListaFiltroAND.Count-1 do
              begin
                if j = 0 then
                begin
                  if ((ListaFiltroAND[j] = 'Vazia')AND(Operador = 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' <> "")'
                  else if ((ListaFiltroAND[j] = 'Vazia')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' = "")'
                  else if ((ListaFiltroAND[j] = 'Nulo')AND(Operador = 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' IS NOT NULL)'
                  else if ((ListaFiltroAND[j] = 'Nulo')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' IS NULL)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'TRUE' then
                    SQLStringAND:= '(('+FieldBusca+ ' = TRUE)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'FALSE' then
                    SQLStringAND:= '(('+FieldBusca+ ' = FALSE)'
                  else
                    SQLStringAND:= '(('+FieldBusca+Sinal+QuotedStr(tudo+ListaFiltroAND[j]+tudo)+')';
                end
                else
                begin
                  if ((ListaFiltroAND[j] = 'Vazia')AND(Operador = 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' <> "")'
                  else if ((ListaFiltroAND[j] = 'Vazia')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = "")'
                  else if ((ListaFiltroAND[j] = 'Nulo')AND(Operador = 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' IS NOT NULL)'
                  else if ((ListaFiltroAND[j] = 'Nulo')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' IS NULL)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'TRUE' then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = TRUE)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'FALSE' then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = FALSE)'
                  else
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroAND[j]+tudo)+')';
                end;
              end;
              SQLStringFINAL:= SQLStringAND+')';
            end
            else
            begin
              if ((ListaFiltroOU[i] = 'Vazia')AND(Operador = 'Diferente')) then
                SQLStringFINAL:= '(('+FieldBusca+ ' <> ""))'
              else if ((ListaFiltroOU[i] = 'Vazia')AND(Operador <> 'Diferente')) then
                SQLStringFINAL:= '(('+FieldBusca+ ' = ""))'
              else if ((ListaFiltroOU[i] = 'Nulo')AND(Operador = 'Diferente')) then
                SQLStringFINAL:= '(('+FieldBusca+ ' IS NOT NULL))'
              else if ((ListaFiltroOU[i] = 'Nulo')AND(Operador <> 'Diferente')) then
                SQLStringFINAL:= '(('+FieldBusca+ ' IS NULL))'
              else if AnsiUpperCase(ListaFiltroOU[i]) = 'TRUE' then
                SQLStringFINAL:= '(('+FieldBusca+ ' = TRUE))'
              else if AnsiUpperCase(ListaFiltroOU[i]) = 'FALSE' then
                SQLStringFINAL:= '(('+FieldBusca+ ' = FALSE))'
              else
                SQLStringFINAL:= '('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroOU[i]+tudo)+')';
            end;
          end
          else
          begin
            //Gerar lista de condições AND
            ListaFiltroAND:= TStringList.Create;
            ListaFiltroAND.delimiter:= '&';
            ListaFiltroAND.StrictDelimiter:= true;
            ListaFiltroAND.DelimitedText:= ListaFiltroOU[i];
            //======================================================
            SQLStringAND:= '';
            if ListaFiltroAND.Count > 1 then
            begin
              for j := 0 to ListaFiltroAND.Count-1 do
              begin
                if j = 0 then
                begin
                  if ((ListaFiltroAND[j] = 'Vazia')AND(Operador = 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' <> "")'
                  else if ((ListaFiltroAND[j] = 'Vazia')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' = "")'
                  else if ((ListaFiltroAND[j] = 'Null')AND(Operador = 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' IS NOT NULL)'
                  else if ((ListaFiltroAND[j] = 'Null')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= '(('+FieldBusca+ ' IS NULL)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'TRUE' then
                    SQLStringAND:= '(('+FieldBusca+ ' = TRUE)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'FALSE' then
                    SQLStringAND:= '(('+FieldBusca+ ' = FALSE)'
                  else
                    SQLStringAND:= '(('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroAND[j]+tudo)+')';
                end
                else
                begin
                  if ((ListaFiltroAND[j] = 'Vazia')AND(Operador = 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' <> "")'
                  else if ((ListaFiltroAND[j] = 'Vazia')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = "")'
                  else if ((ListaFiltroAND[j] = 'Null')AND(Operador = 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' IS NOT NULL)'
                  else if ((ListaFiltroAND[j] = 'Null')AND(Operador <> 'Diferente')) then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' IS NULL)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'TRUE' then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = TRUE)'
                  else if AnsiUpperCase(ListaFiltroAND[j]) = 'FALSE' then
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+ ' = FALSE)'
                  else
                    SQLStringAND:= SQLStringAND+'AND('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroAND[j]+tudo)+')';
                end;
              end;
              SQLStringAND:= SQLStringAND+')';
              if (Operador) = 'Diferente' then
                SQLStringFINAL:= SQLStringFINAL+'AND'+SQLStringAND
              else
                SQLStringFINAL:= SQLStringFINAL+'OR'+SQLStringAND;
            end
            else
            begin
              if ((ListaFiltroOU[i] = 'Vazia')AND(Operador = 'Diferente')) then
                SQLStringFINAL:= SQLStringFINAL+'AND('+FieldBusca+ ' <> "")'
              else if ((ListaFiltroOU[i] = 'Vazia')AND(Operador <> 'Diferente')) then
                SQLStringFINAL:= SQLStringFINAL+'OR('+FieldBusca+ ' = "")'
              else if ((ListaFiltroOU[i] = 'Null')AND(Operador = 'Diferente')) then
                SQLStringFINAL:= SQLStringFINAL+'AND('+FieldBusca+ ' IS NOT NULL)'
              else if ((ListaFiltroOU[i] = 'Null')AND(Operador <> 'Diferente')) then
                SQLStringFINAL:= SQLStringFINAL+'OR('+FieldBusca+ ' IS NULL)'
              else if AnsiUpperCase(ListaFiltroOU[i]) = 'TRUE' then
                SQLStringFINAL:= SQLStringFINAL+'OR('+FieldBusca+ ' = TRUE)'
              else if AnsiUpperCase(ListaFiltroOU[i]) = 'FALSE' then
                SQLStringFINAL:= SQLStringFINAL+'OR('+FieldBusca+ ' = FALSE)'
              else if (Operador = 'Diferente') then
                SQLStringFINAL:= SQLStringFINAL+'AND('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroOU[i]+tudo)+')'
              else
                SQLStringFINAL:= SQLStringFINAL+'OR('+FieldBusca+sinal+QuotedStr(tudo+ListaFiltroOU[i]+tudo)+')';
            end;
          end;
        end;
      end;
    end;
    if (SQLString <> '') then
      SQLStringFINAL:= SQLString+CondicionalVariaveis+'('+SQLStringFINAL+')'
    else
      SQLStringFINAL:= '('+SQLStringFINAL+')';

    Result:= SQLStringFINAL;
  end
  else
    Result:= SQLString;
end;

function TFrmPrincipal.palavraBuscaAND(FieldBusca,PalavraBusca: String): String;
  var
    ListaFiltroOU,ListaFiltroAND: TStringList;
    SQLStringAND,SQLString: String;
    i,j: Integer;
begin
  //Gerar lista de condições OU
  ListaFiltroOU:= TStringList.Create;
  ListaFiltroOU.delimiter:= ';';
  ListaFiltroOU.StrictDelimiter:= true;
  ListaFiltroOU.DelimitedText:= PalavraBusca;
  SQLString:='';
  //======================================================
  if ListaFiltroOU.Count > 0 then
  begin
    for i := 0 to ListaFiltroOU.Count-1 do
    begin
      if i = 0 then
      begin
        //Gerar lista de condições AND
        ListaFiltroAND:= TStringList.Create;
        ListaFiltroAND.delimiter:= '&';
        ListaFiltroAND.StrictDelimiter:= true;
        ListaFiltroAND.DelimitedText:= ListaFiltroOU[i];
        SQLStringAND:= '';
        //======================================================
        if ListaFiltroAND.Count > 1 then
        begin
          for j := 0 to ListaFiltroAND.Count-1 do
          begin
            if j = 0 then
            begin
              SQLStringAND:= '(('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroAND[j]+'%')+')';
            end
            else
            begin
              SQLStringAND:= SQLStringAND+'AND('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroAND[j]+'%')+')';
            end;
          end;
          SQLString:= SQLStringAND+')';
        end
        else
          SQLString:= '('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroOU[i]+'%')+')';
      end
      else
      begin
        //Gerar lista de condições AND
        ListaFiltroAND:= TStringList.Create;
        ListaFiltroAND.delimiter:= '&';
        ListaFiltroAND.StrictDelimiter:= true;
        ListaFiltroAND.DelimitedText:= ListaFiltroOU[i];
        //======================================================
        SQLStringAND:= '';
        if ListaFiltroAND.Count > 1 then
        begin
          for j := 0 to ListaFiltroAND.Count-1 do
          begin
            if j = 0 then
            begin
              SQLStringAND:= '(('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroAND[j]+'%')+')';
            end
            else
            begin
              SQLStringAND:= SQLStringAND+'AND('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroAND[j]+'%')+')';
            end;
          end;
          SQLStringAND:= SQLStringAND+')';
          SQLString:= SQLString+'OR'+SQLStringAND;
        end
        else
          SQLString:= SQLString+'OR('+FieldBusca+' LIKE '+QuotedStr('%'+ListaFiltroOU[i]+'%')+')';
      end;
    end;
  end;
  Result:= '('+SQLString+')';
end;

procedure TFrmPrincipal.Programaodiria1Click(Sender: TObject);
begin
  if not Assigned(FrmConsultaProgramacao) then
    FrmConsultaProgramacao:= TFrmConsultaProgramacao.Create(Application)
  else
    FrmConsultaProgramacao.Show;
end;

procedure TFrmPrincipal.ProgramarExecutantes(Origem: String;idProgramacaoDiaria: Integer;
QueryExecutante,QueryProgramacao: TADOQuery; DataSourceExecutante,DataSourceProgramacao: TDataSource);
begin
  if idProgramacaoDiaria <> 0 then
  begin
    try
      QueryExecutante.Insert;
      DataSourceExecutante.DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger:=idProgramacaoDiaria;
      DataSourceExecutante.DataSet.FieldByName('Origem').AsString:= Origem;
      DataSourceExecutante.DataSet.FieldByName('StatusProgramacao').AsString:= 'Aprovado';
      QueryExecutante.Post;
      //==================================================================
      QueryProgramacao.Edit;
      DataSourceProgramacao.DataSet.FieldByName('NumExecutantes').AsInteger:=
      DataSourceProgramacao.DataSet.FieldByName('NumExecutantes').AsInteger+1;
      DataSourceProgramacao.DataSet.FieldByName('NumAprovados').AsInteger:=
      DataSourceProgramacao.DataSet.FieldByName('NumAprovados').AsInteger+1;
      QueryProgramacao.Post;
      //==================================================================
    except
      QueryExecutante.Cancel;
      QueryProgramacao.Cancel;
    end;
  end;
end;

procedure TFrmPrincipal.ProgramacaoDiaria2Click(Sender: TObject);
begin
  if not Assigned(FrmProgramacaoDiaria) then
    FrmProgramacaoDiaria:= TFrmProgramacaoDiaria.Create(Application)
  else
    FrmProgramacaoDiaria.Show;
end;

procedure TFrmPrincipal.ProgressBarAtualizar;
begin
  // Atualizar ProgressBar
  ProgressBarPrincipal.Value := ProgressBarPrincipal.Max;
  ProgressBarPrincipal.Repaint;
  Timer1.Enabled:= true;
  Screen.Cursor := crDefault;
  MenssagemStatus('');
  StatusBarPrincipal.Panels[5].Text:= '';
  StatusBarPrincipal.Panels[5].Width:= StatusBarPrincipal.Canvas.TextWidth
  (StatusBarPrincipal.Panels[5].Text)+20;
  StatusBarPrincipal.Repaint;
end;

procedure TFrmPrincipal.ProgressBarIncializa(MaxPosition: integer;
  Texto: String);
begin
  txtBarraProgresso:= Texto;
  if MaxPosition <> 0 then
    ProgressBarPrincipal.Max := MaxPosition
  else
    ProgressBarPrincipal.Max := 1;

  ProgressBarPrincipal.Min := 0;
  ProgressBarPrincipal.Value := 0;
  StatusBarPrincipal.Repaint;
  MenssagemStatus(Texto);
  Screen.Cursor := crHourGlass;
end;

procedure TFrmPrincipal.ProgressBarIncremento(Incremento: integer);
begin
  ProgressBarPrincipal.Value := ProgressBarPrincipal.Value + Incremento;
  StatusBarPrincipal.Panels[5].Text:=
  FormatFloat('0',ProgressBarPrincipal.Value)+' de '+FormatFloat('0',ProgressBarPrincipal.Max);
  StatusBarPrincipal.Panels[5].Width:= StatusBarPrincipal.Canvas.TextWidth(StatusBarPrincipal.Panels[5].Text)+15;
  StatusBarPrincipal.Panels[4].Text:= txtBarraProgresso;
  StatusBarPrincipal.Panels[4].Width:= StatusBarPrincipal.Canvas.TextWidth(StatusBarPrincipal.Panels[4].Text)+15;
  StatusBarPrincipal.Repaint;
end;

function TFrmPrincipal.RefToCell(const iCol, iRow: Integer): String;
var
  ACount, APos: Integer;
begin
  ACount := iCol div 26;
  APos := iCol mod 26;
  if APos = 0 then
  begin
    ACount := ACount - 1;
    APos := 26;
  end;

  if ACount = 0 then
    Result := Chr(Ord('A') + iCol - 1) + IntToStr(iRow)
  else if ACount = 1 then
    Result := 'A' + Chr(Ord('A') + APos - 1) + IntToStr(iRow)
  else
    Result := Chr(Ord('A') + ACount - 1) + Chr(Ord('A') + APos - 1) + IntToStr(iRow);
end;

function TFrmPrincipal.registroEndereco(keyString: String): String;
  var
    Registro : TRegistry;
begin
  // Chama o construtor do objeto
  Registro := TRegistry.Create;
  with Registro do
  begin
   RootKey := HKEY_CURRENT_USER;
   // Somente abre se a chave existir
   if OpenKey ('Colibri', False) then
     // Envia as informações ao form, vendo se os valores existem, primeiramente...
     if ValueExists (keyString) then
       Result:= ReadString(keyString);
   // Fecha a chave e o objeto
   Registro.CloseKey;
   Registro.Free;
  end;
end;

procedure TFrmPrincipal.registroEscrever(keyString, Endereco: String);
  var
    Registro : TRegistry;
Begin
  Registro := TRegistry.Create;
  try
    with Registro do
    begin
      RootKey := HKEY_CURRENT_USER;
      OpenKey('Colibri', True);
      WriteString(keyString, Endereco); //Nome dado ao arquivo de sua aplicação.
      CloseKey;
    end;
  finally
    Registro.CloseKey;
    Registro.Free;
  end;
end;

procedure TFrmPrincipal.ResumoProgramao1Click(Sender: TObject);
begin
  if not Assigned(FrmAgendaIntervencao) then
    FrmAgendaIntervencao:= TFrmAgendaIntervencao.Create(Application)
  else
    FrmAgendaIntervencao.Show;
end;


function TFrmPrincipal.RetiraEnter(aText: string): string;
begin
  { Retirando as quebras de linha em campos blob }
  Result := StringReplace(aText, #$D#$A, '', [rfReplaceAll]);

  { Retirando as quebras de linha em campos blob }
  Result := StringReplace(Result, #13#10, '', [rfReplaceAll]);
end;

procedure TFrmPrincipal.rocarLogin1Click(Sender: TObject);
begin
  if not Assigned(FrmLogin) then
    FrmLogin:= TFrmLogin.Create(Application)
  else
    FrmLogin.Show;
end;

function TFrmPrincipal.RT_CentroCusto(Plataforma: String): String;
begin
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active:= false;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Parameters.Items[0].Value:= Plataforma;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active:= true;
  Result:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.FieldByName('CentroCusto').AsString;
end;

function TFrmPrincipal.RT_Local(Plataforma: String): String;
begin
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active:= false;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Parameters.Items[0].Value:= Plataforma;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active:= true;
  Result:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.FieldByName('LocalRTSAP').AsString;
end;

procedure TFrmPrincipal.SalvarBancoDadosComo1Click(Sender: TObject);
var
  arquivo_original,enderecoConsulta_original,enderecoConsulta,
  enderecoRT_original,enderecoRT: string;
begin
  // diretorio e nome do arquivo original
  arquivo_original := registroEndereco('Banco de dados');
  enderecoConsulta_original:= ExtractFilePath(arquivo_original)+'dbConsulta.mdb';
  enderecoRT_original:= ExtractFilePath(arquivo_original)+'dbRT.mdb';
  //====================================================
  SaveDialog1.Filter:= 'Arquivo Colibri|*.colibri';
  SaveDialog1.FileName:= arquivo_original;
  SaveDialog1.DefaultExt:= '*.colibri';
  if SaveDialog1.Execute then
  begin
      try
        if CopyFile(PChar(arquivo_original), PChar(SaveDialog1.FileName), false) then
        begin
          FrmPrincipal.ProgressBarIncializa(4,'Copiando banco de dados');
          registroEscrever('Banco de dados', SaveDialog1.FileName);
          FrmPrincipal.ProgressBarAtualizar;
          //Carregar banco de dados
          conectarBD(SaveDialog1.FileName,FrmDataModule.ADOConnectionColibri);
          FrmPrincipal.ProgressBarIncremento(1);
          //Banco de dados de Consulta
          enderecoConsulta:= ExtractFileDir(SaveDialog1.FileName)+'\dbConsulta.mdb';
          enderecoRT:= ExtractFileDir(SaveDialog1.FileName)+'\dbRT.mdb';

          if CopyFile(PChar(enderecoConsulta_original), PChar(enderecoConsulta), false) then
            FrmPrincipal.conectarBDDireto(enderecoConsulta,FrmDataModule.ADOConnectionConsulta);
          FrmPrincipal.ProgressBarIncremento(1);

          if CopyFile(PChar(enderecoRT_original), PChar(enderecoRT), false) then
            FrmPrincipal.conectarBDDireto(enderecoRT,FrmDataModule.ADOConnectionRT);
          FrmPrincipal.ProgressBarIncremento(1);
        end
        else
        begin
          ShowMessage('Não foi possível copiar os arquivos');
        end;
      except
        ShowMessage('Não foi possível copiar os arquivos');
      end;
  end;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmPrincipal.selComboBox(combo: TComboBox; txt: String);
  var
    i: integer;
    S1,S2: String;
    achou: Boolean;
begin
  achou:= false;
  for I := 0 to combo.Items.Count do
  begin
    S1:= UPPERCASE(txt);
    S2:= UPPERCASE(combo.Text);
    if  S1 = S2 then
    begin
      achou:= true;
      break;
    end
    else
      combo.ItemIndex:= i;
  end;
  if achou = false then
    combo.Text:= '';
end;

procedure TFrmPrincipal.selecaoDBGrid(Grid: TDBGrid; selecao: boolean; Field: String);
  var
    selRegistro: Integer;
begin
  selRegistro:= Grid.DataSource.DataSet.RecNo;
  Grid.DataSource.Enabled:= false;
  Grid.DataSource.DataSet.First;
  while not Grid.DataSource.DataSet.Eof do
  begin
    Grid.DataSource.DataSet.Edit;
    Grid.DataSource.DataSet.FieldByName(Field).AsBoolean:= selecao;
    Grid.DataSource.DataSet.Post;
    Grid.DataSource.DataSet.Next;
  end;
  Grid.DataSource.DataSet.RecNo:= selRegistro;
  Grid.DataSource.Enabled:= true;
end;

procedure TFrmPrincipal.selecaoGrid(Grid: TStringGrid; Status: TStatusBar;
  txt: String);
  var
    i: Integer;
begin
  for I := 1 to Grid.RowCount-1 do
    Grid.Cells[0,i]:= txt;

  Status.Panels[0].Text:=
  'N° Selecionados: '+(IntToStr(FrmPrincipal.CountChecked(Grid)));
  Status.Panels[1].Text:=
  'N° Registros: '+(IntToStr(Grid.RowCount-1));
  AutoFitStatusBar(Status);
end;

procedure TFrmPrincipal.selecaoPlataforma(RChartMapa: TRChart; noPlataforma: TPointFloat;
StatusBarMapa: TStatusBar);
var
  Query: TADOQuery;
  NoSelecionado: TPointFloat;
  Plataforma,maiorX,menorX,maiorY,menorY,SQLBase: String;
  erro: Double;
begin
  Query := TADOQuery.Create(self);
  erro:= 0.55;
  FormatSettings.DecimalSeparator := '.';
  maiorX:= FloatToStr(RoundTo(noPlataforma.X+erro,-3));
  menorX:= FloatToStr(RoundTo(noPlataforma.X-erro,-3));
  maiorY:= FloatToStr(RoundTo(noPlataforma.Y+erro,-3));
  menorY:= FloatToStr(RoundTo(noPlataforma.Y-erro,-3));
  FormatSettings.DecimalSeparator := ',';
  try
    SQLBase:= 'SELECT tblPlataforma.* FROM tblPlataforma '+
    'WHERE (((CoordX <= '+maiorX+') AND (CoordX >= '+menorX+'))AND'+
    '((CoordY <= '+maiorY+') AND (CoordY >= '+menorY+')));';
    Query.Connection := FrmDataModule.ADOConnectionConsulta;
    Query.Close;
    Query.SQL.Clear;
    Query.SQL.Add(SQLBase);
    Query.Open;
    Query.Active:= true;
    if not Query.IsEmpty then
    begin
      Plataforma:= Query.FieldByName('Plataforma').AsString;
      NoSelecionado.X:= Query.FieldByName('CoordX').AsFloat;
      NoSelecionado.Y:= Query.FieldByName('CoordY').AsFloat;
      RChartMapa.DataColor:= clRed;
      RChartMapa.MarkAt(NoSelecionado.X,NoSelecionado.Y,24);
      RChartMapa.Text(NoSelecionado.X,NoSelecionado.Y,8,'  '+Plataforma);
      RChartMapa.Repaint;
      if SelPlataforma = false then
      begin
        SelPlataforma:= true;
        ultimaSelecao:= NoSelecionado;
        StatusBarMapa.Panels[0].Text:= 'Distânica entre as Plataformas:';
        StatusBarMapa.Panels[1].Text:= Plataforma;
        StatusBarMapa.Panels[2].Text:= '';
        StatusBarMapa.Panels[3].Text:= '';
        AutoFitStatusBar(StatusBarMapa);
      end
      else
      begin
        SelPlataforma:= false;
        StatusBarMapa.Panels[2].Text:= Plataforma;
        CotaReta(RChartMapa,ultimaSelecao,NoSelecionado,StatusBarMapa);
      end;
    end;
  except

  end;
end;

function TFrmPrincipal.selecionarServico(GRID: TStringGrid; ServicoPT, TextoBreveOP,
  OrdemManutencao: String): Boolean;
  var
    i: Integer;
begin
  Result:= false;
  for i := 1 to GRID.RowCount-1 do
  begin
    if ((ServicoPT = GRID.Cells[1,i])AND(TextoBreveOP = GRID.Cells[2,i])AND
    (OrdemManutencao = GRID.Cells[3,i])) then
    begin
      GRID.Cells[0,i]:= ' ';
      Result:= true;
      Break;
    end;
  end;
end;

procedure TFrmPrincipal.ServiosProgramados1Click(Sender: TObject);
begin
  if not Assigned(FrmConsultaServicosProgramados) then
    FrmConsultaServicosProgramados:= TFrmConsultaServicosProgramados.Create(Application)
  else
    FrmConsultaServicosProgramados.Show;
end;

procedure TFrmPrincipal.SetupCheckListSQL1(FieldName,txtProcura: String;CheckList: TCheckListBox;
    ListaGroup: TStringList);
var
  ListaCheck: TStringList;
  i,j: Integer;
begin
  ListaCheck:= incialListaBusca(txtProcura);
  try
    //Carregar lista group no CheckListBox
    CheckList.Items.Clear;
    for i := 0 to ListaGroup.Count-1 do
    begin
      if ListaGroup[i] = '' then
      begin
        CheckList.Items.Add('Vazia');
        CheckList.Items.Add('Nulo');
      end
      else
        CheckList.Items.Add(ListaGroup[i]);
    end;
    //Marcar já filtrados
    for i := 0 to ListaCheck.Count-1 do
      for j := 0 to CheckList.Items.Count-1 do
      begin
        if ListaCheck[i] = CheckList.Items[j] then
          CheckList.Checked[j]:= true;
      end;
  except
  end;
end;

procedure TFrmPrincipal.SetupGridPickListLista(FieldName: string; Lista: TStringList; Tabela: TDBGrid);
var
  i: Integer;
begin
  try
    for i := 0 to Tabela.Columns.Count - 1 do
      if Tabela.Columns[i].FieldName = FieldName then
      begin
        Tabela.Columns[i].PickList := Lista;
        Break;
      end;
  except
  end;
end;

procedure TFrmPrincipal.SetupGridPickListSQL(Conection: TADOConnection; const FieldName, sql: String;
  var Tabela: TDBGrid; DeletarRepetidos: Boolean);
var
  slPickList: TStringList;
  Query: TADOQuery;
  i,IndexField: Integer;
begin
  slPickList := TStringList.Create;
  Query := TADOQuery.Create(self);
  try
    Query.Connection := Conection;
    Query.sql.Text := sql;
    Query.Open;
    // Preencher o string list
    IndexField:= 0;
    for i := 0 to Query.Fields.Count-1 do
    begin
      if Query.Fields[i].FieldName = FieldName then
      begin
        IndexField:= i;
        break;
      end;
    end;
    while not Query.Eof do
    begin
      slPickList.Add(Query.Fields[IndexField].AsString);
      Query.Next;
    end; // while
    if DeletarRepetidos then
      deleteRepetidosList(slPickList);
    // colocar a lista na coluna correta
    for i := 0 to Tabela.Columns.Count - 1 do
      if Tabela.Columns[i].FieldName = FieldName then
      begin
        Tabela.Columns[i].PickList := slPickList;
        Break;
      end;
  finally
    slPickList.Free;
    Query.Free;
  end;
end;

procedure TFrmPrincipal.SetupGridFilterPickListSQL(Conection: TADOConnection; const FieldName, sql: String;
  var Tabela: TFilterDBGrid; DeletarRepetidos: Boolean);
var
  slPickList: TStringList;
  Query: TADOQuery;
  i,IndexField: Integer;
begin
  slPickList := TStringList.Create;
  Query := TADOQuery.Create(self);
  try
    Query.Connection := Conection;
    Query.sql.Text := sql;
    Query.Open;
    // Preencher o string list
    IndexField:= 0;
    for i := 0 to Query.Fields.Count-1 do
    begin
      if Query.Fields[i].FieldName = FieldName then
      begin
        IndexField:= i;
        break;
      end;
    end;
    while not Query.Eof do
    begin
      slPickList.Add(Query.Fields[IndexField].AsString);
      Query.Next;
    end; // while
    if DeletarRepetidos then
      deleteRepetidosList(slPickList);
    // colocar a lista na coluna correta
    for i := 0 to Tabela.Columns.Count - 1 do
      if Tabela.Columns[i].FieldName = FieldName then
      begin
        Tabela.Columns[i].PickList := slPickList;
        Break;
      end;
  finally
    slPickList.Free;
    Query.Free;
  end;
end;

procedure TFrmPrincipal.SituaodosEquipamentoseAcessodasPlataformas1Click(
  Sender: TObject);
begin
  if not Assigned(FrmSituacaoEquipamentoAcesso) then
    FrmSituacaoEquipamentoAcesso:= TFrmSituacaoEquipamentoAcesso.Create(Application)
  else
    FrmSituacaoEquipamentoAcesso.Show;
end;

procedure TFrmPrincipal.Sobre1Click(Sender: TObject);
begin
  if not Assigned(FrmSobre) then
    FrmSobre:= TFrmSobre.Create(Application)
  else
    FrmSobre.Show;
end;

function TFrmPrincipal.SomaHoras(Hora1, Hora2: String): String;
  var
    totalHoras,h1,h2: TTime;
begin
  try
    h1:= StrToTime(Hora1);
  except
    h1:= 0;
  end;
  try
    h2:= StrToTime(Hora2);
  except
    h2:= 0;
  end;
  totalHoras:= h1 + h2;
  Result:= TimeToStr(totalHoras);
end;

function TFrmPrincipal.SomenteNumero(Texto: String): String;
  var
    i: integer;
begin
  Result:= '';
  try
    for I := 0 to Length(Texto) do
      if (CharInSet(Texto[i],['0'..'9'])) then
        Result:= Result+Texto[i];
  except
    Result:= '';
  end;
end;

procedure TFrmPrincipal.PanelTituloAjuda1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmPrincipal.Plataforma1Click(Sender: TObject);
begin
  if not Assigned(FrmPlataforma) then
    FrmPlataforma:= TFrmPlataforma.Create(Application)
  else
    FrmPlataforma.Show;
end;

function TFrmPrincipal.PrimeiraLetraMaiscula(Str: string): string;
var
  i: integer;
  esp: boolean;
begin
  str := LowerCase(Trim(str));
  for i := 1 to Length(str) do
  begin
    if i = 1 then
      str[i] := UpCase(str[i])
    else
      begin
        if i <> Length(str) then
        begin
          esp := (str[i] = ' ');
          if esp then
            str[i+1] := UpCase(str[i+1]);
        end;
      end;
  end;
  Result := Str;
end;

procedure TFrmPrincipal.ProcuraQuery(SQLBase: String; sourceQuery: TADOQuery;StatusBar: TStatusBar);
begin
  try
    //=================================
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(SQLBase);
    sourceQuery.Open;
    StatusBar.Panels[0].Text:= 'N° Registros: '+
    IntToStr(sourceQuery.RecordCount);
    AutoFitStatusBar(StatusBar);
  except
    sourceQuery.Connection.Close;
    sourceQuery.Connection.Open;
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(SQLBase);
    sourceQuery.Open;
    StatusBar.Panels[0].Text:= 'N° Registros: '+
    IntToStr(sourceQuery.RecordCount);
    AutoFitStatusBar(StatusBar);
  end;
end;

procedure TFrmPrincipal.procuraQuerySimples(txtField, txtTabela: String;
  edtLocal: TEdit; sourceQuery: TADOQuery);
  var
    SQLString: String;
    Lista: TStrings;
    i: Integer;
begin
  Lista:=TStringList.Create;
  //Definicoes gerais
  Lista.delimiter:= '*';
  Lista.StrictDelimiter:=False;
  Lista.DelimitedText:= edtLocal.Text;
  try
    SQLString:= '';
    if (edtLocal.Text <> '') then
    begin
      for i := 0 to Lista.Count-1 do
      begin
        if i > 0 then
        begin
          SQLString:= SQLString+
          '))OR(('+txtField+' like '+QuotedStr('%'+Lista[i]+'%');
        end
        else
        begin
          SQLString:= SQLString+
          '((('+txtField+' like '+QuotedStr('%'+Lista[i]+'%');
        end;
      end;
      SQLString:= SQLString+')))';
    end;
    //=========================================================
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(
    'SELECT '+txtTabela+'.* FROM '+txtTabela+' '+
    'WHERE '+SQLString+' ORDER BY '+txtField+';');
    sourceQuery.Open;
  except
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(
    'SELECT '+txtTabela+'.* FROM '+txtTabela+' '+
    'ORDER BY '+txtField+';');
    sourceQuery.Open;
  end;
end;

procedure TFrmPrincipal.StatusBarPrincipalDrawPanel(StatusBar: TStatusBar;
  Panel: TStatusPanel; const Rect: TRect);
begin
  if Panel.Index = 6 then
  with ProgressBarPrincipal do
  begin
    Top := Rect.Top;
    Left := Rect.Left;
    Width := Rect.Right - Rect.Left - 15;
    Height := Rect.Bottom - Rect.Top;
  end;
end;

function TFrmPrincipal.SubtrairHoras(Hora1, Hora2: String): String;
  var
    totalHoras,h1,h2: TTime;
begin
  try
    h1:= StrToTime(Hora1);
  except
    h1:= 0;
  end;
  try
    h2:= StrToTime(Hora2);
  except
    h2:= 0;
  end;
  if ((h1 = 0)OR(h2 = 0)) then
    Result:= '00:00:00'
  else
  begin
    totalHoras:= ABS(h1 - h2);
    Result:= TimeToStr(totalHoras);
  end;
end;

function TFrmPrincipal.TextoMaiusculo(Texto: String): String;
  var
    ListaTexto: TStringList;
    TextoFinal: String;
    i: Integer;
begin
  try
    ListaTexto:=TStringList.Create;
    ListaTexto.delimiter:= '*';
    ListaTexto.StrictDelimiter:=false;
    ListaTexto.DelimitedText:= Texto;
    Result:= '';
    TextoFinal:= '';
    for i := 0 to ListaTexto.Count-1 do
    begin
      if i <> 0 then
        TextoFinal:= TextoFinal+' '+AnsiUpperCase(ListaTexto[i])
      else
        TextoFinal:= AnsiUpperCase(ListaTexto[i]);
    end;
    Result:= TextoFinal;
  except
    Result:= '';
  end;
end;

procedure TFrmPrincipal.Timer1Timer(Sender: TObject);
begin
  ProgressBarPrincipal.Value := ProgressBarPrincipal.Min;
  MenssagemStatus('');
  Timer1.Enabled:= false;
  StatusBarPrincipal.Repaint;
end;

function TFrmPrincipal.usuarioWindows: string;
Var
  NetUserNameLength: DWord;
Begin
  NetUserNameLength:=50;
  SetLength(Result, NetUserNameLength);
  GetUserName(pChar(Result),NetUserNameLength);
  SetLength(Result, StrLen(pChar(Result)));
end;

function TFrmPrincipal.verificaCadastroExecutante(OrigemCadastro, NomeExecutante,
  Funcao: String): Boolean;
  var
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblExecutante.* FROM tblExecutante '+
  'WHERE ((NomeExecutante LIKE '+QuotedStr(NomeExecutante)+
  ')AND(txtFuncao LIKE '+QuotedStr(Funcao)+
  ')AND(OrigemCadastro LIKE '+QuotedStr(OrigemCadastro)+'));';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Close;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBConsulta1.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBConsulta1.Open;
  if FrmDataModule.ADOQueryTemporarioDBConsulta1.IsEmpty then
    Result:= false
  else
    Result:= true;
end;

function TFrmPrincipal.VerificaCPF(Texto: String): String;
  var
    Numeros: String;
begin
  try
    Numeros:= SomenteNumero(Texto);
    if Length(Numeros) = 11 then
      Result:= FormataCPF(Numeros)
    else
      Result:= '';
  except
    Result:= '';
  end;
end;

function TFrmPrincipal.CharAscDesc(const Ch: Char; const S: string): Boolean;
var
  Posicao: integer;
begin
  Posicao := Pos(Ch, S);
  if Posicao > 0 then
    Result:= true
  else
    Result:= false;
end;

function TFrmPrincipal.verificaOperacao(StatusSistemaOP: String): String;
  var
    ListaStatusSistemaOP: TStringList;
begin
  if (StatusSistemaOP <> '') then
  begin
    ListaStatusSistemaOP:= TStringList.Create;
    //Definicoes gerais
    ListaStatusSistemaOP.delimiter:= '*';
    ListaStatusSistemaOP.StrictDelimiter:= false;
    ListaStatusSistemaOP.DelimitedText:= StatusSistemaOP;
    Result:= ListaStatusSistemaOP[0];
  end
  else
    Result:= '';
end;

function TFrmPrincipal.versaoAtualizada(versaoPRG, versaoDB: String): Integer;
  var
    charPRG,charDB: array [0..10] of char;
    i,intVersaoPRG,intVersaoDB: Integer;
    strVersaoPRG,strVersaoDB: String;
begin
  for i := 0 to Length(versaoPRG) do
    charPRG[i]:= versaoPRG[i];
  for i := 0 to Length(versaoDB) do
    charDB[i]:= versaoDB[i];

  strVersaoPRG:= charPRG[1]+charPRG[3]+charPRG[5]+charPRG[7];
  strVersaoDB:= charDB[1]+charDB[3]+charDB[5]+charDB[7];
  intVersaoPRG:= strToInt(strVersaoPRG);
  intVersaoDB:= strToInt(strVersaoDB);
  Result:= 99;
  if intVersaoPRG = intVersaoDB then
    Result:= 0
  else if intVersaoPRG < intVersaoDB then
    Result:= 1
  else if intVersaoPRG > intVersaoDB then
    Result:= 2;
end;

function TFrmPrincipal.versaoExe: String;
type
   PFFI = ^vs_FixedFileInfo;
var
   F       : PFFI;
   Handle  : Dword;
   Len     : Longint;
   Data    : Pchar;
   Buffer  : Pointer;
   Tamanho : Dword;
   Parquivo: Pchar;
   Arquivo : String;
begin
   Arquivo  := Application.ExeName;
   Parquivo := StrAlloc(Length(Arquivo) + 1);
   StrPcopy(Parquivo, Arquivo);
   Len := GetFileVersionInfoSize(Parquivo, Handle);
   Result := '';
   if Len > 0 then
   begin
      Data:=StrAlloc(Len+1);
      if GetFileVersionInfo(Parquivo,Handle,Len,Data) then
      begin
         VerQueryValue(Data, '',Buffer,Tamanho);
         F := PFFI(Buffer);
         Result := Format('%d.%d.%d.%d',
                          [HiWord(F^.dwFileVersionMs),
                           LoWord(F^.dwFileVersionMs),
                           HiWord(F^.dwFileVersionLs),
                           Loword(F^.dwFileVersionLs)]
                         );
      end;
      StrDispose(Data);
   end;
   StrDispose(Parquivo);
end;

function TFrmPrincipal.YEquacaoReta(X, X1, Y1, X2, Y2: Double): Double;
begin
  Result:= Y1*(1-((X-X1)/(X2-X1)))+Y2*(1-((X2-X)/(X2-X1)));
end;

end.
