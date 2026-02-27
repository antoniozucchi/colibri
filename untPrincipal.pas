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
  ActiveX, uAccessDBUtils, untDBGridFilter, Vcl.WinXPickers;

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
    rocarLogin1: TMenuItem;
    MenuCadastro: TMenuItem;
    Plataforma1: TMenuItem;
    Executantes1: TMenuItem;
    Embarcacao1: TMenuItem;
    ProgressBarPrincipal: TProgBar;
    N1: TMenuItem;
    CadastroUsuario1: TMenuItem;
    ProgramacaoDiaria2: TMenuItem;
    MenuImportacao: TMenuItem;
    ImportarPlanilhas1: TMenuItem;
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
    ActionManager1: TActionManager;
    actVerificaVersao: TAction;
    actAbrirDB: TAction;
    CompactarBancoDados1: TMenuItem;
    PanelAjuda1: TPanel;
    PanelTituloAjuda1: TPanel;
    SpeedButton4: TSpeedButton;
    mdiChildrenTabs: TTabSet;
    PanelCancelamentoMudanca: TPanel;
    PanelMotivoCancelamento: TPanel;
    actAjudaLimpar: TAction;
    RadioGroupFonteCancelamento: TRadioGroup;
    btnInfoPrioridade: TSpeedButton;
    PanelPrioridadeScrool: TPanel;
    ScrollBoxPrioridade: TScrollBox;
    PanelPrioridade1: TPanel;
    PanelC04: TPanel;
    Panel4: TPanel;
    DBCheckBoxInspecao: TDBCheckBox;
    edtPontoInspecao: TDBEdit;
    Panel30: TPanel;
    DBComboBoxInspecao: TDBComboBox;
    edtLimiteInspecao: TDBEdit;
    Panel34: TPanel;
    edtValorMaxInspecao: TDBEdit;
    PanelC06: TPanel;
    Panel6: TPanel;
    DBCheckBoxFO: TDBCheckBox;
    edtPontoFO: TDBEdit;
    Panel32: TPanel;
    DBComboBoxFO: TDBComboBox;
    edtLimiteFO: TDBEdit;
    Panel36: TPanel;
    edtValorMaxFO: TDBEdit;
    PanelC03: TPanel;
    Panel8: TPanel;
    DBCheckBoxCorretiva: TDBCheckBox;
    edtPontoCorretiva: TDBEdit;
    Panel28: TPanel;
    DBComboBoxCorretiva: TDBComboBox;
    edtLimiteCorretiva: TDBEdit;
    Panel29: TPanel;
    edtValorMaxCorretiva: TDBEdit;
    PanelC01: TPanel;
    Panel10: TPanel;
    DBCheckBoxPreventiva: TDBCheckBox;
    edtPontoPreventiva: TDBEdit;
    edtLimitePreventiva: TDBEdit;
    Panel50: TPanel;
    edtValorMaxPreventiva: TDBEdit;
    Panel2: TPanel;
    DBComboBoxPreventiva: TDBComboBox;
    PanelC12: TPanel;
    Panel12: TPanel;
    DBCheckBoxMarinha: TDBCheckBox;
    edtPontoMarinha: TDBEdit;
    Panel48: TPanel;
    DBComboBoxMarinha: TDBComboBox;
    edtLimiteMarinha: TDBEdit;
    Panel49: TPanel;
    edtValorMaxMarinha: TDBEdit;
    PanelC08D: TPanel;
    Panel14: TPanel;
    DBCheckBoxRTID: TDBCheckBox;
    edtPontoRTID: TDBEdit;
    Panel38: TPanel;
    DBComboBoxRTID: TDBComboBox;
    edtLimiteRTID: TDBEdit;
    Panel39: TPanel;
    edtValorMaxRTID: TDBEdit;
    PanelC07: TPanel;
    Panel16: TPanel;
    DBCheckBoxRO: TDBCheckBox;
    edtPontoRO: TDBEdit;
    Panel33: TPanel;
    DBComboBoxRO: TDBComboBox;
    edtLimiteRO: TDBEdit;
    Panel37: TPanel;
    edtValorMaxRO: TDBEdit;
    PanelC02: TPanel;
    Panel18: TPanel;
    DBCheckBoxICPMA: TDBCheckBox;
    edtPontoICPMA: TDBEdit;
    Panel44: TPanel;
    DBComboBoxICPMA: TDBComboBox;
    edtLimiteICPMA: TDBEdit;
    Panel45: TPanel;
    edtValorMaxICPMA: TDBEdit;
    PanelC13: TPanel;
    Panel20: TPanel;
    DBCheckBoxIBAMA: TDBCheckBox;
    edtPontoIBAMA: TDBEdit;
    Panel42: TPanel;
    DBComboBoxIBAMA: TDBComboBox;
    edtLimiteIBAMA: TDBEdit;
    Panel43: TPanel;
    edtValorMaxIBAMA: TDBEdit;
    PanelC11: TPanel;
    Panel22: TPanel;
    DBCheckBoxATM: TDBCheckBox;
    edtPontoATM: TDBEdit;
    Panel40: TPanel;
    DBComboBoxATM: TDBComboBox;
    edtLimiteATM: TDBEdit;
    Panel41: TPanel;
    edtValorMaxATM: TDBEdit;
    PanelC05: TPanel;
    Panel24: TPanel;
    DBCheckBoxMergulho: TDBCheckBox;
    edtPontoMergulho: TDBEdit;
    Panel31: TPanel;
    DBComboBoxMergulho: TDBComboBox;
    edtLimiteMergulho: TDBEdit;
    Panel35: TPanel;
    edtValorMaxMergulho: TDBEdit;
    PanelC16: TPanel;
    Panel26: TPanel;
    DBCheckBoxEC: TDBCheckBox;
    edtPontoEC: TDBEdit;
    Panel46: TPanel;
    DBComboBoxEC: TDBComboBox;
    edtLimiteEC: TDBEdit;
    Panel47: TPanel;
    edtValorMaxEC: TDBEdit;
    PanelC14: TPanel;
    Panel52: TPanel;
    DBCheckBoxANP: TDBCheckBox;
    edtPontoANP: TDBEdit;
    Panel53: TPanel;
    DBComboBoxANP: TDBComboBox;
    edtLimiteANP: TDBEdit;
    Panel54: TPanel;
    edtValorMaxANP: TDBEdit;
    PanelC00: TPanel;
    Panel5: TPanel;
    DBCheckBoxCusto: TDBCheckBox;
    edtPontoMinimo: TDBEdit;
    edtPontoMaximo: TDBEdit;
    edtCustoMaximo: TDBEdit;
    edtCustoMinimo: TDBEdit;
    PanelC10: TPanel;
    Panel7: TPanel;
    DBCheckBoxNR10: TDBCheckBox;
    edtPontoNR10: TDBEdit;
    Panel9: TPanel;
    DBComboBoxNR10: TDBComboBox;
    edtLimiteNR10: TDBEdit;
    Panel11: TPanel;
    edtValorMaxNR10: TDBEdit;
    PanelC09: TPanel;
    Panel21: TPanel;
    DBCheckBoxINSPLAN: TDBCheckBox;
    edtPontoINSPLAN: TDBEdit;
    Panel23: TPanel;
    DBComboBoxINSPLAN: TDBComboBox;
    edtLimiteINSPLAN: TDBEdit;
    Panel25: TPanel;
    edtValorMaxINSPLAN: TDBEdit;
    PanelC15: TPanel;
    Panel13: TPanel;
    DBCheckBoxANVISA: TDBCheckBox;
    edtPontoANVISA: TDBEdit;
    Panel15: TPanel;
    DBComboBoxANVISA: TDBComboBox;
    edtLimiteANVISA: TDBEdit;
    Panel17: TPanel;
    edtValorMaxANVISA: TDBEdit;
    Panel19: TPanel;
    Panel3: TPanel;
    Panel27: TPanel;
    Panel51: TPanel;
    Panel55: TPanel;
    PanelC08C: TPanel;
    Panel57: TPanel;
    DBCheckBoxRTIC: TDBCheckBox;
    edtPontoRTIC: TDBEdit;
    Panel58: TPanel;
    DBComboBoxRTIC: TDBComboBox;
    edtLimiteRTIC: TDBEdit;
    Panel59: TPanel;
    edtValorMaxRTIC: TDBEdit;
    PanelC08B: TPanel;
    Panel61: TPanel;
    DBCheckBoxRTIB: TDBCheckBox;
    edtPontoRTIB: TDBEdit;
    Panel62: TPanel;
    DBComboBoxRTIB: TDBComboBox;
    edtLimiteRTIB: TDBEdit;
    Panel63: TPanel;
    edtValorMaxRTIB: TDBEdit;
    PanelC08A: TPanel;
    Panel65: TPanel;
    DBCheckBoxRTIA: TDBCheckBox;
    edtPontoRTIA: TDBEdit;
    Panel66: TPanel;
    DBComboBoxRTIA: TDBComboBox;
    edtLimiteRTIA: TDBEdit;
    Panel67: TPanel;
    edtValorMaxRTIA: TDBEdit;
    ToolBar1: TToolBar;
    DBNavigatorPrioridade: TDBNavigator;
    btnConfigPadrao: TBitBtn;
    PanelMagicFiltro1: TPanel;
    PanelFiltroGRID: TPanel;
    PageControl4: TPageControl;
    TabSheet8: TTabSheet;
    TabSheet9: TTabSheet;
    ToolBar3: TToolBar;
    btnSubstituirPor: TToolButton;
    PanelLocalizar: TPanel;
    Panel56: TPanel;
    Panel60: TPanel;
    edtSubstituirPor: TEdit;
    Panel64: TPanel;
    Panel68: TPanel;
    Panel69: TPanel;
    edtLocalizar: TEdit;
    actFiltroCancelar: TAction;
    actGridSelTudo: TAction;
    actGridSelLimpa: TAction;
    PanelTituloMagic: TPanel;
    SpeedButton1: TSpeedButton;
    PanelLayout: TPanel;
    Splitter8: TSplitter;
    RLColunasAtivas: TStringGrid;
    ToolBar8: TToolBar;
    btnCima: TToolButton;
    btnBaixo: TToolButton;
    btnCimaTudo: TToolButton;
    btnBaixoTudo: TToolButton;
    btnClassifica: TToolButton;
    RLColunasOpcoes: TStringGrid;
    MemoPrincipal: TMemo;
    actLayoutSORT: TAction;
    Carregar: TAction;
    btnSalvarLayout: TToolButton;
    PanelFiltrosTabela: TPanel;
    actConfigMagicPanel: TAction;
    actConfigPanelLayout: TAction;
    actConfigFiltrosTabela: TAction;
    Panel70: TPanel;
    BitBtn3: TBitBtn;
    BitBtn4: TBitBtn;
    actExcel: TAction;
    MaskEditData: TMaskEdit;
    actCancelarProgramacao: TAction;
    StatusBar1: TStatusBar;
    PanelCONECTION: TPanel;
    ToolBar7: TToolBar;
    BitBtn7: TBitBtn;
    actLOCAL: TAction;
    actREDE: TAction;
    BitBtn8: TBitBtn;
    actUpload: TAction;
    actDownload: TAction;
    BitBtn9: TBitBtn;
    BitBtn10: TBitBtn;
    ConexoLOCAL1: TMenuItem;
    actSalvatagemPlataforma: TAction;
    actConverterDB: TAction;
    Converterverso1: TMenuItem;
    actConexao: TAction;
    MovimentacaoCarga1: TMenuItem;
    actSalvarLista: TAction;
    actInterromper: TAction;
    ProgBarMagicFiltro: TProgBar;
    Timer2: TTimer;
    actListar: TAction;
    StatusBarMagicFiltro: TStatusBar;
    CondicaoMarEmbarcacao1: TMenuItem;
    PanelDuplicados: TPanel;
    ToolBar6: TToolBar;
    BitBtn13: TBitBtn;
    RLDuplicados: TStringGrid;
    actExcelDuplicados: TAction;
    GroupBox1: TGroupBox;
    CheckBoxColibri: TCheckBox;
    CheckBoxConsulta: TCheckBox;
    CheckBoxMemoria: TCheckBox;
    Panel80: TPanel;
    ToolBar2: TToolBar;
    btnASC: TToolButton;
    btnDESC: TToolButton;
    btnSelTudo: TToolButton;
    btnSelLimpa: TToolButton;
    ToolButton1: TToolButton;
    RadioGroupOperador: TRadioGroup;
    Panel81: TPanel;
    edtProcura: TEdit;
    CheckListBoxFiltroGRID: TCheckListBox;
    btnInserir: TBitBtn;
    btnCancelar: TBitBtn;
    BitBtn12: TBitBtn;
    BitBtn11: TBitBtn;
    CheckBoxListarAutomatico: TCheckBox;
    N3: TMenuItem;
    actDownlodaDBMemoria: TAction;
    actUploadDBMemoria: TAction;
    RadioGroupSubstituir: TRadioGroup;
    DBGridPalavraChave: TDBGrid;
    ColunasLayoutPalavraChave: TStringGrid;
    actFiltroInserir: TAction;
    actGridASC: TAction;
    actGridDESC: TAction;
    actSubstituirPor: TAction;
    actLimparFiltros: TAction;
    actFiltrosTabela: TAction;
    actProcuraFiltrosTabela: TAction;
    actProcurarPalavraChave: TAction;
    ToolBar9: TToolBar;
    BitBtn5: TBitBtn;
    BitBtn6: TBitBtn;
    DBNavigatorPalavraChave: TDBNavigator;
    StatusBarPalavraChave: TStatusBar;
    BitBtn14: TBitBtn;
    BitBtn15: TBitBtn;
    actInserirRegistro: TAction;
    actCancelarRegistro: TAction;
    Ajuda1: TMenuItem;
    actMatrizExecutanteAPLAT: TAction;
    actMatrizExecutanteT31: TAction;
    actMatrizExecutanteCadastro: TAction;
    ResumoProgramao1: TMenuItem;
    SituaodosEquipamentoseAcessodasPlataformas1: TMenuItem;
    actMatrizForaOperacao: TAction;
    Splitter1: TSplitter;
    RLFiltrosTabela: TStringGrid;
    ComboBoxOperador: TComboBox;
    ToolBar4: TToolBar;
    btnProcurarFiltrosTabela: TBitBtn;
    Panel72: TPanel;
    ToolBar10: TToolBar;
    btnSQLFiltroTabela: TToolButton;
    MemoSQL: TMemo;
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
    procedure ImportarPlanilhas1Click(Sender: TObject);
    procedure Sobre1Click(Sender: TObject);
    procedure GerenciarSolicitaes1Click(Sender: TObject);
    procedure GerenciarTransportes1Click(Sender: TObject);
    procedure ServiosProgramados1Click(Sender: TObject);
    procedure ExecutantesProgramadosporPerodo1Click(Sender: TObject);
    procedure ControledeGeradores1Click(Sender: TObject);
    procedure actVerificaVersaoExecute(Sender: TObject);
    procedure actAbrirDBExecute(Sender: TObject);
    procedure CompactarBancoDados1Click(Sender: TObject);
    procedure PanelTituloAjuda1MouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PanelTituloAjuda1MouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PanelTituloAjuda1MouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButton4Click(Sender: TObject);
    procedure actAjudaLimparExecute(Sender: TObject);
    procedure DBComboBoxPreventivaKeyPress(Sender: TObject; var Key: Char);
    procedure DBCheckBoxPreventivaClick(Sender: TObject);
    procedure DBCheckBoxICPMAClick(Sender: TObject);
    procedure DBCheckBoxCorretivaClick(Sender: TObject);
    procedure DBCheckBoxInspecaoClick(Sender: TObject);
    procedure DBCheckBoxMergulhoClick(Sender: TObject);
    procedure DBCheckBoxFOClick(Sender: TObject);
    procedure DBCheckBoxROClick(Sender: TObject);
    procedure DBCheckBoxRTIDClick(Sender: TObject);
    procedure DBCheckBoxMarinhaClick(Sender: TObject);
    procedure DBCheckBoxATMClick(Sender: TObject);
    procedure DBCheckBoxIBAMAClick(Sender: TObject);
    procedure DBCheckBoxANPClick(Sender: TObject);
    procedure DBCheckBoxECClick(Sender: TObject);
    procedure DBCheckBoxCustoClick(Sender: TObject);
    procedure edtCustoMinimoExit(Sender: TObject);
    procedure edtCustoMaximoExit(Sender: TObject);
    procedure DBComboBoxPreventivaChange(Sender: TObject);
    procedure DBComboBoxICPMAChange(Sender: TObject);
    procedure DBComboBoxCorretivaChange(Sender: TObject);
    procedure DBComboBoxInspecaoChange(Sender: TObject);
    procedure DBComboBoxMergulhoChange(Sender: TObject);
    procedure DBComboBoxFOChange(Sender: TObject);
    procedure DBComboBoxROChange(Sender: TObject);
    procedure DBComboBoxRTIDChange(Sender: TObject);
    procedure DBComboBoxMarinhaChange(Sender: TObject);
    procedure DBComboBoxATMChange(Sender: TObject);
    procedure DBComboBoxIBAMAChange(Sender: TObject);
    procedure DBComboBoxANPChange(Sender: TObject);
    procedure DBComboBoxECChange(Sender: TObject);
    procedure DBComboBoxNR10Change(Sender: TObject);
    procedure DBCheckBoxNR10Click(Sender: TObject);
    procedure DBCheckBoxINSPLANClick(Sender: TObject);
    procedure DBCheckBoxANVISAClick(Sender: TObject);
    procedure DBComboBoxANVISAChange(Sender: TObject);
    procedure DBComboBoxINSPLANChange(Sender: TObject);
    procedure DBCheckBoxRTIAClick(Sender: TObject);
    procedure DBCheckBoxRTIBClick(Sender: TObject);
    procedure DBCheckBoxRTICClick(Sender: TObject);
    procedure DBComboBoxRTIAChange(Sender: TObject);
    procedure DBComboBoxRTIBChange(Sender: TObject);
    procedure DBComboBoxRTICChange(Sender: TObject);
    procedure actFiltroCancelarExecute(Sender: TObject);
    procedure actGridSelTudoExecute(Sender: TObject);
    procedure actGridSelLimpaExecute(Sender: TObject);
    procedure SpeedButton1Click(Sender: TObject);
    procedure PanelTituloMagicMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PanelTituloMagicMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure actLayoutSORTExecute(Sender: TObject);
    procedure RLColunasAtivasDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure RLColunasOpcoesDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure RLColunasAtivasDblClick(Sender: TObject);
    procedure RLColunasOpcoesDblClick(Sender: TObject);
    procedure actConfigMagicPanelExecute(Sender: TObject);
    procedure actConfigPanelLayoutExecute(Sender: TObject);
    procedure actConfigFiltrosTabelaExecute(Sender: TObject);
    procedure RLFiltrosTabelaSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ComboBoxOperadorCloseUp(Sender: TObject);
    procedure ComboBoxOperadorMouseLeave(Sender: TObject);
    procedure ComboBoxOperadorKeyPress(Sender: TObject; var Key: Char);
    procedure edtProcuraKeyPress(Sender: TObject; var Key: Char);
    procedure BitBtn4Click(Sender: TObject);
    procedure CarregarExecute(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure actCancelarProgramacaoExecute(Sender: TObject);
    procedure ConexoLOCAL1Click(Sender: TObject);
    procedure actLOCALExecute(Sender: TObject);
    procedure actREDEExecute(Sender: TObject);
    procedure actUploadExecute(Sender: TObject);
    procedure actDownloadExecute(Sender: TObject);
    procedure actSalvatagemPlataformaExecute(Sender: TObject);
    procedure actConverterDBExecute(Sender: TObject);
    procedure MovimentacaoCarga1Click(Sender: TObject);
    procedure actInterromperExecute(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure actListarExecute(Sender: TObject);
    procedure CondicaoMarEmbarcacao1Click(Sender: TObject);
    procedure actSalvarListaExecute(Sender: TObject);
    procedure actExcelDuplicadosExecute(Sender: TObject);
    procedure RLDuplicadosFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure actDownlodaDBMemoriaExecute(Sender: TObject);
    procedure actUploadDBMemoriaExecute(Sender: TObject);
    procedure DBGridPalavraChaveDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure DBGridPalavraChaveTitleClick(Column: TColumn);
    procedure actFiltroInserirExecute(Sender: TObject);
    procedure actGridASCExecute(Sender: TObject);
    procedure actGridDESCExecute(Sender: TObject);
    procedure actSubstituirPorExecute(Sender: TObject);
    procedure actLimparFiltrosExecute(Sender: TObject);
    procedure actFiltrosTabelaExecute(Sender: TObject);
    procedure actProcuraFiltrosTabelaExecute(Sender: TObject);
    procedure actProcurarPalavraChaveExecute(Sender: TObject);
    procedure DBGridPalavraChaveCellClick(Column: TColumn);
    procedure actInserirRegistroExecute(Sender: TObject);
    procedure actCancelarRegistroExecute(Sender: TObject);
    procedure Ajuda1Click(Sender: TObject);
    procedure actMatrizExecutanteAPLATExecute(Sender: TObject);
    procedure actMatrizExecutanteT31Execute(Sender: TObject);
    procedure actMatrizExecutanteCadastroExecute(Sender: TObject);
    procedure ResumoProgramao1Click(Sender: TObject);
    procedure SituaodosEquipamentoseAcessodasPlataformas1Click(Sender: TObject);
    procedure actMatrizForaOperacaoExecute(Sender: TObject);
    type
      TExecutanteProgramado = record
        ListaDestinos: TStrings;
        booleanProgramado: Boolean;
    end;

  private
    { Private declarations }
    function versaoAtualizada(versaoPRG,versaoDB: String): Integer;
    function versaoExe: String;
    function HintRelacaoTempo(Combo: TDBComboBox): String;

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

    MatrizSalvatagem,MatrizExecutanteAPLAT,MatrizForaOperacao,
    MatrizExecutanteT31,MatrizExecutanteCadastro: array of array of String;
    const
    CoresAuto: array[0..14] of TColor =
    (clMaroon,clGreen,clOlive,clNavy,clPurple,clTeal,clRed,clLime,clYellow,clBlue,
    clFuchsia,clAqua,clBlack,clGray,clSilver);
    { Public declarations }
    //Registros Windows
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
    procedure GridZebrado(Grid: TDBGrid; ColunasLayout: TStringGrid; State: TGridDrawState;
    const Rect: TRect; DataCol: Integer; Column: TColumn);
    procedure SetupGridPickListSQL(Conection: TADOConnection; const FieldName: string; const sql: String;
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
    procedure selecaoDBGrid(Grid: TDBGrid; selecao: boolean);
    function palavraBusca(SQLString,FieldBusca,Operador,StringBusca,CondicionalVariaveis: String): String;
    function palavraBuscaAND(FieldBusca,PalavraBusca: String): String;
    function incialListaBusca(CampoString: String): TStringList;
    procedure LayoutPadrao(NomeArquivoMemo:string; GridPadrao: TStringGrid;Tabela: String);
    procedure AutoFitGrade(aGrade: TStringGrid);
    procedure AutoFitStatusBar(aStatus: TStatusBar);
    procedure addColuna(Grid: TDBGrid; Field,Caption,alinha: String;
    numColuna,Largura: Integer; booleanReadOnly: Boolean);
    procedure titleGrid(ColunasLayout: TStringGrid; FonteDB,SQLQuery: String);
    procedure inserirProcura(Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure SubstituirPor(Grid: TDBGrid; ADOQuery: TADOQuery; DataSource: TDataSource);
    procedure limparTitleGrid(Grid: TDBGrid);
    procedure buscaFiledGrid1(FieldName,palavraBusca,Operador:String;ColunasLayout:
    TStringGrid; ACol: Integer;BooleanConcatenar: Boolean);
    procedure AlterarTituloColuna(ColunasLayout:TStringGrid;FieldName,strTitulo: String);
    function GroupFieldDBGrid(SQLQuery,FieldName,FonteDB: String; StatusBar: TStatusBar): TStringList;
    function indexLayoutFieldName(FieldName: String; ColunasLayout: TStringGrid): Integer;
    function indexLayoutCaption(Caption: String; ColunasLayout: TStringGrid): Integer;
    //Tratamento de Strings
    function DeleteChar(const Ch: Char; const S: string): string;
    //Excel
    procedure GerarExcel(Tabela: TDBGrid; const nomeTabela: String);
    procedure ExportarExcel(TableName: String; BancoDados: Integer);
    procedure GeraTexto(DataSource: TDataSource; Separador: Char; nomeTabela: String);
    function RefToCell(const iCol: Integer; const iRow: Integer): String;
    procedure importarExcel(nomeTabela: String;
    ACol,ARow,TabSheet: Integer; Tabela: TDBGrid; ADOQuery: TADOQuery);
    procedure ExcelStringGrid (Grid: TStringGrid; NomePlanilha,Titulo1,Titulo2: String;LinhaInicial: Integer);
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
    procedure GravarCanceladoAprovado(idProgramacaoDiaria: Integer);
    procedure compactarDB(EnderecoArquivo: String;booleanPerguntar,conexaoColibri: Boolean;
    SourceADOConnection: TADOConnection);
    procedure compactarDBMemoria(EnderecoArquivo: String;SourceADOConnection: TADOConnection);
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
    function  GetDistanceBetween(lat1,long1,lat2,long2: Double) : Double;
    function LatLong_XY(Latitude,Longitude: Double): TPointFloat;
    function isNumeric(txt: String): Integer;
    function isData(txt: String): Boolean;
    Function usuarioWindows: string;
    function NomeMaquina: String;
    function TextoMaiusculo(Texto: String): String;
    function dataFormatar(SourceData: String): TDateTime;
    function dataSAP(SourceData: String): String;
    function DataSAP_ToDate(const SourceData: string): TDateTime;
    function dataLimpa(SourceData: String): String;
    function CalcNumCanceladosAprovado(idProgramacaoDiaria,Tipo: Integer): Integer;
    function CalcNumExecutantes(idProgramacaoDiaria: Integer): Integer;
    procedure configPiroridade(edtPonto,edtLimite,edtValorMax: TDBEdit;
    relacaoData: TDBComboBox; check: TDBCheckBox);
    procedure AvaliarProgramacaoExecutante(idProgramacaoExecutante,Fonte: Integer;
    StatusProgramacao,StatusExecucao,Motivo: String);
    procedure RT_ProgramacaoExecutante(idProgramacaoExecutante: Integer;
    RT, TipoEmbarque: String);
    function SQLStringFiltroTabela(ColunasLayout: TStringGrid;BooleanWHERE: Boolean): String;
    procedure ProcuraQuery(SQLBase: String; sourceQuery: TADOQuery;StatusBar: TStatusBar);

    function verificaOperacao(StatusSistemaOP: String): String;
    //Layout Grid
    procedure CarregarColunasGRID(Conection: TADOConnection; tblTabela: String; Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure GravarRLColunas(NomeArquivoMemo:string);
    procedure CarregarRLColunas(NomeArquivoMemo: String; ColunasLayout: TStringGrid);
    procedure CarregarRLColunasAtivasGRID(Grid: TDBGrid);
    procedure moveCima(Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure moveBaixo(Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure botaoMoveTudo(Conection: TADOConnection; NomeArquivoMemo,NomeTabela,NumRLTAB: String;Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure SetUpColunasLayout(Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure LimparColunasFiltro(Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure FiltrosTabela(Grid: TDBGrid; ColunasLayout: TStringGrid);
    //PDF
    procedure ImageExcluir(NomeArquivo: String);
    function dadosColuna(txtCaption: String; ColunasLayout: TStringGrid): TStringList;
    procedure SalvarRLColunasCarregaGRID(Conection: TADOConnection;
    NomeArquivoMemo,NomeTabela: String; Grid: TDBGrid; ColunasLayout: TStringGrid);
    procedure CarregaFiltrosProcura(ColunasLayout: TStringGrid);
    function DiasFinalUtil(DataInicio: TDateTime; DiasUteis: Integer):TDateTime;
    function DiasFinalCorridos(DataInicio: TDateTime; DiasCorridos: Integer):TDateTime;
    function verificaTextoLongoCarteira(OrdemManutencao: String): WideString;
    procedure configurarFiltro(Tipo: Integer;fieldName,ColunaIndex: String;
    ColunaReadOnly: Boolean; actFiltroInserir,actGridASC,actGridDESC,
    actSubstituirPor:TAction);
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
    procedure CopiarProgramacao(DataProgramacao: String;DataSource: TDataSource);
    procedure ProgramarExecutante(Origem: String;idProgramacaoDiaria: Integer;
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

    {procedure ExcluirTableDB(NomeTabela: String; sourceQuery: TADOQuery);
    procedure CriarTableDB(NomeTabela,NomeID,SQLVariaveis: String; sourceQuery: TADOQuery);
    procedure CriarRelacaoTable(TabelaPai,TabelaFilho,IDPai,IDFilho: String; sourceQuery: TADOQuery);
    procedure AlterarNomeFieldDB(CampoNomeNovo,CampoNomeOriginal,NomeTabela,Tipo: String; sourceQuery: TADOQuery);
    procedure AlterarNomeTableDB(TabelaNomeNovo,TabelaNomeOriginal: String; sourceQuery: TADOQuery);}

    function RetiraEnter(aText : string): string;
    procedure ProgramacaoDuplicados(RLGrid: TStringGrid; DataProcura: String);
    procedure PainelDuplicados(DataProcura: String);
    function OrigemPlataformas: String;
    procedure configurarCheckListBox(CheckListBox: TCheckListBox;DBGrid: TDBGrid;Layout: TStringGrid;BooleanLimparColunas: Boolean);
    function buscarString(strBusca,strTexto: String): Boolean;
    procedure DesenharCalendario(Calendario: TStringGrid; PanelNome: TPanel;PrimeiroDiaMes: TDateTime);
    function PrimeiraLetraMaiscula(Str: string): string;
    procedure AbrirBancoDados(enderecoColibri,Chave: String;BooleanSplash: Boolean);
    procedure PainelCancelamentoMudanca(Fonte: Integer; Titulo: String);
    procedure criarVariavelTempoExecucao(Field,Tabela,TipoField: String; sourceQuery: TADOQuery);
    procedure msgSplash(texto1,texto2: String);
    function ForaOperacao(Plataforma: String; IndexSistema: Integer): String;
  end;

var
  FrmPrincipal: TFrmPrincipal;


implementation
  uses untDataModule, untConsultaProgramacao,untFrmlogin,
  untEmbarcacao,untExecutante, untPlataforma, untFrmCadastroUsuario, untProgramacaoDiaria,
  untImportacao,untFrmSobre, untGerenciarSolicitacoes,
  untGerenciarEmbarcacoes, untConsultaServicosProgramados,
  untConsultaExecutantesProgramados, untControleGeradores,
  untAgendaIntervencao,untMovimentacaoCarga,untTelaAbertura,
  untSituacaoEquipamentoAcesso,
  untCondicaoEmbarcacao;

{$R *.dfm}

{ TFrmPrincipal }

procedure TFrmPrincipal.AbrirBancoDados(enderecoColibri,Chave: String;BooleanSplash: Boolean);
  var
    enderecoConsulta,enderecoMemoria: String;
begin
  try
    enderecoColibri:= conectarBD(enderecoColibri,FrmDataModule.ADOConnectionColibri);
    enderecoConsulta:= ExtractFilePath(enderecoColibri)+'dbConsulta.mdb';
    enderecoMemoria:= ExtractFilePath(enderecoColibri)+'dbMemoria.mdb';
    carregarLoginUsuario(Chave);
    //========================================================================
    if BooleanSplash then
    begin
      msgSplash('Conectando ao banco de dados dbColibri.colibri',
      ' no endereço: '+enderecoColibri);

      try
        msgSplash('Conectando ao dbConsulta.mdb',
        ' no endereço: '+enderecoConsulta);
        FrmPrincipal.conectarBDDireto(enderecoConsulta,FrmDataModule.ADOConnectionConsulta);
      except
        MessageBox(0,PChar('Banco de dados de Consulta não encontrado ou corrompido: '+enderecoConsulta),
        'Colibri',MB_ICONEXCLAMATION);
      end;
      try
        msgSplash('Conectando ao dbMemoria.mdb',
        ' no endereço: '+enderecoMemoria);
        FrmPrincipal.conectarBDDireto(enderecoMemoria,FrmDataModule.ADOConnectionMemoria);
      except
        MessageBox(0,PChar('Banco de dados de Memória não encontrado ou corrompido: '+enderecoMemoria),
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
  end;
end;

procedure TFrmPrincipal.actAjudaLimparExecute(Sender: TObject);
begin
  PanelCancelamentoMudanca.Visible:= false;
  PanelPrioridadeScrool.Visible:= false;
  PanelFiltrosTabela.Visible:= false;
  PanelLayout.Visible:= false;
  PanelCONECTION.Visible:= false;
  PanelDuplicados.Visible:= false;
  PanelAjuda1.Visible:= false;
  btnInfoPrioridade.Visible:= false;
  PanelTituloAjuda1.Color:= clGradientActiveCaption;
end;

procedure TFrmPrincipal.CarregarExecute(Sender: TObject);
begin
  FrmPrincipal.btnSalvarLayout.Click;
end;

procedure TFrmPrincipal.actCancelarProgramacaoExecute(Sender: TObject);
begin
  PanelAjuda1.Visible:= false;
  PanelTituloAjuda1.Color:= clGradientActiveCaption;
end;

procedure TFrmPrincipal.actCancelarRegistroExecute(Sender: TObject);
begin
  PanelAjuda1.Visible:= false;
end;

procedure TFrmPrincipal.actConfigFiltrosTabelaExecute(Sender: TObject);
begin
  actAjudaLimpar.Execute;
  PanelFiltrosTabela.Visible:= false;
  PanelFiltrosTabela.Align:= alClient;
  PanelFiltrosTabela.Visible:= true;
  PanelAjuda1.Width:= 600;
  PanelAjuda1.Height:= 500;
  PanelTituloAjuda1.Caption:= 'Filtros da tabela';
  PanelAjuda1.Visible:= true;
  btnProcurarFiltrosTabela.Caption:= 'Procurar';
  ImageList1.GetBitmap(28,btnProcurarFiltrosTabela.Glyph);
  btnProcurarFiltrosTabela.Hint:= 'Procurar';
  RLFiltrosTabela.Hint:= HintPadrao;
end;

procedure TFrmPrincipal.actConfigMagicPanelExecute(Sender: TObject);
begin
  ImageList1.GetBitmap(27,btnInserir.Glyph);
  btnInserir.Caption:= 'Filtrar';
  btnASC.ImageIndex:= 86;
  btnDESC.ImageIndex:= 85;
  btnSubstituirPor.ImageIndex:= 59;
  //===============================
  btnInserir.Hint:= 'Inserir filtro';
  btnASC.Hint:= 'Crescente';
  btnDESC.Hint:= 'Decrescente';
  btnSubstituirPor.Hint:= 'Substituir Por';
  edtProcura.Hint:= 'Localizar';
  PanelMagicFiltro1.Width := 400;
  PanelMagicFiltro1.Height := 387;
end;

procedure TFrmPrincipal.actConfigPanelLayoutExecute(Sender: TObject);
begin
  btnCima.ImageIndex:= 125;
  btnBaixo.ImageIndex:= 126;
  btnCimaTudo.ImageIndex:= 464;
  btnBaixoTudo.ImageIndex:= 465;
  //===============================
  btnCima.Hint:= 'Mover para Colunas Ativas';
  btnBaixo.Hint:= 'Mover para Colunas Optativas';
  btnCimaTudo.Hint:= 'Mover todos para Colunas Ativas';
  btnBaixoTudo.Hint:= 'Mover todos para Colunas Optativas';
  //===============================
  FrmPrincipal.actAjudaLimpar.Execute;
  FrmPrincipal.PanelLayout.Visible:= true;
  FrmPrincipal.PanelLayout.Align:= alClient;
  FrmPrincipal.PanelTituloAjuda1.Caption:= 'Layout';
  FrmPrincipal.PanelAjuda1.Width:= 250;
  FrmPrincipal.PanelAjuda1.Height:= 450;
  FrmPrincipal.PanelAjuda1.Visible:= true;
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
      CriarFieldDB('AtualizadoPor','tblNotaManutencao','VARCHAR(10)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      CriarFieldDB('DataAtualizacao','tblNotaManutencao','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      CriarFieldDB('ComputadorAtualizacao','tblNotaManutencao','VARCHAR(20)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
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
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblCarteiraTrabalho','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblExecutanteAPLAT','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblExecutanteSAP','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblLocalInstalacao','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblNotaManutencao','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblPreventivas','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('ComputadorImportacao','Computador','tblLocalInstalacao','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('TextoBreveOP','TextoBreveOperacao','tblPreventivas','VARCHAR(20)',
      FrmDataModule.ADOConnectionMemoria);
      CriarFieldDB('DataAtendimento','tblMovimentacaoCarga','DATETIME',
      FrmDataModule.ADOQueryTemporarioDBConsulta1);
      AlterarNomeFieldDB('DataNecessidade','Data','tblMovimentacaoCarga','DATETIME',
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
      'tblProgramacaoExecutante','VARCHAR(255)',FrmDataModule.ADOConnectionColibri);
      AlterarNomeFieldDB('PalavraChaveProgramacao','PalavraChaveCancelamento',
      'tblProgramacaoExecutante','VARCHAR(255)',FrmDataModule.ADOConnectionColibri);
      versaoDB:= '1.6.6.3';
    end;
    if versaoDB = '1.6.6.3' then
    begin
      versaoDB:= '1.6.6.4';
    end;
    if versaoDB = '1.6.6.4' then
    begin
      //Tabela de Notas de Manutenção SAP
      AlterarNomeFieldDB('PrazoGM','PrazoGM',
      'tblNotaManutencao','VARCHAR(15)',FrmDataModule.ADOConnectionMemoria);
      AlterarNomeFieldDB('VencimentoAvaria','VencimentoAvaria',
      'tblNotaManutencao','VARCHAR(15)',FrmDataModule.ADOConnectionMemoria);
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
      //Tabela de Notas de Manutenção SAP
      AlterarNomeFieldDB('DenominacaoLocalInstalacao','DenominacaoLocal',
      'tblNotaManutencao','VARCHAR(255)',FrmDataModule.ADOConnectionMemoria);

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
      //LEC Local de Instalação
      CriarFieldDB('CPO','tblCarteiraTrabalho','VARCHAR(2)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      CriarFieldDB('Classe','tblLocalInstalacao','VARCHAR(25)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      CriarFieldDB('BCVT','tblLocalInstalacao','YESNO)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      //Criar Tabela tblRTI
      CriarTableDB('tblRTI','idRTI',
      '[NotaManutencao] INTEGER, [NumItem] VARCHAR(4), [DescricaoFalha] VARCHAR(100), [DataVencimento] DATETIME,'+
      '[Medida] VARCHAR(4),[DataPH] DATETIME,[DataRECB] DATETIME,[CodigoMedidaABCD] VARCHAR(2), '+
      '[DataPlanInicio] DATETIME,[DataPlanFim] DATETIME,[StatusUsuarioMedida] VARCHAR(100),[CentroTrabalhoResp] VARCHAR(25), '+
      '[StatusUsuarioNota] VARCHAR(25), [StatusSistemaMedida] VARCHAR(25), [CodigoMedidas] VARCHAR(4), '+
      '[GrupoCodeMedidas] VARCHAR(25), [LocalInstalacao] VARCHAR(50), [NEquipamento] VARCHAR(25), '+
      '[OrdemManutencao] VARCHAR(25), [StatusSistemaNota] VARCHAR(50), [TextoBreveMedida] VARCHAR(50), '+
      '[CodigoAcao] VARCHAR(25), [AreaOp] VARCHAR(4),[TextoLongoRTI] Memo,[Plataforma] VARCHAR(100), '+
      '[BooleanCritico] YESNO,[BCVT] YESNO,[Classe] VARCHAR(40), [DenominacaoLocalInstalacao] VARCHAR(40),'+
      '[ImportadoPor] VARCHAR(10), [DataImportacao] DATETIME,[ComputadorImportacao] VARCHAR(20),[txtTipoEtapaServico] VARCHAR(100)',
      FrmDataModule.ADOConnectionMemoria);
      //Criar Tabela tblClasseBCVT
      CriarTableDB('tblClasseBCVT','idClasseBCVT',
      '[Classe] VARCHAR(100) ',FrmDataModule.ADOConnectionConsulta);
      versaoDB:= '1.6.7.3';
    end;
    if versaoDB = '1.6.7.3' then
    begin
      //Memoria
      CriarFieldDB('Observacao','tblRTI','VARCHAR(255)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
      CriarFieldDB('StatusProgramacao','tblRTI','VARCHAR(15)',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
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
      //Memoria
      CriarFieldDB('BooleanItem','tblLocalInstalacao','YESNO',
      FrmDataModule.ADOQueryTemporarioDBMemoria);
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

      //dbMemoria
      ExcluirTableDB('tblCarteiraTrabalho',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblICPI',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblICPM',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblItemManutencao',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblLocalInstalacao',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblNotaManutencao',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblPrioridadeForcada',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblRTI',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblTextoLongoCarteira',FrmDataModule.ADOQueryTemporarioDBMemoria);
      ExcluirTableDB('tblTextoLongoOperacao',FrmDataModule.ADOQueryTemporarioDBMemoria);

      versaoDB:= '1.6.9.0';
    end;
    if versaoDB = '1.6.9.0' then
    begin
      //Tabela de dados básicos Colibri
      CriarFieldDB('RT_Tipo','tblColibri','VARCHAR(3)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Requisitante','tblColibri','VARCHAR(8)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_PessoaContato','tblColibri','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_TelContato','tblColibri','VARCHAR(16)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_CentroPlan','tblColibri','VARCHAR(4)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_GrpPlan','tblColibri','VARCHAR(3)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Modal','tblColibri','VARCHAR(3)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_HoraInicio','tblColibri','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_HoraVolta','tblColibri','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      //Excluir
      ExcluirFieldDB('tblColibri','RT_HoraFim',
          FrmDataModule.ADOConnectionColibri);

      //ExcluirFieldDB('tblColibri','RT_HoraCobertura',FrmDataModule.ADOConnectionColibri);
      CriarFieldDB('RT_HoraCobertura','tblColibri','DATETIME',
          FrmDataModule.ADOQueryTemporarioDBColibri);

     //tabela de Programação de Executantes
      CriarFieldDB('RT_Descricao','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Modal','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Classe','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_HoraInicio','tblProgramacaoExecutante','VARCHAR(5)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_CentroCusto','tblProgramacaoExecutante','VARCHAR(20)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Cobertura','tblProgramacaoExecutante','YESNO',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Mensagem','tblProgramacaoExecutante','VARCHAR(100)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Status','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarFieldDB('RT_Erro','tblProgramacaoExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBColibri);

      CriarFieldDB('booleanBateVolta','tblProgramacaoExecutante','YESNO',
          FrmDataModule.ADOQueryTemporarioDBColibri);

      //Tabela Cadastro de Executantes
      CriarFieldDB('CentroCusto','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('DiagramaRede','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('OperRede','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);
      CriarFieldDB('ElementoPEP','tblExecutante','VARCHAR(40)',
          FrmDataModule.ADOQueryTemporarioDBConsulta1);

      //Tabela Programacao de RT
      ExcluirTableDB('tblProgramacaoRT',FrmDataModule.ADOQueryTemporarioDBColibri);
      CriarTableDB(
        'tblProgramacaoRT',
        'idProgramacaoRT',
        '[DataInicio] DATETIME, '+
        '[CodigoSAP] VARCHAR(10), '+
        '[Origem] VARCHAR(20), '+
        '[txtDestino] VARCHAR(20), '+
        '[RT_Tipo] VARCHAR(3), '+
        '[RT_Descricao] VARCHAR(80), '+
        '[RT_Requisitante] VARCHAR(10), '+
        '[RT_TelContato] VARCHAR(16), '+
        '[RT_PessoaContato] VARCHAR(40), '+
        '[RT_CentroPlan] VARCHAR(4), '+
        '[RT_GrpPlan] VARCHAR(3), '+
        '[RT_Modal] VARCHAR(5), '+
        '[RT_Classe] VARCHAR(5), '+
        '[HoraInicio] VARCHAR(5), '+
        '[HoraVolta] VARCHAR(5), '+
        '[RT_CentroCusto] VARCHAR(20), '+
        '[RT_DiagramaRede] VARCHAR(20), '+
        '[RT_OperRede] VARCHAR(4), '+
        '[RT_ElementoPEP] VARCHAR(20), '+
        '[RT_Cobertura] YESNO, '+
        '[RT_Cancelada] YESNO, '+
        '[booleanBateVolta] YESNO, '+
        '[RT] VARCHAR(15), '+
        '[RT_Mensagem] VARCHAR(100), '+
        '[RT_Status] VARCHAR(40), '+
        '[RT_Erro] VARCHAR(100)',
        FrmDataModule.ADOConnectionColibri
      );
    end;
    FrmDataModule.ADOQueryColibri.Edit;
    FrmDataModule.DataSourceColibri.DataSet.FieldByName('versao').AsString:= versaoPRG;
    FrmDataModule.ADOQueryColibri.Post;
    MessageBox(0,'Conversão realizada com sucesso!','Colibri',MB_ICONEXCLAMATION);
  except
    MessageBox(0,'Ocorreu algum erro e a operação foi cancelada!','Colibri',MB_ICONERROR);
  end;
end;

procedure TFrmPrincipal.actDownlodaDBMemoriaExecute(Sender: TObject);
  var
    Caminho_Copia,enderecoMemoria: String;
begin
  ProgressBarIncializa(4,'Copiando dbMemoria.mdb...');
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  enderecoMemoria := FrmPrincipal.registroEndereco('Banco de dados');
  enderecoMemoria:= ExtractFilePath(enderecoMemoria)+'\dbMemoria.mdb';
  Caminho_Copia := ExtractFilePath(Application.ExeName) + 'dbMemoria.mdb';
  DeleteFile(Caminho_Copia);
  ProgressBarIncremento(1);
  CopyFile(PChar(enderecoMemoria), PChar(Caminho_Copia), false);
  ProgressBarIncremento(2);
  FrmPrincipal.conectarBDDireto(Caminho_Copia,FrmDataModule.ADOConnectionMemoria);
  ProgressBarAtualizar;
end;

procedure TFrmPrincipal.actDownloadExecute(Sender: TObject);
  var
    FileName,FilePath,enderecoREDE,
    enderecoREDEdbConsulta,enderecoREDEdbMemoria,
    enderecoLOCALdbColibri,enderecoLOCALdbConsulta,enderecoLOCALdbMemoria: String;
begin
  enderecoREDE:= registroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  //LOCAL
  enderecoLOCALdbColibri:= ExtractFilePath(Application.ExeName)+'LOCAL\'+FileName;
  enderecoLOCALdbConsulta:= ExtractFilePath(Application.ExeName)+'LOCAL\dbConsulta.mdb';
  enderecoLOCALdbMemoria:= ExtractFilePath(Application.ExeName)+'LOCAL\dbMemoria.mdb';
  //REDE
  enderecoREDEdbConsulta:= FilePath+'dbConsulta.mdb';
  enderecoREDEdbMemoria:= FilePath+'dbMemoria.mdb';
  //DOWNPLOAD
  FrmPrincipal.ProgressBarIncializa(4,'Download banco de dados...');
  //=============================================================================
  if CheckBoxColibri.Checked then
  begin
    CopyFile(PChar(enderecoREDE), PChar(enderecoLOCALdbColibri), false);
    conectarBDDireto(enderecoLOCALdbColibri,FrmDataModule.ADOConnectionColibri);
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  if CheckBoxConsulta.Checked then
  begin
    CopyFile(PChar(enderecoREDEdbConsulta), PChar(enderecoLOCALdbConsulta), false);
    conectarBDDireto(enderecoLOCALdbConsulta,FrmDataModule.ADOConnectionConsulta);
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  if CheckBoxMemoria.Checked then
  begin
    CopyFile(PChar(enderecoREDEdbMemoria), PChar(enderecoLOCALdbMemoria), false);
    conectarBDDireto(enderecoLOCALdbMemoria,FrmDataModule.ADOConnectionMemoria);
  end;
  //=============================================================================
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmPrincipal.actExcelDuplicadosExecute(Sender: TObject);
begin
  FrmPrincipal.ExcelStringGrid(RLDuplicados,'Duplicados','','',1);
end;

procedure TFrmPrincipal.actFiltroCancelarExecute(Sender: TObject);
begin
  PanelMagicFiltro1.Visible:= false;
end;

procedure TFrmPrincipal.actFiltroInserirExecute(Sender: TObject);
begin
  FrmPrincipal.inserirProcura(DBGridPalavraChave,ColunasLayoutPalavraChave);
  actProcurarPalavraChave.Execute;
  if (FrmPrincipal.PanelFiltrosTabela.Visible)AND(FrmPrincipal.PanelAjuda1.Visible) then
    actFiltrosTabela.Execute;
end;

procedure TFrmPrincipal.actFiltrosTabelaExecute(Sender: TObject);
begin
  FrmPrincipal.btnProcurarFiltrosTabela.Action:= actProcuraFiltrosTabela;
  FrmPrincipal.FiltrosTabela(DBGridPalavraChave,ColunasLayoutPalavraChave);
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
  SetLength(MatrizForaOperacao, 11);
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
    MatrizForaOperacao[8,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoAcesso').AsString;
    MatrizForaOperacao[9,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('SituacaoDegraus').AsString;
    MatrizForaOperacao[10,i]:= FrmDataModule.DataSourceTemporarioDBConsulta1.
    DataSet.FieldByName('DataRealizacaoDegraus').AsString;

    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
end;

procedure TFrmPrincipal.actInserirRegistroExecute(Sender: TObject);
  var
    PalavraChave: String;
    idProgramacaoExecutante,idProgramacaoDiaria,NumRegistros: Integer;
    passou: Boolean;
begin
  passou:= true;
  PalavraChave:= '';
  while not FrmDataModule.ADOQueryPalavraChave.Eof do
  begin
    if FrmDataModule.ADOQueryPalavraChave.FieldByName('booleanSelecao').AsBoolean then
    begin
      if PalavraChave <> '' then
        PalavraChave:= PalavraChave+'; '+FrmDataModule.ADOQueryPalavraChave.FieldByName('PalavraChave').AsString
      else
        PalavraChave:= FrmDataModule.ADOQueryPalavraChave.FieldByName('PalavraChave').AsString;
    end;
    FrmDataModule.ADOQueryPalavraChave.Next;
  end;
  //================================================================
  if PalavraChave = '' then
  begin
    passou:= false;
    MessageBox(0,'O campo "Palavra Chave" não pode estar em branco!','Avaliação de Programação',MB_ICONERROR);
  end;
  if passou then
  begin
    case RadioGroupFonteCancelamento.ItemIndex of
      0://Gerenciar Solicitações (Cancelamento)
      begin
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
        NumRegistros:= FrmDataModule.ADOQueryGerenciarExecutante.RecordCount;
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Atribuindo status para programação...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('idProgramacaoExecutante').AsInteger;

            FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,0,
            'Cancelado','Não Executada',PalavraChave);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        //Calcular Cancelados
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Calculando: Aprovados, Cancelados e N° de Executantes...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger;
            FrmPrincipal.GravarCanceladoAprovado(idProgramacaoDiaria);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
        FrmGerenciarSolicitacoes.actMatrizOrigemDestino.Execute;
        FrmGerenciarSolicitacoes.StatusLinhaSelecionada;
        FrmGerenciarSolicitacoes.actContadorSolicitacao.Execute;
      end;
      1: //Marcar os não executados
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
            '','Não Executada',PalavraChave);
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
      2: //Gerenciamento de Solicitações (Mudança)
      begin
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
        NumRegistros:= FrmDataModule.ADOQueryGerenciarExecutante.RecordCount;
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Atribuindo status para programação...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('idProgramacaoExecutante').AsInteger;

            FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,0,
            'Mudança','Não Executada',PalavraChave);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        //Calcular Cancelados
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Calculando: Aprovados, Cancelados e N° de Executantes...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger;
            FrmPrincipal.GravarCanceladoAprovado(idProgramacaoDiaria);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
        FrmGerenciarSolicitacoes.actMatrizOrigemDestino.Execute;
        FrmGerenciarSolicitacoes.StatusLinhaSelecionada;
        FrmGerenciarSolicitacoes.actContadorSolicitacao.Execute;
      end;
    end;
  end;
  PanelAjuda1.Visible:= false;
end;

procedure TFrmPrincipal.actInterromperExecute(Sender: TObject);
begin
  Interromper:= true;
end;

procedure TFrmPrincipal.actLayoutSORTExecute(Sender: TObject);
  var
    Ordem: String;
begin
  Ordem:= FrmPrincipal.Clasifica(RLColunasOpcoes,1,true);
  FrmPrincipal.AutoFitGrade(RLColunasOpcoes);
  RLColunasOpcoes.ColWidths[0]:= 18;
  if Ordem = 'DESC' then
  begin
    actLayoutSORT.ImageIndex:= 86;
  end
  else if Ordem = 'ASC' then
  begin
    actLayoutSORT.ImageIndex:= 85;
  end;
end;

procedure TFrmPrincipal.actLimparFiltrosExecute(Sender: TObject);
begin
  FrmPrincipal.LimparColunasFiltro(DBGridPalavraChave,ColunasLayoutPalavraChave);
  actProcurarPalavraChave.Execute;
  if (FrmPrincipal.PanelFiltrosTabela.Visible)AND(FrmPrincipal.PanelAjuda1.Visible) then
    actFiltrosTabela.Execute;
end;

procedure TFrmPrincipal.actListarExecute(Sender: TObject);
  var
    ListaGroup: TStringList;
begin
  ListaGroup:= GroupFieldDBGrid(MagicSQLQuery,MagicFieldName,MagicFonteDB,StatusBarMagicFiltro);
  FrmPrincipal.SetupCheckListSQL1(MagicFieldName,MagicTextoProcura,CheckListBoxFiltroGRID,
  ListaGroup);
end;

procedure TFrmPrincipal.actLOCALExecute(Sender: TObject);
  var
    enderecoLOCAL,
    enderecoREDE,FileName,FilePath,dbConsulta,dbMemoria: String;
begin
  enderecoREDE:= registroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  enderecoLOCAL:= ExtractFilePath(Application.ExeName);
  if Application.MessageBox(PChar(
  'Deseja substituir os bancos de dados LOCAL atual pelos bancos dados da REDE selecionados abaixo?'),
  '.::ATENÇÃO::.',36) = 6 then
  begin
    //Copiar e Substituir Arquivos
    dbConsulta:= FilePath+'\dbConsulta.mdb';
    dbMemoria:= FilePath+'\dbMemoria.mdb';
    FrmPrincipal.ProgressBarIncializa(7,'Conexão LOCAL...');
    CopyFile(PChar(enderecoREDE), PChar(enderecoLOCAL+'\LOCAL\'+FileName), false);
    FrmPrincipal.ProgressBarIncremento(1);
    CopyFile(PChar(dbConsulta), PChar(enderecoLOCAL+'\LOCAL\dbConsulta.mdb'), false);
    FrmPrincipal.ProgressBarIncremento(1);
    CopyFile(PChar(dbMemoria), PChar(enderecoLOCAL+'\LOCAL\dbMemoria.mdb'), false);
    FrmPrincipal.ProgressBarIncremento(1);
    //Conexão
    conectarBDDireto(enderecoLOCAL+'\LOCAL\'+FileName,FrmDataModule.ADOConnectionColibri);
    FrmPrincipal.ProgressBarIncremento(1);
    conectarBDDireto(enderecoLOCAL+'\LOCAL\dbConsulta.mdb',FrmDataModule.ADOConnectionConsulta);
    FrmPrincipal.ProgressBarIncremento(1);
    conectarBDDireto(enderecoLOCAL+'\LOCAL\dbMemoria.mdb',FrmDataModule.ADOConnectionMemoria);
  end
  else
  begin
    conectarBDDireto(enderecoLOCAL+'\LOCAL\'+FileName,FrmDataModule.ADOConnectionColibri);
    FrmPrincipal.ProgressBarIncremento(1);
    conectarBDDireto(enderecoLOCAL+'\LOCAL\dbConsulta.mdb',FrmDataModule.ADOConnectionConsulta);
    FrmPrincipal.ProgressBarIncremento(1);
    conectarBDDireto(enderecoLOCAL+'\LOCAL\dbMemoria.mdb',FrmDataModule.ADOConnectionMemoria);
  end;
  FrmPrincipal.ProgressBarAtualizar;
  FrmPrincipal.Caption:= 'Colibri: LOCAL->REDE ';
  booLOCAL:= true;
  actLOCAL.Enabled:= false;
  actREDE.Enabled:= true;
  PanelTituloAjuda1.Caption:= 'Conectado LOCAL';
  PanelTituloAjuda1.Color:= clRed;
  if  (logPerfil = 'Administrador')OR
      (logPerfil = 'Supervisor')OR
      (logPerfil = 'Programação') then
    actUpload.Enabled:= true
  else
    actUpload.Enabled:= false;

  actDownload.Enabled:= true;
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
end;

procedure TFrmPrincipal.actMatrizExecutanteAPLATExecute(Sender: TObject);
  var
    NumRegistros,i: Integer;
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblExecutanteAPLAT.* FROM tblExecutanteAPLAT;';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBMemoria.Close;
  FrmDataModule.ADOQueryTemporarioDBMemoria.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBMemoria.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBMemoria.Open;
  NumRegistros:= FrmDataModule.ADOQueryTemporarioDBMemoria.RecordCount;
  //=========================================================================
  MatrizExecutanteAPLAT:= nil;
  SetLength(MatrizExecutanteAPLAT, 7);
  for i := 0 to High(MatrizExecutanteAPLAT) do
    SetLength(MatrizExecutanteAPLAT[i], NumRegistros);

  for i := 0 to High(MatrizExecutanteAPLAT[0]) do
  begin

    MatrizExecutanteAPLAT[0,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtTipoEtapaServico').AsString;
    MatrizExecutanteAPLAT[1,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtFuncao').AsString;
    MatrizExecutanteAPLAT[2,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtEmpresa').AsString;
    MatrizExecutanteAPLAT[3,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtNomeExecutante').AsString;
    MatrizExecutanteAPLAT[4,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('codigoSAP').AsString;
    MatrizExecutanteAPLAT[5,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('DataEmbarque').AsString;
    MatrizExecutanteAPLAT[6,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('DataDesembarque').AsString;

    FrmDataModule.ADOQueryTemporarioDBMemoria.Next;
  end;
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
  SetLength(MatrizExecutanteCadastro, 11);
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


    FrmDataModule.ADOQueryTemporarioDBConsulta1.Next;
  end;
end;

procedure TFrmPrincipal.actMatrizExecutanteT31Execute(Sender: TObject);
  var
    NumRegistros,i: Integer;
    SQLBase: String;
begin
  SQLBase:= 'SELECT tblExecutanteSAP.* FROM tblExecutanteSAP;';
  //=========================================================================
  FrmDataModule.ADOQueryTemporarioDBMemoria.Close;
  FrmDataModule.ADOQueryTemporarioDBMemoria.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBMemoria.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBMemoria.Open;
  NumRegistros:= FrmDataModule.ADOQueryTemporarioDBMemoria.RecordCount;
  //=========================================================================
  MatrizExecutanteT31:= nil;
  SetLength(MatrizExecutanteT31, 11);
  for i := 0 to High(MatrizExecutanteT31) do
    SetLength(MatrizExecutanteT31[i], NumRegistros);

  for i := 0 to High(MatrizExecutanteT31[0]) do
  begin

    MatrizExecutanteT31[0,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtTipoEtapaServico').AsString;
    MatrizExecutanteT31[1,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtFuncao').AsString;
    MatrizExecutanteT31[2,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtEmpresa').AsString;
    MatrizExecutanteT31[3,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('txtNomeExecutante').AsString;
    MatrizExecutanteT31[4,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('codigoSAP').AsString;
    MatrizExecutanteT31[5,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('Documento').AsString;
    MatrizExecutanteT31[6,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('TipoNota').AsString;
    MatrizExecutanteT31[7,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('Nota').AsString;
    MatrizExecutanteT31[8,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('DataEmbarque').AsString;
    MatrizExecutanteT31[9,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('Origem').AsString;
    MatrizExecutanteT31[10,i]:= FrmDataModule.DataSourceTemporarioDBMemoria.
    DataSet.FieldByName('Destino').AsString;

    FrmDataModule.ADOQueryTemporarioDBMemoria.Next;
  end;
end;

procedure TFrmPrincipal.actProcuraFiltrosTabelaExecute(Sender: TObject);
begin
  frmPrincipal.CarregaFiltrosProcura(ColunasLayoutPalavraChave);
  actProcurarPalavraChave.Execute;
end;

procedure TFrmPrincipal.actProcurarPalavraChaveExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutPalavraChave,true);
  SQLBase:= 'SELECT tblPalavraChave.* FROM tblPalavraChave'+
  SQLString+' ORDER BY PalavraChave;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryPalavraChave,StatusBarPalavraChave);
end;

procedure TFrmPrincipal.actGridASCExecute(Sender: TObject);
begin
  FrmPrincipal.ClassificaDBGrid(DBGridPalavraChave,
  FrmDataModule.ADOQueryPalavraChave,0);
end;

procedure TFrmPrincipal.actGridDESCExecute(Sender: TObject);
begin
  FrmPrincipal.ClassificaDBGrid(DBGridPalavraChave,
  FrmDataModule.ADOQueryPalavraChave,1);
end;

procedure TFrmPrincipal.actGridSelLimpaExecute(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to FrmPrincipal.CheckListBoxFiltroGRID.Count-1 do
    FrmPrincipal.CheckListBoxFiltroGRID.Checked[i]:= false;
end;

procedure TFrmPrincipal.actGridSelTudoExecute(Sender: TObject);
var
  I: Integer;
begin
  for I := 0 to FrmPrincipal.CheckListBoxFiltroGRID.Count-1 do
    FrmPrincipal.CheckListBoxFiltroGRID.Checked[i]:= true;
end;

procedure TFrmPrincipal.actREDEExecute(Sender: TObject);
  var
    enderecoREDE,FilePath,dbConsulta,dbMemoria: String;
begin
  enderecoREDE:= registroEndereco('Banco de dados');
  FilePath:= ExtractFilePath(enderecoREDE);
  dbConsulta:= FilePath+'\dbConsulta.mdb';
  dbMemoria:= FilePath+'\dbMemoria.mdb';
  FrmPrincipal.ProgressBarIncializa(2,'Conexão REDE...');
  FrmPrincipal.ProgressBarIncremento(1);
  conectarBD(enderecoREDE,FrmDataModule.ADOConnectionColibri);
  conectarBDDireto(dbConsulta,FrmDataModule.ADOConnectionConsulta);
  conectarBDDireto(dbMemoria,FrmDataModule.ADOConnectionMemoria);
  FrmPrincipal.ProgressBarAtualizar;
  booLOCAL:= false;
  actLOCAL.Enabled:= true;
  actREDE.Enabled:= false;
  PanelTituloAjuda1.Caption:= 'Conectado REDE';
  PanelTituloAjuda1.Color:= clGreen;
  actUpload.Enabled:= false;
  actDownload.Enabled:= false;
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
end;

procedure TFrmPrincipal.actSalvarListaExecute(Sender: TObject);
  var
    txtPontoVirgula: String;
    i: Integer;
begin
  SaveDialog1.Filter := 'Texto|*.txt';
  SaveDialog1.DefaultExt := '*.txt';
  SaveDialog1.FileName:= PanelTituloMagic.Caption;
  if SaveDialog1.Execute then
  begin
    CheckListBoxFiltroGRID.Items.SaveToFile(SaveDialog1.FileName);
    ShellExecute(0,'open',PChar(SaveDialog1.FileName),'','',SW_SHOWNORMAL);
    txtPontoVirgula:= '';
    for I := 0 to CheckListBoxFiltroGRID.Items.Count-1 do
    begin
      if txtPontoVirgula = '' then
        txtPontoVirgula:= CheckListBoxFiltroGRID.Items[i]
      else
        txtPontoVirgula:= txtPontoVirgula+';'+CheckListBoxFiltroGRID.Items[i];
    end;
    ClipBoard.AsText := txtPontoVirgula;
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

procedure TFrmPrincipal.actSubstituirPorExecute(Sender: TObject);
begin
  FrmPrincipal.SubstituirPor(DBGridPalavraChave,FrmDataModule.ADOQueryPalavraChave,
  FrmDataModule.DataSourcePalavraChave);
end;

procedure TFrmPrincipal.actUploadDBMemoriaExecute(Sender: TObject);
  var
    Caminho_Copia,enderecoMemoria: String;
begin
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  enderecoMemoria := FrmPrincipal.registroEndereco('Banco de dados');
  enderecoMemoria:= ExtractFilePath(enderecoMemoria)+'\dbMemoria.mdb';
  Caminho_Copia := ExtractFilePath(Application.ExeName) + 'dbMemoria.mdb';
  FrmPrincipal.compactarDB(Caminho_Copia,false,false,FrmDataModule.ADOConnectionMemoria);
  FrmDataModule.ADOConnectionMemoria.Connected:= false;
  CopyFile(PChar(Caminho_Copia), PChar(enderecoMemoria), false);
  FrmDataModule.ADOConnectionMemoria.Connected:= true;
  //FrmPrincipal.conectarBDDireto(enderecoMemoria,FrmDataModule.ADOConnectionMemoria);
end;

procedure TFrmPrincipal.actUploadExecute(Sender: TObject);
  var
    FileName,FilePath,enderecoREDE,
    enderecoREDEdbConsulta,enderecoREDEdbMemoria,
    enderecoLOCALdbColibri,enderecoLOCALdbConsulta,enderecoLOCALdbMemoria: String;
begin
  enderecoREDE:= registroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  //LOCAL
  enderecoLOCALdbColibri:= ExtractFilePath(Application.ExeName)+'LOCAL\'+FileName;
  enderecoLOCALdbConsulta:= ExtractFilePath(Application.ExeName)+'LOCAL\dbConsulta.mdb';
  enderecoLOCALdbMemoria:= ExtractFilePath(Application.ExeName)+'LOCAL\dbMemoria.mdb';
  //REDE
  enderecoREDEdbConsulta:= FilePath+'dbConsulta.mdb';
  enderecoREDEdbMemoria:= FilePath+'dbMemoria.mdb';
  //DOWNPLOAD
  FrmPrincipal.ProgressBarIncializa(4,'Upload banco de dados...');
  //============================================================================
  if CheckBoxColibri.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbColibri da REDE?'),'.::ATENÇÃO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbColibri), PChar(enderecoREDE), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  if CheckBoxConsulta.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbConsulta da REDE?'),'.::ATENÇÃO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbConsulta), PChar(enderecoREDEdbConsulta), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  if CheckBoxMemoria.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbMemoria da REDE?'),'.::ATENÇÃO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbMemoria), PChar(enderecoREDEdbMemoria), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  FrmPrincipal.ProgressBarAtualizar;
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

procedure TFrmPrincipal.AutoFitGrade(aGrade: TStringGrid);
var
  i, j: Integer;
  x: array of Integer;
begin
  // ==============STRINGGRIDCALCULO====================
  SetLength(x, aGrade.ColCount);
  for i := 0 to aGrade.ColCount - 1 do
    x[i] := 0;
  for i := 0 to aGrade.RowCount - 1 do
  begin
    for j := 0 to aGrade.ColCount - 1 do
    begin
      if aGrade.Canvas.TextWidth(aGrade.Cells[j, i]) > x[j] then
        x[j] := aGrade.Canvas.TextWidth(aGrade.Cells[j, i]);
    end;
  end;
  for i := 0 to aGrade.ColCount - 1 do
  begin
    if x[i] = 0 then
      x[i] := 89;
  end;
  for i := 0 to aGrade.RowCount - 1 do
  begin
    for j := 0 to aGrade.ColCount - 1 do
      aGrade.ColWidths[j] := x[j] + 14;
  end;
end;

procedure TFrmPrincipal.AutoFitStatusBar(aStatus: TStatusBar);
var
  i: Integer;
  str: String;
begin
  aStatus.Canvas.Font := aStatus.Font;
  for i := 0 to aStatus.Panels.Count-1 do
  begin
    //Escrever novamente
    str:= aStatus.Panels[i].Text;
    aStatus.Panels[i].Text:= '';
    aStatus.Panels[i].Text:= str;
    //=======================================================
    aStatus.Panels[i].Width:= aStatus.Canvas.TextWidth(aStatus.Panels[i].Text)+15;
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
  idProgramacaoExecutante,Fonte: Integer; StatusProgramacao,StatusExecucao,
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
        //Marcar como Não Executada
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('StatusExecucao').AsString:= StatusExecucao;
        //=====================================================
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('AvaliadoPorExecucao').AsString:= logChave;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('ComputadorExecucao').AsString:= logMaquina;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('DataAvaliacaoExecucao').AsDateTime:= now;
        //=====================================================
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
      end;
      1://Não Executa (Consulta Executantes)
      begin
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('StatusExecucao').AsString:= StatusExecucao;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('MotivoNaoExecucao').AsString:= Motivo;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('AvaliadoPorExecucao').AsString:= logChave;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('ComputadorExecucao').AsString:= logMaquina;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
        FieldByName('DataAvaliacaoExecucao').AsDateTime:= now;
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
      end;
    end;
  except
    ShowMessage('Erro');
  end;
end;

procedure TFrmPrincipal.BitBtn4Click(Sender: TObject);
begin
  PanelAjuda1.Visible:= false;
end;

procedure TFrmPrincipal.botaoMoveTudo(Conection: TADOConnection; NomeArquivoMemo,NomeTabela,NumRLTAB: String;Grid: TDBGrid; ColunasLayout: TStringGrid);
begin
  FrmPrincipal.LayoutPadrao(NomeArquivoMemo,ColunasLayout,NumRLTAB);
  FrmPrincipal.CarregarRLColunas(NomeArquivoMemo, ColunasLayout);
  FrmPrincipal.CarregarColunasGRID(Conection,NomeTabela,Grid,ColunasLayout);
end;

procedure TFrmPrincipal.buscaFiledGrid1(FieldName, palavraBusca,Operador: String;
  ColunasLayout: TStringGrid; ACol: Integer;BooleanConcatenar: Boolean);
var
  I: Integer;
begin
  for I := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[0,i] = FieldName then
    begin
      if (BooleanConcatenar)AND(ColunasLayout.Cells[ACol,i]<>'') then
      begin
        ColunasLayout.Cells[ACol,i]:= ColunasLayout.Cells[ACol,i]+';'+palavraBusca;
        if ACol = 4 then
          ColunasLayout.Cells[5,i]:= Operador;
      end
      else
      begin
        ColunasLayout.Cells[ACol,i]:= palavraBusca;
        if ACol = 4 then
          ColunasLayout.Cells[5,i]:= Operador;
      end;
      break;
    end;
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
    if ((logPerfil = 'Administrador')OR
    (logPerfil = 'Programacao')OR
    (logPerfil = 'Supervisão')) then
      DateTimePickerMinima.DateTime:= EncodeDate(2012, 01, 01)
    else
      DateTimePickerMinima.DateTime:= now;//IncDay(now,1);
  end
  else
  begin
    if (logPerfil = 'Administrador') then
      DateTimePickerMinima.DateTime:= EncodeDate(2012, 01, 01)
    else
      DateTimePickerMinima.DateTime:= now;//IncDay(now,1);
  end;

  DateTimePickerMinima.Time:= 0;
  Result:= DateTimePickerMinima.DateTime;
end;

procedure TFrmPrincipal.CarregaFiltrosProcura(ColunasLayout: TStringGrid);
    var
      i,index: Integer;
begin
  for I := 1 to RLFiltrosTabela.RowCount-1 do
  begin
    index:= indexLayoutCaption(RLFiltrosTabela.Cells[0,i],ColunasLayout);
    RLFiltrosTabela.Cells[4,index]:= '';
    RLFiltrosTabela.Cells[5,index]:= '';
    ColunasLayout.Cells[4,index]:= RLFiltrosTabela.Cells[1,i];
    ColunasLayout.Cells[5,index]:= RLFiltrosTabela.Cells[2,i];
  end;
end;

procedure TFrmPrincipal.CarregarColunasGRID(Conection: TADOConnection;
tblTabela: String; Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    i,coluna: Integer;
    Titulo: String;
    ListaColuna: TStringList;
    booleanReadyOnly: Boolean;
begin
  Grid.DataSource.Enabled := false;
  Grid.columns.Clear;
  if FrmPrincipal.RLColunasAtivas.RowCount>1 then
  begin
    coluna:= 0;
    FrmPrincipal.ProgressBarIncializa(FrmPrincipal.
    RLColunasAtivas.RowCount,'Carregando colunas...');
    for I := 0 to FrmPrincipal.RLColunasAtivas.RowCount-1 do
    begin
      Titulo:= FrmPrincipal.RLColunasAtivas.Cells[1,i];
      ListaColuna:= dadosColuna(Titulo,ColunasLayout);
      if ListaColuna.Count > 1 then
      begin
        if ListaColuna[3] = 'TRUE' then
          booleanReadyOnly:= TRUE
        else
        begin
          booleanReadyOnly:= false;
        end;
        FrmPrincipal.addColuna(Grid,
        ListaColuna[0],ListaColuna[1],ListaColuna[2],Coluna,
        StrToInt(ListaColuna[4]),booleanReadyOnly);
        Coluna:= Coluna+1;
        {if ((booleanReadyOnly = false)AND(tblTabela<>'')) then
        begin
          try
            FrmPrincipal.SetupGridPickListSQL(Conection, ListaColuna[0],
            'SELECT '+tblTabela+'.'+ListaColuna[0]+' FROM '
            +tblTabela+' GROUP BY '+tblTabela+'.'+ListaColuna[0]+';',
            Grid,0);
          except
          end;
        end;}
      end;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
    Grid.DataSource.Enabled := true;
  end;
  FrmPrincipal.LimparColunasFiltro(Grid,ColunasLayout);
  FrmPrincipal.ProgressBarAtualizar;
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

procedure TFrmPrincipal.CarregarRLColunas(NomeArquivoMemo: String; ColunasLayout: TStringGrid);
  var
    i: integer;
    ListaColunas: TStringList;
    Endereco,strTitulo: String;
begin
  ListaColunas:=TStringList.Create;
  ListaColunas.delimiter:= ';';
  ListaColunas.StrictDelimiter:=true;
  Endereco:= ExtractFilePath(Application.ExeName)+'Layout\'+NomeArquivoMemo;
  if FileExists(Endereco) then
    MemoPrincipal.Lines.LoadFromFile(Endereco)
  else
    FrmPrincipal.LayoutPadrao(ExtractFilePath(Application.ExeName)+
    'Layout\'+NomeArquivoMemo,ColunasLayout,'TAB1;');

  FrmPrincipal.RLColunasAtivas.RowCount:=1;
  FrmPrincipal.RLColunasOpcoes.RowCount:=1;
  FrmPrincipal.RLColunasAtivas.Cells[1,0]:= 'COLUNAS ATIVAS';
  FrmPrincipal.RLColunasOpcoes.Cells[1,0]:= 'COLUNAS OPTATIVAS';
  for i := 0 to MemoPrincipal.Lines.Count-1 do
  begin
    ListaColunas.Clear;
    ListaColunas.DelimitedText:= MemoPrincipal.Lines[i];
    if ListaColunas[0] = 'TAB1' then
    begin
      strTitulo:= FrmPrincipal.DeleteChar(char(9650),ListaColunas[1]);
      strTitulo:= FrmPrincipal.DeleteChar(char(9660),strTitulo);
      strTitulo:= FrmPrincipal.DeleteChar(char(8722),strTitulo);
      FrmPrincipal.RLColunasAtivas.RowCount:= FrmPrincipal.RLColunasAtivas.RowCount+1;
      FrmPrincipal.RLColunasAtivas.Cells[1,FrmPrincipal.RLColunasAtivas.RowCount-1]:=
      strTitulo;
    end
    else if ListaColunas[0] = 'TAB2' then
    begin
      strTitulo:= FrmPrincipal.DeleteChar(char(9650),ListaColunas[1]);
      strTitulo:= FrmPrincipal.DeleteChar(char(9660),strTitulo);
      strTitulo:= FrmPrincipal.DeleteChar(char(8722),strTitulo);
      FrmPrincipal.RLColunasOpcoes.RowCount:= FrmPrincipal.RLColunasOpcoes.RowCount+1;
      FrmPrincipal.RLColunasOpcoes.Cells[1,FrmPrincipal.RLColunasOpcoes.RowCount-1]:=
      strTitulo;
    end;
  end;
  try
    FrmPrincipal.RLColunasAtivas.FixedRows:= 1;
  except
  end;
  try
    FrmPrincipal.RLColunasOpcoes.FixedRows:= 1;
  except
  end;
  FrmPrincipal.AutoFitGrade(FrmPrincipal.RLColunasAtivas);
  FrmPrincipal.AutoFitGrade(FrmPrincipal.RLColunasOpcoes);
  FrmPrincipal.RLColunasAtivas.ColWidths[0]:= 18;
  FrmPrincipal.RLColunasOpcoes.ColWidths[0]:= 18;
end;

procedure TFrmPrincipal.CarregarRLColunasAtivasGRID(Grid: TDBGrid);
  var
    i: integer;
    strTitulo: String;
begin
  FrmPrincipal.RLColunasAtivas.RowCount:=1;
  FrmPrincipal.RLColunasAtivas.Cells[1,0]:= 'COLUNAS ATIVAS';
  for i := 0 to Grid.Columns.Count-1 do
  begin
    strTitulo:= FrmPrincipal.DeleteChar(char(9650),Grid.Columns.Items[i].Title.Caption);
    strTitulo:= FrmPrincipal.DeleteChar(char(9660),strTitulo);
    strTitulo:= FrmPrincipal.DeleteChar(char(8722),strTitulo);
    FrmPrincipal.RLColunasAtivas.RowCount:= FrmPrincipal.RLColunasAtivas.RowCount+1;
    FrmPrincipal.RLColunasAtivas.Cells[1,FrmPrincipal.RLColunasAtivas.RowCount-1]:=strTitulo;
  end;
  try
    FrmPrincipal.RLColunasAtivas.FixedRows:= 1;
  except
  end;
  FrmPrincipal.AutoFitGrade(FrmPrincipal.RLColunasAtivas);
  FrmPrincipal.RLColunasAtivas.ColWidths[0]:= 18;
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
    if logPerfil = 'Administrador' then //Administrador
    begin
      CadastroUsuario1.Enabled:= true;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      ImportarPlanilhas1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      actUpload.Enabled:= true;
      actDownload.Enabled:= true;
      Converterverso1.Enabled:= true;
    end
    else if logPerfil = 'Supervisor' then
    begin
      CadastroUsuario1.Enabled:= true;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      ImportarPlanilhas1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      actUpload.Enabled:= true;
      actDownload.Enabled:= true;
      Converterverso1.Enabled:= false;
    end
    else if logPerfil = 'Programação' then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      ImportarPlanilhas1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= true;
      actUpload.Enabled:= true;
      actDownload.Enabled:= true;
      Converterverso1.Enabled:= false;
    end
    else if logPerfil = 'Hotelaria' then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= true;
      SalvarBancoDadosComo1.Enabled:= true;
      ImportarPlanilhas1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= true;
      CompactarBancoDados1.Enabled:= false;
      actUpload.Enabled:= true;
      actDownload.Enabled:= true;
      Converterverso1.Enabled:= false;
    end
    else //if logPerfil = 'Consulta' then
    begin
      CadastroUsuario1.Enabled:= false;
      ProgramacaoDiaria2.Enabled:= false;
      SalvarBancoDadosComo1.Enabled:= true;
      ImportarPlanilhas1.Enabled:= true;
      GerenciarTransportes1.Enabled:= true;
      Executantes1.Enabled:= false;
      CompactarBancoDados1.Enabled:= false;
      actUpload.Enabled:= false;
      actDownload.Enabled:= false;
      Converterverso1.Enabled:= false;
    end;
  end;
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
  FrmPrincipal.PanelMagicFiltro1.Visible:= false;
end;

procedure TFrmPrincipal.ComboBoxOperadorCloseUp(Sender: TObject);
begin
  RLFiltrosTabela.Cells[2,RLFiltrosTabela.Row]:= (ComboBoxOperador.Text);
  ComboBoxOperador.Visible := False;
end;

procedure TFrmPrincipal.ComboBoxOperadorKeyPress(Sender: TObject;
  var Key: Char);
begin
  Key:= #0;
end;

procedure TFrmPrincipal.ComboBoxOperadorMouseLeave(Sender: TObject);
begin
  ComboBoxOperador.Visible := False;
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
    //Memoria
    caminhoReg:= ExtractFilePath(caminhoReg)+'dbMemoria.mdb';
    compactarDB(caminhoReg,false,false,FrmDataModule.ADOConnectionMemoria);
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

procedure TFrmPrincipal.compactarDBMemoria(EnderecoArquivo: String;SourceADOConnection: TADOConnection);
var
  dao: OleVariant;
  Caminho_Temp: String;
begin
  // Compactar e Reparar Banco de dados
  try
    FrmPrincipal.ProgressBarIncializa(18,'Compactando dbMemoria.mdb');
    //Endereço do arquivo temporário
    Caminho_Temp := ExtractFilePath(Application.ExeName) + 'TEMP.mdb';
    //Copiar Arquivo para o Temp.
    CopyFile(PChar(EnderecoArquivo), PChar(Caminho_Temp), false);
    //Excluir arquivo temporário antigo caso exista
    FrmPrincipal.ProgressBarIncremento(2);
    DeleteFile(Caminho_Temp);
    //Fechar Conexão
    SourceADOConnection.Connected := false;
    FrmPrincipal.ProgressBarIncremento(2);
    SourceADOConnection.Close;
    FrmPrincipal.ProgressBarIncremento(3);
    //Criar banco
    dao := CreateOleObject('DAO.DBEngine.36');
    FrmPrincipal.ProgressBarIncremento(5);
    //Compactar banco
    dao.CompactDatabase(EnderecoArquivo, Caminho_Temp);
    FrmPrincipal.ProgressBarIncremento(5);
    CopyFile(PChar(Caminho_Temp), PChar(EnderecoArquivo), false);
    //Excluir arquivo temporário
    FrmPrincipal.ProgressBarIncremento(1);
    DeleteFile(Caminho_Temp);
    FrmPrincipal.ProgressBarAtualizar;
  except
    FrmPrincipal.ProgressBarAtualizar;
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
  actAjudaLimpar.Execute;
  PanelCONECTION.Visible:= true;
  PanelCONECTION.Align:= alClient;
  PanelAjuda1.Width:= 330;
  PanelAjuda1.Height:= 160;

  if booLOCAL then
  begin
    PanelTituloAjuda1.Caption:= 'Conectado LOCAL';
    PanelTituloAjuda1.Color:= clRed;
    actUpload.Enabled:= true;
    actDownload.Enabled:= true;
  end
  else
  begin
    PanelTituloAjuda1.Caption:= 'Conectado REDE';
    PanelTituloAjuda1.Color:= clGreen;
    actUpload.Enabled:= false;
    actDownload.Enabled:= false;
  end;
  PanelAjuda1.Visible:= true;
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

procedure TFrmPrincipal.configPiroridade(edtPonto, edtLimite,
  edtValorMax: TDBEdit; relacaoData: TDBComboBox; check: TDBCheckBox);
begin
  if check.Checked then
  begin
    edtPonto.Enabled:= true;
    edtLimite.Enabled:= true;
    edtValorMax.Enabled:= true;
    relacaoData.Enabled:= true;
    relacaoData.Hint:= HintRelacaoTempo(relacaoData);
    edtPonto.Color:= clWhite;
    edtLimite.Color:= clWhite;
    edtValorMax.Color:= clWhite;
    relacaoData.Color:= clWhite;
  end
  else
  begin
    edtPonto.Enabled:= false;
    edtLimite.Enabled:= false;
    edtValorMax.Enabled:= false;
    relacaoData.Enabled:= false;
    relacaoData.Hint:= HintRelacaoTempo(relacaoData);
    edtPonto.Color:= clSilver;
    edtLimite.Color:= clSilver;
    edtValorMax.Color:= clSilver;
    relacaoData.Color:= clSilver;
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
    FrmPrincipal.LimparColunasFiltro(DBGrid,Layout);
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
          FrmPrincipal.buscaFiledGrid1(FieldName,Busca,Condicional,Layout,4,true);
        end;
      end;
    end;
  end;
end;

procedure TFrmPrincipal.configurarFiltro(Tipo: Integer; fieldName,ColunaIndex: String;
ColunaReadOnly: Boolean; actFiltroInserir,actGridASC,actGridDESC,
actSubstituirPor:TAction);
begin
  case Tipo of
    0:
    begin
      FieldName:= FieldName;
      FrmPrincipal.btnInserir.Action:= actFiltroInserir;
      FrmPrincipal.btnASC.Action:= actGridASC;
      FrmPrincipal.btnDESC.Action:= actGridDESC;
      //======================================================
      FrmPrincipal.MagicFieldName:= fieldName;
      FrmPrincipal.MagicIndexColuna:= ColunaIndex;
      FrmPrincipal.edtProcura.Text:= '';
      FrmPrincipal.edtLocalizar.Text:= '';
      FrmPrincipal.edtSubstituirPor.Text:= '';
      //======================================================
      FrmPrincipal.btnSubstituirPor.Enabled:= false;
      FrmPrincipal.btnSubstituirPor.ImageIndex:= 59;
      FrmPrincipal.edtLocalizar.Enabled:= false;
      FrmPrincipal.edtSubstituirPor.Enabled:= false;
      FrmPrincipal.edtLocalizar.Color:= clSilver;
      FrmPrincipal.edtSubstituirPor.Color:= clSilver;
    end;
    1:
    begin
      FrmPrincipal.btnInserir.Action:= actFiltroInserir;
      FrmPrincipal.btnASC.Action:= actGridASC;
      FrmPrincipal.btnDESC.Action:= actGridDESC;
      FrmPrincipal.btnSubstituirPor.Action:= actSubstituirPor;
      //======================================================
      FrmPrincipal.MagicFieldName:= fieldName;
      FrmPrincipal.MagicIndexColuna:= ColunaIndex;
      FrmPrincipal.edtProcura.Text:= '';
      FrmPrincipal.edtLocalizar.Text:= '';
      FrmPrincipal.edtSubstituirPor.Text:= '';
      //======================================================
      if ColunaReadOnly then
      begin
        actSubstituirPor.Enabled:= false;
        FrmPrincipal.edtLocalizar.Enabled:= false;
        FrmPrincipal.edtSubstituirPor.Enabled:= false;
        FrmPrincipal.edtLocalizar.Color:= clSilver;
        FrmPrincipal.edtSubstituirPor.Color:= clSilver;
      end
      else
      begin
        actSubstituirPor.Enabled:= true;
        FrmPrincipal.edtLocalizar.Enabled:= true;
        FrmPrincipal.edtSubstituirPor.Enabled:= true;
        FrmPrincipal.edtLocalizar.Color:= clWhite;
        FrmPrincipal.edtSubstituirPor.Color:= clWhite;
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

procedure TFrmPrincipal.CopiarProgramacao(DataProgramacao: String;DataSource: TDataSource);
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
    Dimensao:= FormatFloat('0.00km',
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
    FrmPrincipal.AutoFitStatusBar(StatusBarMapa);
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

procedure TFrmPrincipal.DBCheckBoxANPClick(Sender: TObject);
begin
  configPiroridade(edtPontoANP,edtLimiteANP,edtValorMaxANP,
  DBComboBoxANP,DBCheckBoxANP);
end;

procedure TFrmPrincipal.DBCheckBoxANVISAClick(Sender: TObject);
begin
  configPiroridade(edtPontoANVISA,edtLimiteANVISA,edtValorMaxANVISA,
  DBComboBoxANVISA,DBCheckBoxANVISA);
end;

procedure TFrmPrincipal.DBCheckBoxATMClick(Sender: TObject);
begin
  configPiroridade(edtPontoATM,edtLimiteATM,edtValorMaxATM,
  DBComboBoxATM,DBCheckBoxATM);
end;

procedure TFrmPrincipal.DBCheckBoxCorretivaClick(Sender: TObject);
begin
  configPiroridade(edtPontoCorretiva,edtLimiteCorretiva,edtValorMaxCorretiva,
  DBComboBoxCorretiva,DBCheckBoxCorretiva);
end;

procedure TFrmPrincipal.DBCheckBoxCustoClick(Sender: TObject);
begin
  if DBCheckBoxCusto.Checked then
  begin
    edtCustoMaximo.Enabled:= true;
    edtCustoMinimo.Enabled:= true;
    edtPontoMinimo.Enabled:= true;
    edtPontoMaximo.Enabled:= true;
    edtCustoMaximo.Color:= clWhite;
    edtCustoMinimo.Color:= clWhite;
    edtPontoMinimo.Color:= clWhite;
    edtPontoMaximo.Color:= clWhite;
  end
  else
  begin
    edtCustoMaximo.Enabled:= false;
    edtCustoMinimo.Enabled:= false;
    edtPontoMinimo.Enabled:= false;
    edtPontoMaximo.Enabled:= false;
    edtCustoMaximo.Color:= clSilver;
    edtCustoMinimo.Color:= clSilver;
    edtPontoMinimo.Color:= clSilver;
    edtPontoMaximo.Color:= clSilver;
  end;
end;

procedure TFrmPrincipal.DBCheckBoxECClick(Sender: TObject);
begin
  configPiroridade(edtPontoEC,edtLimiteEC,edtValorMaxEC,
  DBComboBoxEC,DBCheckBoxEC);
end;

procedure TFrmPrincipal.DBCheckBoxFOClick(Sender: TObject);
begin
  configPiroridade(edtPontoFO,edtLimiteFO,edtValorMaxFO,
  DBComboBoxFO,DBCheckBoxFO);
end;

procedure TFrmPrincipal.DBCheckBoxNR10Click(Sender: TObject);
begin
  configPiroridade(edtPontoNR10,edtLimiteNR10,edtValorMaxNR10,
  DBComboBoxNR10,DBCheckBoxNR10);
end;

procedure TFrmPrincipal.DBCheckBoxIBAMAClick(Sender: TObject);
begin
  configPiroridade(edtPontoIBAMA,edtLimiteIBAMA,edtValorMaxIBAMA,
  DBComboBoxIBAMA,DBCheckBoxIBAMA);
end;

procedure TFrmPrincipal.DBCheckBoxICPMAClick(Sender: TObject);
begin
  configPiroridade(edtPontoICPMA,edtLimiteICPMA,edtValorMaxICPMA,
  DBComboBoxICPMA,DBCheckBoxICPMA);
end;

procedure TFrmPrincipal.DBCheckBoxInspecaoClick(Sender: TObject);
begin
  configPiroridade(edtPontoInspecao,edtLimiteInspecao,edtValorMaxInspecao,
  DBComboBoxInspecao,DBCheckBoxInspecao);
end;

procedure TFrmPrincipal.DBCheckBoxINSPLANClick(Sender: TObject);
begin
  configPiroridade(edtPontoINSPLAN,edtLimiteINSPLAN,edtValorMaxINSPLAN,
  DBComboBoxINSPLAN,DBCheckBoxINSPLAN);
end;

procedure TFrmPrincipal.DBCheckBoxMarinhaClick(Sender: TObject);
begin
  configPiroridade(edtPontoMarinha,edtLimiteMarinha,edtValorMaxMarinha,
  DBComboBoxMarinha,DBCheckBoxMarinha);
end;

procedure TFrmPrincipal.DBCheckBoxMergulhoClick(Sender: TObject);
begin
  configPiroridade(edtPontoMergulho,edtLimiteMergulho,edtValorMaxMergulho,
  DBComboBoxMergulho,DBCheckBoxMergulho);
end;

procedure TFrmPrincipal.DBCheckBoxPreventivaClick(Sender: TObject);
begin
  configPiroridade(edtPontoPreventiva,edtLimitePreventiva,edtValorMaxPreventiva,
  DBComboBoxPreventiva,DBCheckBoxPreventiva);
end;

procedure TFrmPrincipal.DBCheckBoxROClick(Sender: TObject);
begin
  configPiroridade(edtPontoRO,edtLimiteRO,edtValorMaxRO,
  DBComboBoxRO,DBCheckBoxRO);
end;

procedure TFrmPrincipal.DBCheckBoxRTIAClick(Sender: TObject);
begin
  configPiroridade(edtPontoRTIA,edtLimiteRTIA,edtValorMaxRTIA,
  DBComboBoxRTIA,DBCheckBoxRTIA);
end;

procedure TFrmPrincipal.DBCheckBoxRTIBClick(Sender: TObject);
begin
  configPiroridade(edtPontoRTIB,edtLimiteRTIB,edtValorMaxRTIB,
  DBComboBoxRTIB,DBCheckBoxRTIB);
end;

procedure TFrmPrincipal.DBCheckBoxRTICClick(Sender: TObject);
begin
  configPiroridade(edtPontoRTIC,edtLimiteRTIC,edtValorMaxRTIC,
  DBComboBoxRTIC,DBCheckBoxRTIC);
end;

procedure TFrmPrincipal.DBCheckBoxRTIDClick(Sender: TObject);
begin
  configPiroridade(edtPontoRTID,edtLimiteRTID,edtValorMaxRTID,
  DBComboBoxRTID,DBCheckBoxRTID);
end;

procedure TFrmPrincipal.DBComboBoxANPChange(Sender: TObject);
begin
DBComboBoxANP.Hint:= HintRelacaoTempo(DBComboBoxANP);
end;

procedure TFrmPrincipal.DBComboBoxANVISAChange(Sender: TObject);
begin
DBComboBoxANVISA.Hint:= HintRelacaoTempo(DBComboBoxANVISA);
end;

procedure TFrmPrincipal.DBComboBoxATMChange(Sender: TObject);
begin
DBComboBoxATM.Hint:= HintRelacaoTempo(DBComboBoxATM);
end;

procedure TFrmPrincipal.DBComboBoxCorretivaChange(Sender: TObject);
begin
DBComboBoxCorretiva.Hint:= HintRelacaoTempo(DBComboBoxCorretiva);
end;

procedure TFrmPrincipal.DBComboBoxECChange(Sender: TObject);
begin
DBComboBoxEC.Hint:= HintRelacaoTempo(DBComboBoxEC);
end;

procedure TFrmPrincipal.DBComboBoxFOChange(Sender: TObject);
begin
DBComboBoxFO.Hint:= HintRelacaoTempo(DBComboBoxFO);
end;

procedure TFrmPrincipal.DBComboBoxNR10Change(Sender: TObject);
begin
DBComboBoxNR10.Hint:= HintRelacaoTempo(DBComboBoxNR10);
end;

procedure TFrmPrincipal.DBComboBoxIBAMAChange(Sender: TObject);
begin
DBComboBoxIBAMA.Hint:= HintRelacaoTempo(DBComboBoxIBAMA);
end;

procedure TFrmPrincipal.DBComboBoxICPMAChange(Sender: TObject);
begin
DBComboBoxICPMA.Hint:= HintRelacaoTempo(DBComboBoxICPMA);
end;

procedure TFrmPrincipal.DBComboBoxInspecaoChange(Sender: TObject);
begin
DBComboBoxInspecao.Hint:= HintRelacaoTempo(DBComboBoxInspecao);
end;

procedure TFrmPrincipal.DBComboBoxINSPLANChange(Sender: TObject);
begin
DBComboBoxINSPLAN.Hint:= HintRelacaoTempo(DBComboBoxINSPLAN);
end;

procedure TFrmPrincipal.DBComboBoxMarinhaChange(Sender: TObject);
begin
DBComboBoxMarinha.Hint:= HintRelacaoTempo(DBComboBoxMarinha);
end;

procedure TFrmPrincipal.DBComboBoxMergulhoChange(Sender: TObject);
begin
DBComboBoxMergulho.Hint:= HintRelacaoTempo(DBComboBoxMergulho);
end;

procedure TFrmPrincipal.DBComboBoxPreventivaChange(Sender: TObject);
begin
DBComboBoxPreventiva.Hint:= HintRelacaoTempo(DBComboBoxPreventiva);
end;

procedure TFrmPrincipal.DBComboBoxPreventivaKeyPress(Sender: TObject; var Key: Char);
begin
key:= #0;
end;

procedure TFrmPrincipal.DBComboBoxROChange(Sender: TObject);
begin
DBComboBoxRO.Hint:= HintRelacaoTempo(DBComboBoxRO);
end;

procedure TFrmPrincipal.DBComboBoxRTIAChange(Sender: TObject);
begin
DBComboBoxRTIA.Hint:= HintRelacaoTempo(DBComboBoxRTIA);
end;

procedure TFrmPrincipal.DBComboBoxRTIBChange(Sender: TObject);
begin
DBComboBoxRTIB.Hint:= HintRelacaoTempo(DBComboBoxRTIB);
end;

procedure TFrmPrincipal.DBComboBoxRTICChange(Sender: TObject);
begin
DBComboBoxRTIC.Hint:= HintRelacaoTempo(DBComboBoxRTIC);
end;

procedure TFrmPrincipal.DBComboBoxRTIDChange(Sender: TObject);
begin
DBComboBoxRTID.Hint:= HintRelacaoTempo(DBComboBoxRTID);
end;

procedure TFrmPrincipal.DBGridPalavraChaveCellClick(Column: TColumn);
begin
  if ((FrmPrincipal.logPerfil = 'Administrador')OR
  (FrmPrincipal.logPerfil = 'Programação')OR
  (FrmPrincipal.logPerfil = 'Supervisor')) then
  begin
    try
      if (Self.DBGridPalavraChave.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'booleanSelecao') then
      begin
        DBGridPalavraChave.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOQueryPalavraChave.Edit;
        FrmDataModule.DataSourcePalavraChave.DataSet.
        FieldByName('booleanSelecao').AsBoolean:= not Self.DBGridPalavraChave.SelectedField.AsBoolean;
        try
          FrmDataModule.ADOQueryPalavraChave.Post;
        except
          MessageBox(0,'Banco de dados alterado, necessário atualização!','Banco de Dados',MB_ICONINFORMATION);
          FrmDataModule.ADOQueryPalavraChave.Cancel;
        end;
      end
      else if (Self.DBGridPalavraChave.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'Kit_PS') then
      begin
        DBGridPalavraChave.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOQueryPalavraChave.Edit;
        FrmDataModule.DataSourcePalavraChave.DataSet.
        FieldByName('Kit_PS').AsBoolean:= not Self.DBGridPalavraChave.SelectedField.AsBoolean;
        try
          FrmDataModule.ADOQueryPalavraChave.Post;
        except
          MessageBox(0,'Banco de dados alterado, necessário atualização!','Banco de Dados',MB_ICONINFORMATION);
          FrmDataModule.ADOQueryPalavraChave.Cancel;
        end;
      end
      else
        DBGridPalavraChave.Options:=
        [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
    except
    end;
  end;
end;

procedure TFrmPrincipal.DBGridPalavraChaveDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridPalavraChave,ColunasLayoutPalavraChave,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'booleanSelecao')then   //  mudança no código
  begin
    Self.DBGridPalavraChave.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridPalavraChave.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end
end;

procedure TFrmPrincipal.DBGridPalavraChaveTitleClick(Column: TColumn);
begin
  FrmPrincipal.configurarFiltro(1,Column.FieldName,IntToStr(Column.Index),
  Column.ReadOnly,actFiltroInserir,actGridASC,actGridDESC,actSubstituirPor);
  //======================================================
  FrmPrincipal.titleGrid(ColunasLayoutPalavraChave,'Consulta',FrmDataModule.ADOQueryPalavraChave.SQL.Text);
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
          GravarDataSource.DataSet.FieldByName('StatusExecucao').AsString:= 'Executado';
          GravarDataSource.DataSet.FieldByName('MotivoProgramacao').AsString:= '';
          GravarDataSource.DataSet.FieldByName('MotivoNaoExecucao').AsString:= '';
          GravarDataSource.DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean:= false;
          GravarDataSource.DataSet.FieldByName('RT').AsString:= 'SEM RT';
          GravarDataSource.DataSet.FieldByName('TipoEmbarque').AsString:= 'SEM RT';
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

procedure TFrmPrincipal.edtCustoMaximoExit(Sender: TObject);
begin
  edtCustoMaximo.Text:= FormatFloat('R$ ###,##0.00',StrToFloat(edtCustoMaximo.Text));
end;

procedure TFrmPrincipal.edtCustoMinimoExit(Sender: TObject);
begin
  edtCustoMinimo.Text:= FormatFloat('R$ ###,##0.00',StrToFloat(edtCustoMinimo.Text));
end;

procedure TFrmPrincipal.edtProcuraKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    btnInserir.Click;
end;

procedure TFrmPrincipal.Embarcacao1Click(Sender: TObject);
begin
  if not Assigned(FrmEmbarcacao) then
    FrmEmbarcacao:= TFrmEmbarcacao.Create(Application)
  else
    FrmEmbarcacao.Show;
end;

procedure TFrmPrincipal.ExcelStringGrid(Grid: TStringGrid;
  NomePlanilha,Titulo1,Titulo2: String; LinhaInicial: Integer);
  var
    Excel, Sheet: Variant;
    NumLinhas, NumColunas, Linha, Coluna, i, j, tipoNumero: Integer;
    celula1: String;
begin
  Try
    Excel := CreateOleObject('Excel.Application');
    Excel.Workbooks.Add(1);
    Excel.Caption := NomePlanilha;
    Excel.Workbooks[1].WorkSheets[1].Name := NomePlanilha;
    Sheet := Excel.Workbooks[1].WorkSheets[NomePlanilha];
    // CONFIGURAÇÃO DO PROGRESSBAR
    FrmPrincipal.ProgressBarIncializa(Grid.ColCount*Grid.RowCount,'Gerando Excel...');
    Screen.Cursor := crHourGlass;
    // ==========CARREGAR OS DADOS===============
    NumLinhas := Grid.RowCount;
    NumColunas := Grid.ColCount;
    Linha := LinhaInicial;
    Coluna := 1;
    for j := 0 to Grid.ColCount - 1 do
      for i := 0 to Grid.RowCount - 1 do
      begin
        tipoNumero:= isNumeric(Grid.Cells[j, i]);
        case tipoNumero of
          0://texto
          begin
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j].NumberFormat:= '@';
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j] := Grid.Cells[j, i];
          end;
          1://Inteiro
          begin
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j].NumberFormat:= '0';
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j] := Grid.Cells[j, i];
          end;
          2://Float
          begin
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j].NumberFormat:= '0,00';
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j] := Grid.Cells[j, i];
          end;
          3://Data
          begin
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j].NumberFormat:= 'dd/mm/aaaa';
            Excel.Workbooks[1].Sheets[1].Cells[Linha+i,Coluna+j] := StrToDate(Grid.Cells[j, i]);
          end;
        end;
        {// COLORIR LINHAS ALTERNADAMENTE
        if (Linha + i) MOD 2 = 0 then
        begin
          // ANÁLISE 1
          Sheet.Range['A' + IntToStr(Linha + i),
            FrmPrincipal.RefToCell(NumColunas, Linha + i)].Interior.Color :=
            $00FFCF9F;
        end;}
        // ProgressBar
        FrmPrincipal.ProgressBarIncremento(1);
      end;
    // =========LETRAS CABEÇALHOS=========
    if Titulo1 <> '' then
      Excel.Workbooks[1].Sheets[1].Cells[1,1] := Titulo1;
    if Titulo2 <> '' then
      Excel.Workbooks[1].Sheets[1].Cells[2,1] := Titulo2;
    Sheet.Rows[LinhaInicial].Font.Bold := true;
    Sheet.Rows[LinhaInicial].Font.Italic := true;
    Sheet.Rows[Linha + NumLinhas + 2].Font.Bold := true;
    Sheet.Rows[Linha + NumLinhas + 2].Font.Italic := true;
    //Sheet.Columns[1].Font.Bold := true;
    //Sheet.Columns[1].Font.Italic := true;
    // ====================GRADE=================
    celula1 := FrmPrincipal.RefToCell(NumColunas, NumLinhas + Linha - 1);
    Sheet.Range['A'+IntToStr(LinhaInicial), celula1].Borders.LineStyle := 1;
    Sheet.Range['A'+IntToStr(LinhaInicial), celula1].Borders.Weight := 2;
    Sheet.Range['A'+IntToStr(LinhaInicial), celula1].Borders.ColorIndex := 1;
    Sheet.Range['A'+IntToStr(LinhaInicial), celula1].HorizontalAlignment := 3; // 3=Center
    // cor do cabeçalho
    Sheet.Range['A'+IntToStr(LinhaInicial), FrmPrincipal.RefToCell(NumColunas, Linha)].Interior.Color
      := clSilver; // Cor da Célula
    //===================================================
    Excel.Visible := true;
    Excel.Columns.Autofit;
  Except
    MessageBox(0, 'Esta operação não pode ser executada.' + #13 + #10 +
      'Verifique se o Microsoft Office Excel esta instalado na sua máquina.',
      'Excel', MB_ICONERROR);
  End;
  // Atualizar ProgressBar
  FrmPrincipal.ProgressBarAtualizar;
  Screen.Cursor := crDefault;
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

procedure TFrmPrincipal.ExportarExcel(TableName: String; BancoDados: Integer);
  Const
    acQuitSaveAll             = $00000001;
    acExport                  = $00000001;
    acSpreadsheetTypeExcel9   = $00000008;
    acSpreadsheetTypeExcel12  = $00000009;
  var
   LAccess : OleVariant;
   enderecoDB: String;
begin
  case BancoDados of
    0: //Colibri
    begin
      enderecoDB := FrmPrincipal.registroEndereco('Banco de dados');
    end;
    1: //Memoria
    begin
      enderecoDB := FrmPrincipal.registroEndereco('Banco de dados');
      enderecoDB:= ExtractFilePath(enderecoDB)+'\dbMemoria.mdb';
    end;
    2: //Consulta
    begin
      enderecoDB := FrmPrincipal.registroEndereco('Banco de dados');
      enderecoDB:= ExtractFilePath(enderecoDB)+'\dbConsulta.mdb';
    end;
  end;
  SaveDialog1.DefaultExt:= '*.xls';
  SaveDialog1.Filter:= 'Excel|*.xls|Excel|*.xlsx';
  if SaveDialog1.Execute then
  begin
    //create the COM Object
    LAccess := CreateOleObject('Access.Application');
    //open the access database
    LAccess.OpenCurrentDatabase(enderecoDB);
    //export the data
    LAccess.DoCmd.TransferSpreadsheet( acExport, acSpreadsheetTypeExcel9,
    TableName, SaveDialog1.FileName, True);
    LAccess.CloseCurrentDatabase;
    LAccess.Quit(1);
  end;
end;

procedure TFrmPrincipal.Fechar1Click(Sender: TObject);
begin
  Application.Terminate;
end;

procedure TFrmPrincipal.FiltrosTabela(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    i,index: Integer;
begin
  RLFiltrosTabela.RowCount:= Grid.Columns.Count+1;
  RLFiltrosTabela.ColCount:= 3;
  RLFiltrosTabela.Cells[0,0]:= 'Colunas';
  RLFiltrosTabela.Cells[1,0]:= 'Filtros';
  RLFiltrosTabela.Cells[2,0]:= 'Operador';
  for I := 0 to Grid.Columns.Count-1 do
  begin
    index:= indexLayoutFieldName(Grid.Columns[i].FieldName,ColunasLayout);
    RLFiltrosTabela.Cells[0,i+1]:= '';
    RLFiltrosTabela.Cells[1,i+1]:= '';
    RLFiltrosTabela.Cells[2,i+1]:= '';
    RLFiltrosTabela.Cells[0,i+1]:= ColunasLayout.Cells[1,index];
    RLFiltrosTabela.Cells[1,i+1]:= ColunasLayout.Cells[4,index];
    RLFiltrosTabela.Cells[2,i+1]:= ColunasLayout.Cells[5,index];
  end;
  AutoFitGrade(RLFiltrosTabela);
  RLFiltrosTabela.ColWidths[1]:= 300;
  actConfigFiltrosTabela.Execute;
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
  MatrizForaOperacao[8,i] SituacaoAcesso;
  MatrizForaOperacao[9,i] SituacaoDegraus;
  MatrizForaOperacao[10,i] DataRealizacaoDegraus;}
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
  actLOCAL.Enabled:= true;
  actREDE.Enabled:= false;
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

procedure TFrmPrincipal.GerarExcel(Tabela: TDBGrid; const nomeTabela: String);
var
  Excel, Sheet: Variant;
  i,j,matrizLinha,matrizColuna,selRegistro,tipoNumero,TipoArquivo: Integer;
  Celula, Titulo,campoProcura: String;
begin
  if Application.MessageBox(PChar(
  'Clique em SIM para exportar para arquivo Excel.'+#13+
  'Clique em NÃO para exportar para arquivo txt, delimitado por "@"'),'.::ATENÇÃO::.',36) = 6 then
    TipoArquivo:= 0
  else
    TipoArquivo:= 1;

  case TipoArquivo of
    0: //Excel
    begin
      selRegistro:= 0;
      try
        Excel := CreateOleObject('Excel.Application');
        Excel.Workbooks.Add;
        Excel.Workbooks[1].WorkSheets[1].Name := nomeTabela;
        Sheet := Excel.Workbooks[1].WorkSheets[nomeTabela];
        Excel.Caption := nomeTabela;
        //============================================
        selRegistro:= Tabela.DataSource.DataSet.RecNo;
        matrizColuna := Tabela.Columns.Count;
        matrizLinha :=  Tabela.DataSource.DataSet.RecordCount;
        Tabela.DataSource.Enabled:=false;
        // CONFIGURAÇÃO DO PROGRESSBAR
        ProgressBarIncializa(matrizColuna*matrizLinha,'Gerando Excel...');
        //============================================
        for j := 0 to matrizColuna - 1 do
        begin
          Titulo:= Tabela.Columns.Items[j].Title.Caption;
          campoProcura:= Tabela.Columns.Items[j].FieldName;
          // Preencher os titulos
          Excel.Workbooks[1].Sheets[1].Cells[1, j+1] := Titulo;
          for i := 1 to matrizLinha do
          begin
            // Preencher as linhas
            Tabela.DataSource.DataSet.RecNo:= i;
            tipoNumero:= isNumeric(Tabela.DataSource.DataSet.FieldByName(campoProcura).AsString);
            case tipoNumero of
              0://texto
              begin
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1].NumberFormat:= '@';
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1] :=
                Tabela.DataSource.DataSet.FieldByName(campoProcura).AsString;
              end;
              1://Inteiro
              begin
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1].NumberFormat:= '0';
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1] :=
                Tabela.DataSource.DataSet.FieldByName(campoProcura).AsInteger;
              end;
              2://Float
              begin
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1].NumberFormat:= '0,00';
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1] :=
                Tabela.DataSource.DataSet.FieldByName(campoProcura).AsFloat;
              end;
              3://Data
              begin
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1].NumberFormat:= 'dd/mm/aaaa';
                Excel.Workbooks[1].Sheets[1].Cells[i+1, j+1] :=
                Tabela.DataSource.DataSet.FieldByName(campoProcura).AsDateTime;
              end;
            end;
            // CONFIGURAÇÃO DAS CÉLULAS DAS LINHAS
            Sheet.Rows[i+1].VerticalAlignment := 2;
            // 1=Top - 2=Center - 3=Bottom
            Sheet.Rows[i+1].HorizontalAlignment := 3; // 3=Center }
            // Atualizar ProgressBar
            ProgressBarIncremento(1);
          end;
        end;
        // letras do cabeçalho
        Sheet.Rows[1].Font.Bold := true;
        Sheet.Rows[1].Font.Italic := true;
        Sheet.Rows[1].VerticalAlignment := 2; // 1=Top - 2=Center - 3=Bottom
        Sheet.Rows[1].HorizontalAlignment := 3; // 3=Center
        // desenha as grades da tabela
        Celula := RefToCell(matrizColuna, matrizLinha+1);
        Sheet.Range['A1', Celula].Borders.LineStyle := 1;
        Sheet.Range['A1', Celula].Borders.Weight := 2;
        Sheet.Range['A1', Celula].Borders.ColorIndex := 1;
        Sheet.Range['A1', RefToCell(matrizColuna, 1)].Interior.Color := clSilver;
        // Cor da Célula
        Excel.Visible := true;
        Excel.columns.Autofit;
      finally
        Tabela.DataSource.Enabled:=true;
        Tabela.DataSource.DataSet.RecNo:= selRegistro;
        // Atualizar ProgressBar
        ProgressBarAtualizar;
      end;
    end;
    1: //txt
    begin
      GeraTexto(Tabela.DataSource,'@',nomeTabela);
    end;
  end;
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

function TFrmPrincipal.GetDistanceBetween(lat1,long1,lat2,long2: Double): Double;
  var
    F,G,L : Double;
    SF, SG, SL,
    CF, CG, CL : Double;
    W1, W2 : Double;
    S, C : Double;
    O,R,D : Double;
    H1, H2 : Double;
    ff : Double;
begin
    ff := 1 / 298.257;
    F := (lat1 + lat2) / 2;
    G := (lat1 - lat2) / 2;
    L := (long1 - long2) / 2;

    SF := Sin(F*Pi/180);
    SG := Sin(G*Pi/180);
    SL := Sin(L*Pi/180);
    CF := Cos(F*Pi/180);
    CG := Cos(G*Pi/180);
    CL := Cos(L*Pi/180);

    W1 := sqr(SG * CL);
    W2 := sqr(CF * SL);
    S := W1 + W2;

    W1 := sqr(CG * CL);
    W2 := sqr(SF * SL);
    C := W1 + W2;
    O := ArcTan(Sqrt(S/C));
    R := Sqrt(S*C) / O;
    D := 2 * O * 6378.14;

    H1 := (3*R-1) / (2*C);
    H2 := (3*R+1) / (2*S);

    W1 := sqr(SF * CG) * H1 * ff + 1;
    W2 := sqr(CF * SG) * H2 * ff;

    result := D * (W1 - W2);// * 1.609344
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

procedure TFrmPrincipal.GridZebrado(Grid: TDBGrid; ColunasLayout: TStringGrid;
 State: TGridDrawState; const Rect: TRect; DataCol: Integer; Column: TColumn);
begin
  //DBGRIDZEBRADO
  if gdSelected in State then
    Grid.Canvas.Brush.Color:= $004080FF
  else
  begin
    if Grid.DataSource.DataSet.RecNo mod 2 = 0  then
      Grid.Canvas.Brush.Color:= clInfoBk
    else
      Grid.Canvas.Brush.Color:= clMoneyGreen;
  end;
  if ColunasLayout.Cells[4,indexLayoutFieldName(Column.FieldName,ColunasLayout)] <> '' then
  begin
    Column.Title.Font.Style := [FSBOLD];
    Column.Title.Font.Color:= clRed;
  end
  else
  begin
    Column.Title.Font.Style := [];
    Column.Title.Font.Color:= clBlack;
  end;
  Grid.Canvas.Font.Color:= clBlack;
  Grid.Canvas.FillRect(Rect);
  Grid.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

function TFrmPrincipal.GroupFieldDBGrid(SQLQuery,FieldName,FonteDB: String;
StatusBar: TStatusBar): TStringList;
  var
    Anterior,Atual,StrTotal: String;
    vetorCalc: array of Double;
    booleanCalculo: Boolean;
    I,NumRegistros: Integer;
    Soma: Double;
    Query: TADOQuery;
begin
  booleanCalculo:= true;
  Soma:= 0;
  NumRegistros:= 0;
  StatusBar.Panels[0].Text:= '';
  StatusBar.Panels[1].Text:= '';
  StatusBar.Panels[2].Text:= '';
  StatusBar.Panels[3].Text:= '';
  StatusBar.Panels[4].Text:= '';
  Result:= TStringList.Create;
  Query := TADOQuery.Create(self);
  try
    if FonteDB = 'Colibri' then
    begin
      Query.Connection := FrmDataModule.ADOConnectionColibri;
    end
    else if FonteDB = 'Consulta' then
    begin
      Query.Connection := FrmDataModule.ADOConnectionConsulta;
    end
    else if FonteDB = 'Memoria' then
    begin
      Query.Connection := FrmDataModule.ADOConnectionMemoria;
    end
    else if FonteDB = 'Importar' then
    begin
      Query.Connection := FrmDataModule.ADOConnectionImportar;
    end;
    Query.Close;
    Query.SQL.Clear;
    Query.SQL.Add(SQLQuery);
    Query.Open;
    Query.Sort:= FieldName+' ASC';
    NumRegistros:= Query.RecordCount;
    //======================================
    SetLength(vetorCalc,NumRegistros);
    Anterior:= 'zebrabrabadoscaraiodoido';
    I:= 0;
    actInterromper.Enabled:= true;
    StrTotal:= IntToSTr(NumRegistros);
    ProgBarMagicFiltro.Max:= NumRegistros;
    ProgBarMagicFiltro.Min:= 0;
    ProgBarMagicFiltro.Value:= 0;
    while not Query.Eof do
    begin
      //==================================
      Application.ProcessMessages;
      if Interromper then
      begin
        Interromper := False;
        Timer2.Enabled:= true;
        Exit;
      end;
      //==================================
      Atual:= Query.FieldByName(FieldName).AsString;
      if booleanCalculo then
      begin
        try
          vetorCalc[i]:= Query.FieldByName(FieldName).AsFloat;
          Soma:= Soma+Query.FieldByName(FieldName).AsFloat;
        except
          booleanCalculo:= false;
          Soma:= 0;
        end;
      end;
      if Anterior <> Atual then
      begin
        Result.Add(Atual);
        Anterior:= Atual;
      end;
      Query.Next;
      i:= i+1;
      ProgBarMagicFiltro.Value:= i;
      StatusBar.Panels[0].Text:= IntToStr(i)+' de '+StrTotal;
      AutoFitStatusBar(StatusBar);
    end;
    Timer2.Enabled:= true;
  except
  end;
  if booleanCalculo then
  begin
    try
      StatusBar.Panels[1].Text:= 'Máx.: '+FormatFloat('0',MaxValue(vetorCalc));
      StatusBar.Panels[2].Text:= 'Mín.: '+FormatFloat('0',MinValue(vetorCalc));
      StatusBar.Panels[3].Text:= '∑: '+FormatFloat('0',Soma);
      StatusBar.Panels[4].Text:= 'Média: '+FormatFloat('0.0',Soma/NumRegistros);
      AutoFitStatusBar(StatusBar);
    except
    end;
  end;
end;

function TFrmPrincipal.HintRelacaoTempo(Combo: TDBComboBox): String;
begin
  if Combo.Text = 'Data Base Início' then
  begin
    result:=
    'if (now >= DataBaseInicio) then'+#13+
    'begin'+#13+
    '   TotalDias:= DaysBetween(now,DataBaseInicio);'+#13+
    '   if (TotalDias >= LimiteDias) then'+#13+
    '      Result:= ValorMax'+#13+
    '   else'#13+
    '      Result:= RoundTo(((TotalDias*ValorMax)/LimiteDias),-3);'+#13+
    'end;';
  end
  else if Combo.Text = 'Data Base Fim' then
  begin
    result:=
    'if (now < incDay(DataBaseFim,-LimiteDias)) then'+#13+
    '   Result:= 0'+#13+
    'else'+#13+
    '   Result:= ValorMax;';
  end;
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

procedure TFrmPrincipal.ImportarPlanilhas1Click(Sender: TObject);
begin
  if not Assigned(FrmImportacao) then
    FrmImportacao:= TFrmImportacao.Create(Application)
  else
    FrmImportacao.Show;
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
  FrmDataModule.ADOConnectionMemoria.Connected:= false;
  FrmDataModule.ADOConnectionMemoria.Connected:= true;
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

procedure TFrmPrincipal.inserirProcura(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    i: Integer;
    txtString: String;
begin
  //Limpar simbolo de classificação ASC DESC
  limparTitleGrid(Grid);
  //Marcar check nas string já filtradas
  txtString:= '';
  if edtProcura.Text = '' then
  begin
    for i := 0 to CheckListBoxFiltroGRID.Items.Count-1 do
      if CheckListBoxFiltroGRID.Checked[i] then
      begin
        if txtString <> '' then
          txtString:= txtString+';'+CheckListBoxFiltroGRID.Items[i]
        else
          txtString:= CheckListBoxFiltroGRID.Items[i];
      end;
  end
  else
    txtString:= edtProcura.Text;
  //Gravar filtro na tabela de filtros
  for i := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[0,i] = MagicFieldName then
    begin
      ColunasLayout.Cells[4,i]:= txtString;
      case RadioGroupOperador.ItemIndex of
        0: ColunasLayout.Cells[5,i]:= 'Contem';
        1: ColunasLayout.Cells[5,i]:= 'Exato';
        2: ColunasLayout.Cells[5,i]:= 'Diferente';
      end;
      Break;
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
  actAjudaLimpar.Execute;
  FrmPrincipal.PanelMagicFiltro1.Visible:= false;
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

procedure TFrmPrincipal.moveBaixo(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    Titulo: String;
    Linha: Integer;
begin
  if FrmPrincipal.RLColunasAtivas.Row > 0 then
  begin
    Titulo:= FrmPrincipal.RLColunasAtivas.Cells[1,FrmPrincipal.RLColunasAtivas.Row];
    FrmPrincipal.DeleteRow(FrmPrincipal.RLColunasAtivas,FrmPrincipal.RLColunasAtivas.Row);
    Linha:= FrmPrincipal.InsertRow1(FrmPrincipal.RLColunasOpcoes,FrmPrincipal.RLColunasOpcoes.Row);
    FrmPrincipal.RLColunasOpcoes.Cells[1,Linha]:= Titulo;
  end;
end;

procedure TFrmPrincipal.moveCima(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    Titulo: String;
    Linha: Integer;
begin
  if FrmPrincipal.RLColunasOpcoes.Row > 0 then
  begin
    Titulo:= FrmPrincipal.RLColunasOpcoes.Cells[1,FrmPrincipal.RLColunasOpcoes.Row];
    FrmPrincipal.DeleteRow(FrmPrincipal.RLColunasOpcoes,FrmPrincipal.RLColunasOpcoes.Row);
    Linha:= FrmPrincipal.InsertRow1(FrmPrincipal.RLColunasAtivas,FrmPrincipal.RLColunasAtivas.Row);
    FrmPrincipal.RLColunasAtivas.Cells[1,Linha]:= Titulo;
  end;
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
    FrmTelaAbertura.MemoMensagem.Lines.Add('--------------------------------------------------------------------------------------------------------');
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

procedure TFrmPrincipal.PainelCancelamentoMudanca(Fonte: Integer; Titulo: String);
begin
  FrmPrincipal.SetUpColunasLayout(DBGridPalavraChave, ColunasLayoutPalavraChave);
  actProcurarPalavraChave.Execute;
  FrmPrincipal.PanelTituloAjuda1.Caption:= Titulo;
  FrmPrincipal.actAjudaLimpar.Execute;
  FrmPrincipal.PanelCancelamentoMudanca.Align:= alClient;
  FrmPrincipal.PanelCancelamentoMudanca.Visible:= true;
  FrmPrincipal.PanelAjuda1.Top:= 100;
  FrmPrincipal.PanelAjuda1.Left:= 200;
  FrmPrincipal.PanelAjuda1.Height:= 440;
  FrmPrincipal.PanelAjuda1.Width:= 500;
  FrmPrincipal.PanelAjuda1.Visible:= true;
  FrmPrincipal.RadioGroupFonteCancelamento.ItemIndex:= Fonte;
  if ((FrmPrincipal.logPerfil = 'Administrador')OR
  (FrmPrincipal.logPerfil = 'Programação')OR
  (FrmPrincipal.logPerfil = 'Supervisor')) then
  begin
    DBNavigatorPalavraChave.Enabled:= true;
    DBGridPalavraChave.ReadOnly:= false;
    actInserirRegistro.Enabled:= true;
    actCancelarRegistro.Enabled:= true;
  end
  else
  begin
    DBNavigatorPalavraChave.Enabled:= false;
    DBGridPalavraChave.ReadOnly:= true;
    actInserirRegistro.Enabled:= false;
    actCancelarRegistro.Enabled:= false;
  end;
end;

procedure TFrmPrincipal.PainelDuplicados(DataProcura: String);
begin
  ProgramacaoDuplicados(RLDuplicados,DataProcura);
  actAjudaLimpar.Execute;
  PanelTituloAjuda1.Caption:= 'Executantes Duplicados (Transbordos)';
  PanelDuplicados.Align:= alClient;
  PanelDuplicados.Visible:= true;
  PanelAjuda1.Visible:= true;
  PanelAjuda1.Width:= 770;
  PanelAjuda1.Height:= 450;
  PanelAjuda1.Top:= 180;
  PanelAjuda1.Left:= 400;
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

procedure TFrmPrincipal.ProgramarExecutante(Origem: String;idProgramacaoDiaria: Integer;
QueryExecutante,QueryProgramacao: TADOQuery; DataSourceExecutante,DataSourceProgramacao: TDataSource);
begin
  if idProgramacaoDiaria <> 0 then
  begin
    try
      QueryExecutante.Insert;
      DataSourceExecutante.DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger:=idProgramacaoDiaria;
      DataSourceExecutante.DataSet.FieldByName('Origem').AsString:= Origem;
      DataSourceExecutante.DataSet.FieldByName('StatusProgramacao').AsString:= 'Aprovado';
      DataSourceExecutante.DataSet.FieldByName('StatusExecucao').AsString:='Executado';
      DataSourceExecutante.DataSet.FieldByName('RT').AsString:= 'SEM RT';
      DataSourceExecutante.DataSet.FieldByName('TipoEmbarque').AsString:= 'SEM RT';
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

procedure TFrmPrincipal.ProgramacaoDuplicados(RLGrid: TStringGrid; DataProcura: String);
  var
    SQLBase,NomeExecutante1,NomeExecutante2,Origem1,Origem2,Destino1,Destino2,
    TipoEtapaServico1,TipoEtapaServico2,Funcao1,Funcao2,Empresa1,Empresa2: String;
    linha: Integer;
begin
  RLGrid.RowCount:= 1;
  RLGrid.ColCount:= 6;
  RLGrid.Rows[0].Clear;
  RLGrid.Cells[0,0]:= 'Origem';
  RLGrid.Cells[1,0]:= 'Destino';
  RLGrid.Cells[2,0]:= 'Nome do Executante';
  RLGrid.Cells[3,0]:= 'Tipo de Etapa de Serviço';
  RLGrid.Cells[4,0]:= 'Função';
  RLGrid.Cells[5,0]:= 'Empresa';
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
  NomeExecutante1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('NomeExecutante').AsString;
  Origem1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Origem').AsString;
  Destino1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtDestino').AsString;
  TipoEtapaServico1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtTipoEtapaServico').AsString;
  Funcao1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Funcao').AsString;
  Empresa1:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Empresa').AsString;
  FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    NomeExecutante2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('NomeExecutante').AsString;
    Origem2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Origem').AsString;
    Destino2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtDestino').AsString;
    TipoEtapaServico2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('txtTipoEtapaServico').AsString;
    Funcao2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Funcao').AsString;
    Empresa2:= FrmDataModule.ADOQueryTemporarioDBColibri.FieldByName('Empresa').AsString;
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
      //========================================================
      linha:= RLGrid.RowCount;
      RLGrid.RowCount:= RLGrid.RowCount+1;
      RLGrid.Cells[0,linha]:= Origem2;
      RLGrid.Cells[1,linha]:= Destino2;
      RLGrid.Cells[2,linha]:= NomeExecutante2;
      RLGrid.Cells[3,linha]:= TipoEtapaServico2;
      RLGrid.Cells[4,linha]:= Funcao2;
      RLGrid.Cells[5,linha]:= Empresa2;
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
  FrmPrincipal.AutoFitGrade(RLGrid);
  try
    RLGrid.FixedRows:= 1;
  except
  end;
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

procedure TFrmPrincipal.RLColunasAtivasDblClick(Sender: TObject);
begin
btnBaixo.Click
end;

procedure TFrmPrincipal.RLColunasAtivasDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  FrmPrincipal.ConfigGridLayout(RLColunasAtivas,ACol,ARow,Rect);
end;

procedure TFrmPrincipal.RLColunasOpcoesDblClick(Sender: TObject);
begin
btnCima.Click
end;

procedure TFrmPrincipal.RLColunasOpcoesDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
begin
  FrmPrincipal.ConfigGridLayout(RLColunasOpcoes,ACol,ARow,Rect);
end;

procedure TFrmPrincipal.RLDuplicadosFixedCellClick(Sender: TObject; ACol,
  ARow: Integer);
begin
  FrmPrincipal.clasifica(RLDuplicados,ACol,true);
end;

procedure TFrmPrincipal.RLFiltrosTabelaSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
  var
    R: TRect;
begin
  if ((ACol = 2) AND(ARow > 0)) then
  begin
    if RLFiltrosTabela.Cells[1, ARow] = '' then
    begin
      RLFiltrosTabela.Cells[2, ARow] := '';
    end
    else
    begin
      R := RLFiltrosTabela.CellRect(ACol, ARow);
      R.Left := R.Left + RLFiltrosTabela.Left;
      R.Right := R.Right + RLFiltrosTabela.Left;
      R.Top := R.Top + RLFiltrosTabela.Top;
      R.Bottom := R.Bottom + RLFiltrosTabela.Top;
      ComboBoxOperador.Left := R.Left + 1;
      ComboBoxOperador.Top := R.Top + 1;
      ComboBoxOperador.Width := (R.Right + 1) - R.Left;
      ComboBoxOperador.Height := (R.Bottom + 1) - R.Top;
      ComboBoxOperador.Visible := True;
      FrmPrincipal.selComboBox(ComboBoxOperador,(RLFiltrosTabela.Cells[2, ARow]));
      ComboBoxOperador.SetFocus;
    end;
  end;
  CanSelect := True;
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

procedure TFrmPrincipal.RT_ProgramacaoExecutante(idProgramacaoExecutante: Integer;
RT, TipoEmbarque: String);
begin
  try
    //Procurar Programação do Executante
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('RT').AsString:= RT;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('TipoEmbarque').AsString:= TipoEmbarque;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
  except
  end;
end;

procedure TFrmPrincipal.SalvarBancoDadosComo1Click(Sender: TObject);
var
  arquivo_original,enderecoConsulta_original,enderecoConsulta,
  enderecoMemoria_original,enderecoMemoria: string;
begin
  // diretorio e nome do arquivo original
  arquivo_original := registroEndereco('Banco de dados');
  enderecoConsulta_original:= ExtractFilePath(arquivo_original)+'dbConsulta.mdb';
  enderecoMemoria_original:= ExtractFilePath(arquivo_original)+'dbMemoria.mdb';
  //====================================================
  SaveDialog1.Filter:= 'Arquivo Colibri|*.colibri';
  SaveDialog1.FileName:= arquivo_original;
  SaveDialog1.DefaultExt:= '*.colibri';
  if SaveDialog1.Execute then
  begin
      try
        if CopyFile(PChar(arquivo_original), PChar(SaveDialog1.FileName), false) then
        begin
          FrmPrincipal.ProgressBarIncializa(5,'Copiando banco de dados');
          registroEscrever('Banco de dados', SaveDialog1.FileName);
          FrmPrincipal.ProgressBarAtualizar;
          //Carregar banco de dados
          conectarBD(SaveDialog1.FileName,FrmDataModule.ADOConnectionColibri);
          FrmPrincipal.ProgressBarIncremento(1);
          //Banco de dados de Consulta
          enderecoConsulta:= ExtractFileDir(SaveDialog1.FileName)+'\dbConsulta.mdb';
          enderecoMemoria:= ExtractFileDir(SaveDialog1.FileName)+'\dbMemoria.mdb';
          if CopyFile(PChar(enderecoConsulta_original), PChar(enderecoConsulta), false) then
            FrmPrincipal.conectarBDDireto(enderecoConsulta,FrmDataModule.ADOConnectionConsulta);
          FrmPrincipal.ProgressBarIncremento(1);
          if CopyFile(PChar(enderecoMemoria_original), PChar(enderecoMemoria), false) then
            FrmPrincipal.conectarBDDireto(enderecoMemoria,FrmDataModule.ADOConnectionMemoria);
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

procedure TFrmPrincipal.GravarRLColunas(NomeArquivoMemo:string);
  var
    i: Integer;
begin
  MemoPrincipal.Lines.Clear;
  for i := 1 to RLColunasAtivas.RowCount-1 do
    MemoPrincipal.Lines.Add('TAB1;'+RLColunasAtivas.Cells[1,i]);
  for i := 1 to RLColunasOpcoes.RowCount-1 do
    MemoPrincipal.Lines.Add('TAB2;'+RLColunasOpcoes.Cells[1,i]);

  MemoPrincipal.Lines.SaveToFile(
  ExtractFilePath(Application.ExeName)+'Layout\'+NomeArquivoMemo);
end;

procedure TFrmPrincipal.SalvarRLColunasCarregaGRID(Conection: TADOConnection;
NomeArquivoMemo,NomeTabela: String; Grid: TDBGrid; ColunasLayout: TStringGrid);
begin
  FrmPrincipal.GravarRLColunas(NomeArquivoMemo);
  FrmPrincipal.CarregarColunasGRID(Conection,NomeTabela,Grid,ColunasLayout);
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

procedure TFrmPrincipal.selecaoDBGrid(Grid: TDBGrid; selecao: boolean);
  var
    selRegistro: Integer;
begin
  selRegistro:= Grid.DataSource.DataSet.RecNo;
  Grid.DataSource.Enabled:= false;
  Grid.DataSource.DataSet.First;
  while not Grid.DataSource.DataSet.Eof do
  begin
    Grid.DataSource.DataSet.Edit;
    Grid.DataSource.DataSet.FieldByName('booleanSelecao').AsBoolean:= selecao;
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
  FrmPrincipal.AutoFitStatusBar(Status);
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
        FrmPrincipal.AutoFitStatusBar(StatusBarMapa);
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

procedure TFrmPrincipal.LimparColunasFiltro(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    i: Integer;
begin
  limparTitleGrid(Grid);
  for I := 0 to ColunasLayout.RowCount-1 do
  begin
    ColunasLayout.Cells[4,i]:= '';
    ColunasLayout.Cells[5,i]:= '';
  end;
end;

procedure TFrmPrincipal.SetUpColunasLayout(Grid: TDBGrid; ColunasLayout: TStringGrid);
  var
    i: Integer;
    alinha: TAlignment;
    strTitulo: String;
begin
  ColunasLayout.RowCount:= Grid.Columns.Count;
  ColunasLayout.ColCount:= 7;
  for i := 0 to ColunasLayout.RowCount-1 do
  begin
    //FieldName
    ColunasLayout.Cells[0,i]:= Grid.Columns[i].FieldName;
    //Titulo da coluna
    strTitulo:= FrmPrincipal.DeleteChar(char(9650),Grid.Columns[i].Title.Caption);
    strTitulo:= FrmPrincipal.DeleteChar(char(9660),strTitulo);
    strTitulo:= FrmPrincipal.DeleteChar(char(8722),strTitulo);
    ColunasLayout.Cells[1,i]:= strTitulo;
    //Filtro
    ColunasLayout.Cells[4,i]:= '';
    //Operador
    ColunasLayout.Cells[5,i]:= '';
    //Largura da coluna
    ColunasLayout.Cells[6,i]:= IntToStr(Grid.Columns[i].Width);
    //Alinhamento
    alinha:= Grid.Columns[i].Alignment;
    if alinha = taCenter then
      ColunasLayout.Cells[2,i]:= 'CENTER'
    else if alinha = taLeftJustify then
      ColunasLayout.Cells[2,i]:= 'LEFT'
    else if alinha = taRightJustify then
      ColunasLayout.Cells[2,i]:= 'RIGHT';
    //ReadOnly
    if Grid.Columns[i].ReadOnly then
      ColunasLayout.Cells[3,i]:= 'TRUE'
    else
      ColunasLayout.Cells[3,i]:= 'FALSE';
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

procedure TFrmPrincipal.SpeedButton1Click(Sender: TObject);
begin
  PanelMagicFiltro1.Visible:= false;
end;

procedure TFrmPrincipal.SpeedButton4Click(Sender: TObject);
begin
  PanelAjuda1.Visible:= false;
  PanelTituloAjuda1.Color:= clGradientActiveCaption;
end;

function TFrmPrincipal.SQLStringFiltroTabela(ColunasLayout: TStringGrid;BooleanWHERE: Boolean): String;
  var
    SQLString: String;
    i: Integer;
begin
  SQLString:= '';
  for i := 0 to ColunasLayout.RowCount do
  begin
    if ColunasLayout.Cells[4,i] <> '' then
    begin
      SQLString:= FrmPrincipal.palavraBusca(SQLString,ColunasLayout.Cells[0,i],
      ColunasLayout.Cells[5,i],ColunasLayout.Cells[4,i],'AND');
    end;
  end;
  if ((SQLString <> '')AND(BooleanWHERE)) then
    SQLString:= ' WHERE '+SQLString;

  Result:= SQLString;
end;

procedure TFrmPrincipal.PanelTituloAjuda1MouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmPrincipal.PanelTituloAjuda1MouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    PanelAjuda1.Left := PanelAjuda1.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda1.Top := PanelAjuda1.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmPrincipal.PanelTituloAjuda1MouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    FrmPrincipal.Capturing := false;
    PanelAjuda1.Left := PanelAjuda1.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda1.Top := PanelAjuda1.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
    if PanelAjuda1.Top < 50 then
      PanelAjuda1.Top := 50;
  end;
end;

procedure TFrmPrincipal.PanelTituloMagicMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    PanelMagicFiltro1.Left := PanelMagicFiltro1.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelMagicFiltro1.Top := PanelMagicFiltro1.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmPrincipal.PanelTituloMagicMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    FrmPrincipal.Capturing := false;
    PanelMagicFiltro1.Left := PanelMagicFiltro1.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelMagicFiltro1.Top := PanelMagicFiltro1.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
    if PanelMagicFiltro1.Top < 50 then
      PanelMagicFiltro1.Top := 50;
  end;
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
    if btnSQLFiltroTabela.Down then
      SQLBase:= MemoSQL.Text;
    //=================================
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(SQLBase);
    sourceQuery.Open;
    StatusBar.Panels[0].Text:= 'N° Registros: '+
    IntToStr(sourceQuery.RecordCount);
    FrmPrincipal.AutoFitStatusBar(StatusBar);
    MemoSQL.Text:= SQLBase;
  except
    sourceQuery.Connection.Close;
    sourceQuery.Connection.Open;
    sourceQuery.Close;
    sourceQuery.SQL.Clear;
    sourceQuery.SQL.Add(SQLBase);
    sourceQuery.Open;
    StatusBar.Panels[0].Text:= 'N° Registros: '+
    IntToStr(sourceQuery.RecordCount);
    FrmPrincipal.AutoFitStatusBar(StatusBar);
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

procedure TFrmPrincipal.SubstituirPor(Grid: TDBGrid; ADOQuery: TADOQuery;DataSource: TDataSource);
  var
    txtProcura,FieldName,txtCorrigido: String;
    i: Integer;
begin
  try
    FieldName:= MagicFieldName;
    DataSource.Enabled:= false;
    ADOQuery.First;
    ProgBarMagicFiltro.Max:= ADOQuery.RecordCount;
    ProgBarMagicFiltro.Min:= 0;
    ProgBarMagicFiltro.Value:= 0;
    i:= 0;
    case RadioGroupSubstituir.ItemIndex of
      0:
      begin
        while not ADOQuery.Eof do
        begin
          txtProcura:= DataSource.DataSet.FieldByName(FieldName).AsString;
          if txtProcura = edtLocalizar.Text then
          begin
            try
              ADOQuery.Edit;
              DataSource.DataSet.FieldByName(FieldName).AsString:= edtSubstituirPor.Text;
              ADOQuery.Post;
            except
              ADOQuery.Cancel;
              ADOQuery.Refresh;
            end;
          end;
          ADOQuery.Next;
          i:= i+1;
          ProgBarMagicFiltro.Value:= i;
        end;
        Timer2.Enabled:= true;
      end;
      1:
      begin
        while not ADOQuery.Eof do
        begin
          txtProcura:= DataSource.DataSet.FieldByName(FieldName).AsString;
          //txtCorrigido:= SubstituirCaracteres(txtProcura,edtLocalizar.Text,edtSubstituirPor.Text);
          txtCorrigido:= StringReplace(txtProcura,edtLocalizar.Text,edtSubstituirPor.Text,[rfReplaceAll]);
          if txtCorrigido<>txtProcura then
          begin
            try
              ADOQuery.Edit;
              DataSource.DataSet.FieldByName(FieldName).AsString:= txtCorrigido;
              ADOQuery.Post;
            except
              ADOQuery.Cancel;
              ADOQuery.Refresh;
            end;
          end;
          ADOQuery.Next;
          i:= i+1;
          ProgBarMagicFiltro.Value:= i;
        end;
        Timer2.Enabled:= true;
      end;
    end;
  finally
    DataSource.Enabled:= true;
    ADOQuery.First;
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

procedure TFrmPrincipal.Timer2Timer(Sender: TObject);
begin
  ProgBarMagicFiltro.Value := ProgBarMagicFiltro.Min;
  actInterromper.Enabled:= false;
  Interromper:= false;
  StatusBarMagicFiltro.Panels[0].Text:= '';
  AutoFitStatusBar(StatusBarMagicFiltro);
  Timer2.Enabled:= false;
end;

procedure TFrmPrincipal.titleGrid(ColunasLayout: TStringGrid;
FonteDB,SQLQuery: String);
  var
    txtProcura,titulo: String;
    i: Integer;
begin
  actConfigMagicPanel.Execute;
  CheckListBoxFiltroGRID.Items.Clear;
  for i := 0 to ColunasLayout.RowCount-1 do
  begin
    if ColunasLayout.Cells[0,i] = MagicFieldName then
    begin
      txtProcura:= ColunasLayout.Cells[4,i];
      titulo:= ColunasLayout.Cells[1,i];
      Break;
    end;
  end;
  PanelTituloMagic.Caption:= titulo;
  PanelMagicFiltro1.Visible := True;
  MagicSQLQuery:= SQLQuery;
  MagicFonteDB:= FonteDB;
  MagicTextoProcura:= txtProcura;
  if CheckBoxListarAutomatico.Checked then
    actListar.Execute;
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

function TFrmPrincipal.verificaTextoLongoCarteira(
  OrdemManutencao: String): WideString;
begin
  if ((OrdemManutencao = '')or(OrdemManutencao = '0')) then
    OrdemManutencao:= '999999999';

  try
    FrmDataModule.ADOQueryTextoLongoOM.Active:= false;
    FrmDataModule.ADOQueryTextoLongoOM.Parameters.Items[0].Value:= OrdemManutencao;
    FrmDataModule.ADOQueryTextoLongoOM.Active:= true;

    if not FrmDataModule.ADOQueryTextoLongoOM.IsEmpty then
      Result:= FrmDataModule.DataSourceTextoLongoOM.
      DataSet.FieldByName('TextoLongoCarteira').AsWideString
    else
      Result:= '';
  except
    Result:= 'NÃO DEFINIDO';
  end;
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
