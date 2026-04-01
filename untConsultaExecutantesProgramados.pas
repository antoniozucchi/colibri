unit untConsultaExecutantesProgramados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ComCtrls, Vcl.ToolWin, Vcl.CheckLst, Vcl.ExtCtrls, Vcl.Grids, Vcl.DBGrids,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,DateUtils,
  Vcl.ExtDlgs, Vcl.Menus, Vcl.AppEvnts,ComOBJ, ADODB,
  untDBGridFilter, System.StrUtils, Vcl.DBCtrls, System.Math,
  System.Generics.Collections, System.Generics.Defaults, uProgramacaoRTUtils, System.UITypes,
  uSimulacaoLogistica, uAccessDBUtils,uZucchi, Vcl.Mask;

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
    actImprmir: TAction;
    actNumDias: TAction;
    actStatusSELECIONADO: TAction;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    DBGridExecutantesProgramados: TFilterDBGrid;
    StatusBarExecutantes: TStatusBar;
    RLTemporario: TStringGrid;
    ToolBar1: TToolBar;
    BitBtn3: TBitBtn;
    SavePictureDialog1: TSavePictureDialog;
    PopupMenuStatus: TPopupMenu;
    CalcularStatusExecSelecionado1: TMenuItem;
    CalcularStatusExec1: TMenuItem;
    actStatusTODOS: TAction;
    SaveDialog1: TSaveDialog;
    ColunasLayoutExecutanteProgramado: TStringGrid;
    CheckBoxOrigemDestino: TCheckBox;
    ToolButton11: TToolButton;
    actTransbordo: TAction;
    actMotivo: TAction;
    MemoSAP: TMemo;
    actGerarMultiplasRTs: TAction;
    ToolButton4: TToolButton;
    actConfigurarRT: TAction;
    btnLayoutExecutante: TToolButton;
    btnClearFiltroExecutante: TToolButton;
    btnExcelExecutante: TToolButton;
    btnSelAll: TToolButton;
    btnSelClear: TToolButton;
    actCancelarRTForcado: TAction;
    TabSheet2: TTabSheet;
    ToolBar3: TToolBar;
    BitBtn10: TBitBtn;
    btnClearFiltroRT: TToolButton;
    actProcurarProgramacaoRT: TAction;
    StatusBarGestaoRT: TStatusBar;
    DBGridProgramcaoRT: TFilterDBGrid;
    btnCancelarForcado: TBitBtn;
    actPrepararDadosRT: TAction;
    BitBtn9: TBitBtn;
    btnExcelRT: TToolButton;
    ColunasLayoutRT: TStringGrid;
    btnLaypoutRT: TToolButton;
    actProgramacaoRTxProgramacaoExecutante: TAction;
    BitBtn2: TBitBtn;
    PopupMenuSelAll: TPopupMenu;
    PopupMenuSelClear: TPopupMenu;
    actSelAllSelecao: TAction;
    actSelAllRecolhimento: TAction;
    actSelClearSelecao: TAction;
    actSelClearRecolhimento: TAction;
    MarcartodosRecolhimento1: TMenuItem;
    MarcartodosSeleo1: TMenuItem;
    DesmarcartodosRecolhimento1: TMenuItem;
    DesmarcartodosSeleo1: TMenuItem;
    TabSheet4: TTabSheet;
    DBGridRTSapImport: TFilterDBGrid;
    ToolBar4: TToolBar;
    btnClearFiltroSAPImport: TToolButton;
    btnExcelSAPImport: TToolButton;
    btnLayoutSAPImport: TToolButton;
    ColunasLayoutSAPImport: TStringGrid;
    actProcurarSAPImport: TAction;
    StatusBarSAPImport: TStatusBar;
    BitBtn7: TBitBtn;
    actImportRTSAP: TAction;
    PopupMenuFuncoes: TPopupMenu;
    AvaliarnecessidadedecriaodeRTs1: TMenuItem;
    Aplicarregrasautomticas1: TMenuItem;
    SincronizarProgramaodeExecutantecomasRTsExistentesnoSAP1: TMenuItem;
    actHistorizarProgramacao_DadosSAP: TAction;
    CarregaronmerodaRTsimportadasdoSAPparaastabelasRegistrodeRTseExecutantesProgramados1: TMenuItem;
    Panel1: TPanel;
    Panel2: TPanel;
    Panel5: TPanel;
    CarregardadosdatabelaRegistrosdeRTparaatabelaExecutantesProgramados1: TMenuItem;
    Preparao1: TMenuItem;
    actConcliliar1: TMenuItem;
    actReprocessar1: TMenuItem;
    actAuditarSemExecutante1: TMenuItem;
    actResumoAuditoriaSemExecutante1: TMenuItem;
    BitBtn11: TBitBtn;
    ConciliarRTsOrfas: TAction;
    actRT_TMIB: TAction;
    BitBtn13: TBitBtn;
    RLImpressao: TStringGrid;
    BitBtn12: TBitBtn;
    btnFiltros: TToolButton;
    actProcurarSomenteTransbordos: TAction;
    PopupMenuFiltros: TPopupMenu;
    actProcurarSomenteTransbordos1: TMenuItem;
    actLimpar: TAction;
    BitBtn1: TBitBtn;
    TabSheet3: TTabSheet;
    DBGridResumoFrequencia: TFilterDBGrid;
    RLLayoutBuscaEmbarque: TStringGrid;
    ToolBar2: TToolBar;
    btnClearFiltroBuscaEmbarque: TToolButton;
    btnExcelBuscaEmbarque: TToolButton;
    btnLauoutBuscaEmbarque: TToolButton;
    Panel6: TPanel;
    actProcurarBuscaEmbarque: TAction;
    ToolButton1: TToolButton;
    StatusBarFrequenciaResumo: TStatusBar;
    TabSheet5: TTabSheet;
    PanelSimulacao: TPanel;
    ToolBarSimulacao: TToolBar;
    btnSimulacaoPadrao: TBitBtn;
    btnSimulacaoBaseReal: TBitBtn;
    btnSimulacaoRodar: TBitBtn;
    btnSimulacaoComparar: TBitBtn;
    btnSimulacaoExportar: TBitBtn;
    GridSimulacaoParametros: TStringGrid;
    SplitterSimulacao: TSplitter;
    MemoSimulacaoRelatorio: TMemo;
    StatusBarSimulacao: TStatusBar;
    procedure actProcurarProgramacaoExecutanteExecute(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridExecutantesProgramadosDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure edtExecutanteKeyPress(Sender: TObject; var Key: Char);
    procedure actImprmirExecute(Sender: TObject);
    procedure dataInicioKeyPress(Sender: TObject; var Key: Char);
    procedure dataFimKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridExecutantesProgramadosCellClick(Column: TColumn);
    procedure actNumDiasExecute(Sender: TObject);
    procedure actStatusSELECIONADOExecute(Sender: TObject);
    procedure actStatusTODOSExecute(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    procedure actTransbordoExecute(Sender: TObject);
    procedure actMotivoExecute(Sender: TObject);
    procedure RLImpressaoFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure actGerarMultiplasRTsExecute(Sender: TObject);
    procedure actConfigurarRTExecute(Sender: TObject);
    procedure actCancelarRTForcadoExecute(Sender: TObject);
    procedure actProcurarProgramacaoRTExecute(Sender: TObject);
    procedure actPrepararDadosRTExecute(Sender: TObject);
    procedure actProgramacaoRTxProgramacaoExecutanteExecute(Sender: TObject);
    procedure dataInicioCloseUp(Sender: TObject);
    procedure actSelAllSelecaoExecute(Sender: TObject);
    procedure actSelAllRecolhimentoExecute(Sender: TObject);
    procedure actSelClearSelecaoExecute(Sender: TObject);
    procedure actSelClearRecolhimentoExecute(Sender: TObject);
    procedure actProcurarSAPImportExecute(Sender: TObject);
    procedure actImportRTSAPExecute(Sender: TObject);
    procedure actHistorizarProgramacao_DadosSAPExecute(Sender: TObject);
    procedure dataFimCloseUp(Sender: TObject);
    procedure DBGridProgramcaoRTDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure ConciliarRTsOrfasExecute(Sender: TObject);
    procedure actRT_TMIBExecute(Sender: TObject);
    procedure actProcurarSomenteTransbordosExecute(Sender: TObject);
    procedure actLimparExecute(Sender: TObject);
    procedure ToolButton1Click(Sender: TObject);
    procedure DBEdit1Change(Sender: TObject);
    procedure actProcurarBuscaEmbarqueExecute(Sender: TObject);
    procedure btnSimulacaoPadraoClick(Sender: TObject);
    procedure btnSimulacaoBaseRealClick(Sender: TObject);
    procedure btnSimulacaoRodarClick(Sender: TObject);
    procedure btnSimulacaoCompararClick(Sender: TObject);
    procedure btnSimulacaoExportarClick(Sender: TObject);
  type
      TDadosRT = record
        idProgramacaoExecutante, idProgramacaoRT, TipoCusto: Integer;
        DataProgramacao,HoraCobertura: TDateTime;
        RTNumero, TipoRT,DescricaoRT,Requisitante,PessoaContato,TelContato,CentroPlan,GrpPlan,
        MatriculaPax, CPF, Passaporte, Modal,Classe,
        Origem, Destino, Retorno, DataIda, DataVolta, HoraIda, HoraVolta,
        CentroCusto, DiagramaRede, OperRede, ElementoPEP,
        ChavePassageiro, ChaveIda, ChaveVolta, ChaveCompleta: string;
        RTCobertura, booRTCriada, booleanRecolhimento, booleanNaoCriarRT: Boolean;
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
  type
    THorario = record
      HoraIda,HoraVolta: String;
      DataIda,DataVolta: TDateTime;
  end;

    type
    TRTSapImportDados = record
      DataImportacao: TDateTime;
      OrigemConsulta: string;      // PONTUAL / PERIODO
      PeriodoInicio: TDateTime;
      PeriodoFim: TDateTime;

      QMNUM: string;
      QMART: string;
      IWERK: string;
      INGRP: string;

      Origem: string;
      Destino: string;
      DataViagem: TDateTime;
      HoraViagem: string;

      PERNR: string;
      TipoDoc: string;
      Documento: string;
      Passageiro: string;
      QMTXT: string;

      RT_Modal: string;
      RT_Classe: string;

      StatusItem: string;
      StatusDescricao: string;

      idProgramacaoRT: Integer;
      idProgramacaoExecutante: Integer;
      idRTSapImport: Integer;
      ImportadoColibri: Boolean;
      Observacao: string;
    end;

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

  type
  TLegRT = record
    IDProgExec: Integer;
    DataProgramacao: TDateTime;
    NomeExecutante: string;
    CodigoSAP: string;
    Origem: string;
    Destino: string;
    DistKm: Double;
    TravelMin: Integer;
    HasDist: Boolean;
    HoraIda: TDateTime;
    HoraChegada: TDateTime;
  end;

  TLegRTList = class(TList<TLegRT>)
  end;

  type
  TLegTransbordo = record
    IdProgramacaoExecutante: Integer;
    DataProgramacao: TDateTime;
    Origem: string;
    Destino: string;
    HoraIda: string;
    HoraVolta: string;
    DataVolta: TDateTime;
    TemDataVolta: Boolean;
  end;

  TLegTransbordoList = class(TList<TLegTransbordo>)
  end;

  private
    { Private declarations }
    enderecoColibriRegistro: String;
    function RegistroDentroDoPeriodo(const DS: TDataSet; const ADataIni,
      ADataFim: TDateTime): Boolean;
    procedure DebugCamposDataset(DS: TDataSet);
    function RateioValido(const ACentroCusto, ADiagramaRede, AOperRede,
      AElementoPEP: string): Boolean;
    function ObterInfoTblPlataforma(const APlataforma: string;
      out ABooleanProntidao: Boolean; out ACentroCusto, ADiagramaRede,
      AOperRede, AElementoPEP: string): Boolean;
    function ObterRateioExecutanteComFallback(
      const AIdProgramacaoExecutante: Integer; const ACodigoSAPAtual,
      ANomeExecutante, AEmpresa: string; out ACodigoSAPUsado, ACentroCusto,
      ADiagramaRede, AOperRede, AElementoPEP: string): Boolean;

    procedure CarregarDetalheFrequencia;
    function StatusRTNormalizado(const AStatus: string): string;
    function StatusUnicoPorEventoSAP(const AEvento, AValor,
      ARTNumero: string): string;

    function StatusRTSustentaRT(const AStatus: string): Boolean;
    function StatusImportacaoSAP(const ARTNumero, AStatusItem,
      AStatusDescricao: string): string;
    procedure ReclassificarExecutantesParaR3NoPeriodo(const ADataIni,
      ADataFim: TDateTime; const AForcarTudo: Boolean);
    procedure InicializarSimulacaoLogistica;
    procedure AtualizarDiasDisponiveisSimulacao;
    procedure PreencherCenarioBaseSimulacao;
    procedure DefinirLinhaSimulacao(const ARow: Integer; const ANome, AValor,
      AObservacao: string);
    function ValorSimulacaoTexto(const ARow: Integer): string;
    function TryParseNumeroSimulacao(const ATexto: string;
      out AValor: Double): Boolean;
    function LerInteiroSimulacao(const ARow: Integer;
      const ANome: string): Integer;
    function LerFloatSimulacao(const ARow: Integer;
      const ANome: string): Double;
    function FormatarNumeroSimulacao(const AValor: Double;
      const ACasas: Integer = 1): string;
    function LerParametrosSimulacaoAtual: TSimulacaoParametros;
    procedure CarregarParametrosSimulacaoComBaseReal;
    function ClassificarModalSimulacao(const ANomeEmbarcacao,
      ATipoEmbarcacao: string; const ACapacidade: Integer;
      const AUsaBridgeMesmoGrupo: Boolean): string;
    function EstimarDuracaoRotaModalHoras(const AModal,
      ARotaSequencia: string; const ATempoTransbordo: Integer): Double;
    procedure AplicarMetricasModalNaGrade(const AModal: string;
      const AQtdEmbarcacoes: Integer; const ADisponibilidade,
      AHorasUteis, AProdutividade: Double);
    function DescreverFrotaSimulacao(
      const AParam: TSimulacaoParametros): string;
    function SimularCenarioLogistico(const ANome: string;
      const AParam: TSimulacaoParametros): TSimulacaoCenario;
    procedure MontarCenariosComparativos(
      const ABase: TSimulacaoParametros; var ACenarios: array of TSimulacaoCenario);
    function GerarRelatorioComparativoTexto(
      const ACenarios: array of TSimulacaoCenario): string;
    function GerarRelatorioComparativoCSV(
      const ACenarios: array of TSimulacaoCenario): string;

  const
    SIM_COL_PARAMETRO         = 0;
    SIM_COL_VALOR             = 1;
    SIM_COL_OBSERVACAO        = 2;
    SIM_ROW_QTD_SURFER        = 1;
    SIM_ROW_QTD_SOV           = 2;
    SIM_ROW_QTD_BAE           = 3;
    SIM_ROW_QTD_AQUA          = 4;
    SIM_ROW_BACKLOG           = 5;
    SIM_ROW_DIAS_DISPONIVEIS  = 6;
    SIM_ROW_CUSTO_SURFER      = 7;
    SIM_ROW_CUSTO_SOV         = 8;
    SIM_ROW_CUSTO_BAE         = 9;
    SIM_ROW_CUSTO_AQUA        = 10;
    SIM_ROW_PROD_SURFER       = 11;
    SIM_ROW_PROD_SOV          = 12;
    SIM_ROW_PROD_BAE          = 13;
    SIM_ROW_PROD_AQUA         = 14;
    SIM_ROW_HORAS_SURFER      = 15;
    SIM_ROW_HORAS_SOV         = 16;
    SIM_ROW_HORAS_BAE         = 17;
    SIM_ROW_HORAS_AQUA        = 18;
    SIM_ROW_DISP_SURFER       = 19;
    SIM_ROW_DISP_SOV          = 20;
    SIM_ROW_DISP_BAE          = 21;
    SIM_ROW_DISP_AQUA         = 22;
    SIM_ROW_MOB_SURFER        = 23;
    SIM_ROW_MOB_SOV           = 24;
    SIM_ROW_MOB_BAE           = 25;
    SIM_ROW_MOB_AQUA          = 26;

    RT_STATUS_NAO_CRIAR        = 'NAO_CRIAR';
    RT_STATUS_PENDENTE         = 'PENDENTE';
    RT_STATUS_PRONTO_EMITIR    = 'PRONTO_EMITIR';
    RT_STATUS_EMITIDA          = 'EMITIDA';
    RT_STATUS_JA_EXISTE        = 'JA_EXISTE';
    RT_STATUS_CANCELADA        = 'CANCELADA';
    RT_STATUS_ORFA             = 'ORFA';
    RT_STATUS_ERRO_NAO_ATIVO   = 'ERRO_NAO_ATIVO';
    RT_STATUS_ERRO_CONFLITO_RT = 'ERRO_CONFLITO_RT';
    RT_STATUS_ERRO_EMISSAO     = 'ERRO_EMISSAO';

    MSG_RT_NAO_CRIAR        = 'Registro não requer emissão de RT.';
    MSG_RT_PENDENTE         = 'Registro precisa de revisão.';
    MSG_RT_PRONTO_EMITIR    = 'Registro apto para emissão de RT.';
    MSG_RT_EMITIDA          = 'RT emitida com sucesso no SAP.';
    MSG_RT_JA_EXISTE        = 'Registro já possui RT válida.';
    MSG_RT_CANCELADA        = 'RT cancelada.';
    MSG_RT_ORFA             = 'RT órfã identificada.';
    MSG_RT_ERRO_NAO_ATIVO   = 'Código SAP do passageiro não está ativo.';
    MSG_RT_ERRO_CONFLITO_RT = 'Conflito com outra RT já existente.';
    MSG_RT_ERRO_EMISSAO     = 'Erro ao processar RT no SAP.';

    //Dados da embarcação
    const HoraBase: string = '07:00';
    const VelocidadeKmH: Double = 35;
    const MinTravelMin: Integer = 15;
    const TempoConexaoMin: Integer = 30;
    const DefaultNoDistMin: Integer = 60;
    const TempoPermanenciaMin: Integer = 120; // 2h


    function CampoAsString(DS: TDataSet; const NomeCampo: string): string;
    function PodeGerarRTNoRegistro(DS: TDataSet;
      out MotivoBloqueio: string): Boolean;
    procedure AplicarTransbordoNoPeriodo(const ADataIni, ADataFim: TDateTime);
    function BuscarProgramacaoExecutantePorImportacaoSAP(
      const DadosSAP: TRTSapImportDados;
      out AIdProgramacaoExecutante: Integer): Boolean;
    function CriarRTLocalViaImportacaoSAP(const DadosSAP: TRTSapImportDados;
      const AIdProgramacaoExecutante: Integer): Integer;
    function PlataformaPorNomeSAP(const ANomeSAP: string): string;
    function BuscarProgramacaoExecutanteParaRecuperacao(
      const DadosSAP: TRTSapImportDados): Integer;
    procedure ComplementarProgramacaoExecutanteViaSAPImport(
      const AIdProgramacaoExecutante: Integer;
      const DadosSAP: TRTSapImportDados);
    procedure ComplementarProgramacaoRTViaSAPImport(
      const AIdProgramacaoRT: Integer; const DadosSAP: TRTSapImportDados);
    procedure RecuperarRetornoRTsViaImportacaoSAP(const ADataIni,
      ADataFim: TDateTime);

    procedure AddKeySet(ALst: TStringList; const AValor: string);
    function HasKeySet(ALst: TStringList; const AValor: string): Boolean;

    procedure AtualizarMarcacaoRTOrfa(const AIdProgramacaoRT: Integer;
      const AOrfa, APendenteCancelamento: Boolean; const AMotivo: string);
    procedure MontarBaseSustentacaoRTAtual(const ADataIni, ADataFim: TDateTime;
      AIdsValidos, ARTsValidas, AChavesCompletas, AChavesIda,
      AChavesVolta: TStringList);
    function MontarDadosRTDoExecutanteAtual(DS: TDataSet;
      out Dados: TDadosRT): Boolean;
    procedure AtualizarMarcacaoImportacaoRTOrfa(const AIdRTSapImport: Integer;
      const AOrfa, APendenteCancelamento: Boolean; const AMotivo: string);
    function CriarEspelhoLocalRTOrfaViaImportacaoSAP(
      const DadosSAP: TRTSapImportDados): Integer;
    procedure ConciliarRTsOrfasNoPeriodo(const ADataIni, ADataFim: TDateTime);
    function TryParseDataChave(const S: string; out AData: TDateTime): Boolean;
    function NormalizarDataChave(const S: string): string; overload;
    function NormalizarDataChave(const AData: TDateTime): string; overload;
    function AjustarTextoParaCampo(AConn: TADOConnection; const ATabela, ACampo,
      AValor: string; const APadrao: Integer = 255): string;
    function TamanhoCampoTexto(AConn: TADOConnection; const ATabela, ACampo: string;
      const APadrao: Integer = 255): Integer;
    procedure CarregarRegrasRecolhimentoRT(ALista: TList<TRegraRecolhimentoRT>);
    function RegraRecolhimentoConfere(const ARegra: TRegraRecolhimentoRT;
      const AOrigem, ADestino: string;
      const ADestinoPlataforma: TDadosPlataforma): Boolean;
    function ResolverRecolherParaRegra(const ARegra: TRegraRecolhimentoRT;
      const AOrigem, ADestino: string): string;
    procedure AtualizarRecolhimentoExecutante(
      const AIdProgramacaoExecutante: Integer;
      const ARecolhe: Boolean;
      const ARecolherPara, AHoraIda, AHoraVolta: string;
      const ADataVolta: TDateTime;
      const ATemDataVolta: Boolean);
    procedure AplicarRegrasRecolhimentoBanco(
      const ADataIni, ADataFim: TDateTime;
      const AForcarTudo: Boolean = False);
    procedure GarantirTabelaRTRegraRecolhimento;
    procedure RecalcularClassesRTNoPeriodo(
      const ADataIni, ADataFim: TDateTime;
      const AForcarTudo: Boolean = False);
    function ExecutanteSustentaRT(DS: TDataSet): Boolean;
    procedure AtualizarVinculoProgramacaoRTExecutante(const AIdProgramacaoRT,
      AIdProgramacaoExecutante: Integer);
    function BuscarProgramacaoExecutanteParaSincronizarRT(
      DSRT: TDataSet): Integer;
    function ExecutanteAceitaRT(DS: TDataSet): Boolean;
    function ExecutanteIdAceitaRT(
      const AIdProgramacaoExecutante: Integer): Boolean;
    procedure LimparProgramacaoExecutanteRT(
      const AIdProgramacaoExecutante: Integer; const AMotivo: string);
    function CodigoSAPEhNumerico(const CodigoSAP: string): Boolean;
    function BuscarNovoCodigoSAPExecutante(const ANomeExecutante,
      AEmpresa: string; out ACodigoSAP: string): Boolean;
    procedure AtualizarCodigoSAPProgramacaoExecutante(
      const AIdProgramacaoExecutante: Integer;
      const ACodigoSAPNovo: string);
    function TentarReclassificarExecutanteParaR3(
      const AIdProgramacaoExecutante: Integer; const ANomeExecutante, AEmpresa,
      ACodigoSAPAtual: string; out ACodigoSAPFinal,
      ATipoRTFinal: string): Boolean;
    procedure AtualizarCodigoSAPTblExecutante(const ANomeExecutante, AEmpresa,
      ACodigoSAPNovo: string);
    function NormalizarCodigoSAP(const ACodigo: string): string;
    procedure PersistirCodigoSAPNormalizado(
      const AIdProgramacaoExecutante: Integer; const ANomeExecutante, AEmpresa,
      ACodigoSAPOriginal: string; out ACodigoSAPNormalizado: string);
    function PrecisaPrepararRegistro(
      DS: TDataSet; const AForcarTudo: Boolean): Boolean;
    function PrecisaReprocessarRecolhimentoOuClasse(DS: TDataSet;
      const AForcarTudo: Boolean): Boolean;

    function NormalizarMensagemLog(const AMensagem: string): string;
    function RTLocalIndicaCancelamentoReal(const AStatus: string; const ARTCancelada: Boolean): Boolean;


    procedure AtualizarMensagemRetornoSemRebaixarStatus(const AIdExec,
      AIdRT: Integer; const AMensagem: string);



    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE); message WM_MDIACTIVATE;
    procedure ProcessarRetornoSAP(const LinhaLog: string);

    function Existe_RT(const Dados: TDadosRT): TRT;
    function CriarRegistroTabelaRT(Dados: TDadosRT): Integer;
    function CustoExecutante(const CodigoSAP, CPF, Passaporte: String): TCodicoCusto;
    procedure GerarMultiplasRTsArray;

    procedure AtualizarProgramacaoRT_Cancelamento(
      const idProgramacaoRT: Integer; const Cancelada: Boolean; const Mensagem,
      RT_Status: string);
    procedure ProcessarRetornoCancelamentoSAP(const LinhaLog: string);

    function HoraIdaHoraVolta(booleanTurno2: Boolean;
      DataIda: TDateTime): THorario;
    function DataSAP(const AData: TDateTime): string;

    procedure AtualizarProgramacaoExecutante(
      const idProgramacaoExecutante: Integer;
      const MensagemRT, RT_Numero, RT_Status, RT_HoraIda, RT_HoraVolta, RT_Modal,
      RT_Classe, RecolherPara: string; booleanRecolhimento: Boolean;
      DataVolta: TDateTime);

    function DetermineClasse(const ATipoRT, AOrigem, ADestino,
      ListaOrigens, Classe: string; Recolhimento: Boolean): string;
    procedure AtualizarRetornoExecutante(  const AIdExec: Integer;
      const AMensagem, ANumeroRT, AStatus: string);
    procedure AtualizarRetornoTabelaRT(const AIdRT: Integer; const AMensagem,
      ANumeroRT, AStatus: string);

    //---------------------------------------------------
    function NormalizarTextoChave(const S: string): string;
    function NormalizarHoraChave(const S: string): string;
    function NormalizarDocumentoChave(const S: string): string;

    function ChavePassageiro(
      const CodigoSAP, TipoDoc, Documento: string): string;

    function ChaveRTIda(const Dados: TDadosRT): string;
    function ChaveRTVolta(const Dados: TDadosRT): string;
    function ChaveRTCompleta(const Dados: TDadosRT): string;

    //---------------------------------------------------
    function ChavePassageiroSAPImport(
      const PERNR, TipoDoc, Documento: string): string;
    function ChaveRTIdaSAPImport(
      const DadosSAP: TRTSapImportDados): string;
    procedure InserirOuAtualizarRTSapImport(
      const DadosSAP: TRTSapImportDados);

    //--------------------------------------------------
    function TryParseDataSAPBr(const S: string; out AData: TDateTime): Boolean;
    function StripPrefixoLog(const LinhaLog: string): string;

    //--------------------------------------------------
    procedure ProcessarRetornoImportacaoRTSAP(
      const LinhaLog, AOrigemConsulta: string;
      const APeriodoInicio, APeriodoFim: TDateTime);

    //--------------------------------------------------
    procedure ImportarRTsSAPPeriodoPorTipo(
      const ATipoRT: string;
      const APeriodoInicio, APeriodoFim: TDateTime);

    //--------------------------------------------------
    procedure ImportarRTsSAPPeriodo;
    function DescricaoStatusRTSAP(const AStatus: string): string;
    procedure LimparRTSapImportPeriodo(const ADataIni, ADataFim: TDateTime);
    procedure ReimportarRTsSAPPeriodo;
    procedure AtualizarProgramacaoExecutanteViaImportacaoSAP(
      const AIdProgramacaoExecutante: Integer; const ARTNumero, AStatusItem,
      AStatusDescricao, AOrigemVinculo: string);
    procedure AtualizarProgramacaoRTViaImportacaoSAP(
      const AIdProgramacaoRT: Integer; const ARTNumero, AStatusItem,
      AStatusDescricao, AOrigemVinculo: string);
    procedure AtualizarVinculoRTSapImport(const AIdRTSapImport,
      AIdProgramacaoRT, AIdProgramacaoExecutante: Integer;
      const AImportadoColibri: Boolean; const AObservacao: string);
    function BuscarProgramacaoRTPorChave(const ACampoChave, AValorChave: string;
      out AIdProgramacaoRT, AIdProgramacaoExecutante: Integer): Integer;
    function BuscarProgramacaoRTPorNumeroRT(const ART: string;
      out AIdProgramacaoRT, AIdProgramacaoExecutante: Integer): Integer;
    procedure VincularRTSapImportComProgramacao(const APeriodoInicio,
      APeriodoFim: TDateTime);
    function BuscarRTLocalPorChave(
      const Dados: TDadosRT;
      out AIdRT: Integer;
      out ART, AMensagem, AStatus: string): Boolean;
    function BuscarRTSapImportPorChave(const Dados: TDadosRT;
      out ADadosSAP: TRTSapImportDados): Boolean;
    procedure PrepararProgramacaoExecutanteParaRT(
      const ADataIni, ADataFim: TDateTime;
      const AForcarTudo: Boolean = False);
    procedure SincronizarProgramacaoExecutanteComRTExistente(const ADataIni, ADataFim: TDateTime);
    procedure AplicarRegrasAutomaticasRT(
      const ADataIni, ADataFim: TDateTime;
      const AForcarTudo: Boolean = False);
    //-------------------------------------------------
    function AvaliarNecessidadeCriacaoRT(
        const DS: TDataSet;
        out Status, Mensagem: string): Boolean; overload;

    procedure AvaliarNecessidadeCriacaoRT(
      const ADataIni, ADataFim: TDateTime;
      const AForcarTudo: Boolean = False); overload;
    //-------------------------------------------------
    function DeterminarModalAutomatico(const AOrigem, ADestino: string): string;
    function DeterminarTipoRTAutomatico(const CodigoSAP: string): string;
    function DeveCriarRTAutomaticamente(const AOrigem, ADestino: string;
      const PlataformaOrigem,
      PlataformaDestino: TDadosPlataforma): Boolean;
    procedure AplicarRecolhimentoNoRegistro(DS: TDataSet;
      const Horario: THorario; const AOrigem: string;
      const ABooleanRecolhimento: Boolean);
    procedure PreencherCamposAutomaticosRTNoRegistro(DS: TDataSet; const TipoRT,
      Modal, Classe: string; const Horario: THorario);

  public
    ExcelRT,SheetRT: Variant;
    { Public declarations }
  end;

var
  FrmConsultaExecutantesProgramados: TFrmConsultaExecutantesProgramados;

implementation
  uses untPrincipal,untDataModule,untFrmPreview, untFrmConfigRT, untFrmTabela,
  untMotivoCancelamento;
{$R *.dfm}

type
  TInfoEmbarcacaoSimulacao = record
    TipoEmbarcacao: string;
    UsaBridgeMesmoGrupo: Boolean;
  end;

  TMetricasModalReal = class
  public
    TotalMovimentacoes: Integer;
    TotalHorasRota: Double;
    TotalCapacidadeOferta: Double;
    Embarcacoes: TStringList;
    DiasAtivos: TStringList;
    EmbarcacaoDiasAtivos: TStringList;
    RotasProcessadas: TStringList;
    constructor Create;
    destructor Destroy; override;
    function OcupacaoMediaPercentual: Double;
  end;

constructor TMetricasModalReal.Create;
begin
  inherited Create;
  Embarcacoes := TStringList.Create;
  Embarcacoes.Sorted := True;
  Embarcacoes.Duplicates := dupIgnore;
  DiasAtivos := TStringList.Create;
  DiasAtivos.Sorted := True;
  DiasAtivos.Duplicates := dupIgnore;
  EmbarcacaoDiasAtivos := TStringList.Create;
  EmbarcacaoDiasAtivos.Sorted := True;
  EmbarcacaoDiasAtivos.Duplicates := dupIgnore;
  RotasProcessadas := TStringList.Create;
  RotasProcessadas.Sorted := True;
  RotasProcessadas.Duplicates := dupIgnore;
end;

destructor TMetricasModalReal.Destroy;
begin
  RotasProcessadas.Free;
  EmbarcacaoDiasAtivos.Free;
  DiasAtivos.Free;
  Embarcacoes.Free;
  inherited;
end;

function TMetricasModalReal.OcupacaoMediaPercentual: Double;
begin
  if TotalCapacidadeOferta > 0 then
    Result := (TotalMovimentacoes / TotalCapacidadeOferta) * 100
  else
    Result := 0;
end;

function TFrmConsultaExecutantesProgramados.StatusRTNormalizado(
  const AStatus: string): string;
var
  S: string;
begin
  S := UpperCase(Trim(AStatus));

  if S = RT_STATUS_NAO_CRIAR then Exit(RT_STATUS_NAO_CRIAR);
  if S = RT_STATUS_PENDENTE then Exit(RT_STATUS_PENDENTE);
  if S = RT_STATUS_PRONTO_EMITIR then Exit(RT_STATUS_PRONTO_EMITIR);
  if S = RT_STATUS_EMITIDA then Exit(RT_STATUS_EMITIDA);
  if S = RT_STATUS_JA_EXISTE then Exit(RT_STATUS_JA_EXISTE);
  if S = RT_STATUS_CANCELADA then Exit(RT_STATUS_CANCELADA);
  if S = RT_STATUS_ORFA then Exit(RT_STATUS_ORFA);
  if S = RT_STATUS_ERRO_NAO_ATIVO then Exit(RT_STATUS_ERRO_NAO_ATIVO);
  if S = RT_STATUS_ERRO_CONFLITO_RT then Exit(RT_STATUS_ERRO_CONFLITO_RT);
  if S = RT_STATUS_ERRO_EMISSAO then Exit(RT_STATUS_ERRO_EMISSAO);

  Result := '';
end;

procedure TFrmConsultaExecutantesProgramados.CarregarRegrasRecolhimentoRT(
  ALista: TList<TRegraRecolhimentoRT>);
var
  Q: TADOQuery;
  R: TRegraRecolhimentoRT;
begin
  if ALista = nil then
    Exit;

  ALista.Clear;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := False;
    Q.SQL.Text :=
      'SELECT tblRTRegraRecolhimento.* '+
      'FROM tblRTRegraRecolhimento '+
      'WHERE Ativa = True '+
      'ORDER BY Prioridade, idRegraRecolhimento';
    Q.Open;

    while not Q.Eof do
    begin
      R := Default(TRegraRecolhimentoRT);

      R.ID := Q.FieldByName('idRegraRecolhimento').AsInteger;
      R.Ativa := Q.FieldByName('Ativa').AsBoolean;
      R.Prioridade := Q.FieldByName('Prioridade').AsInteger;
      R.Descricao := Trim(Q.FieldByName('Descricao').AsString);
      R.Origem := Trim(Q.FieldByName('Origem').AsString);
      R.Destino := Trim(Q.FieldByName('txtDestino').AsString);
      R.DestinoNaoCriarRT := UpperCase(Trim(Q.FieldByName('DestinoNaoCriarRT').AsString));
      R.RecolherParaTipo := UpperCase(Trim(Q.FieldByName('RecolherParaTipo').AsString));
      R.RecolherParaValor := Trim(Q.FieldByName('RecolherParaValor').AsString);
      R.Observacao := Trim(Q.FieldByName('Observacao').AsString);

      ALista.Add(R);
      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.RegraRecolhimentoConfere(
  const ARegra: TRegraRecolhimentoRT;
  const AOrigem, ADestino: string;
  const ADestinoPlataforma: TDadosPlataforma
): Boolean;
begin
  Result := False;

  if not ARegra.Ativa then
    Exit;

  if (Trim(ARegra.Origem) <> '') and
     (not SameText(Trim(ARegra.Origem), Trim(AOrigem))) then
    Exit;

  if (Trim(ARegra.Destino) <> '') and
     (not SameText(Trim(ARegra.Destino), Trim(ADestino))) then
    Exit;

  if ARegra.DestinoNaoCriarRT = 'SIM' then
  begin
    if not ADestinoPlataforma.booleanNaoCriarRT then
      Exit;
  end
  else if ARegra.DestinoNaoCriarRT = 'NAO' then
  begin
    if ADestinoPlataforma.booleanNaoCriarRT then
      Exit;
  end;

  Result := True;
end;

function TFrmConsultaExecutantesProgramados.ResolverRecolherParaRegra(
  const ARegra: TRegraRecolhimentoRT;
  const AOrigem, ADestino: string
): string;
begin
  Result := '';

  if ARegra.RecolherParaTipo = 'ORIGEM' then
    Exit(Trim(AOrigem));

  if ARegra.RecolherParaTipo = 'DESTINO' then
    Exit(Trim(ADestino));

  if ARegra.RecolherParaTipo = 'FIXO' then
    Exit(Trim(ARegra.RecolherParaValor));
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarRecolhimentoExecutante(
  const AIdProgramacaoExecutante: Integer;
  const ARecolhe: Boolean;
  const ARecolherPara, AHoraIda, AHoraVolta: string;
  const ADataVolta: TDateTime;
  const ATemDataVolta: Boolean);
var
  Q: TADOQuery;
begin
  if AIdProgramacaoExecutante <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;

    if ARecolhe then
    begin
      if ATemDataVolta then
      begin
        Q.SQL.Text :=
          'UPDATE tblProgramacaoExecutante SET '+
          '  booleanRecolhimento = True, '+
          '  RecolherPara = :RecolherPara, '+
          '  RT_HoraIda = IIf((RT_HoraIda IS NULL OR Trim(RT_HoraIda) = ''''), :RT_HoraIda, RT_HoraIda), '+
          '  RT_HoraVolta = :RT_HoraVolta, '+
          '  DataVolta = :DataVolta '+
          'WHERE idProgramacaoExecutante = :ID '+
          '  AND (RT_Transbordo = False OR RT_Transbordo IS NULL)';

        Q.Parameters.ParamByName('RecolherPara').Value := Trim(ARecolherPara);
        Q.Parameters.ParamByName('RT_HoraIda').Value := Trim(AHoraIda);
        Q.Parameters.ParamByName('RT_HoraVolta').Value := Trim(AHoraVolta);
        Q.Parameters.ParamByName('DataVolta').Value := ADataVolta;
      end
      else
      begin
        Q.SQL.Text :=
          'UPDATE tblProgramacaoExecutante SET '+
          '  booleanRecolhimento = True, '+
          '  RecolherPara = :RecolherPara, '+
          '  RT_HoraIda = IIf((RT_HoraIda IS NULL OR Trim(RT_HoraIda) = ''''), :RT_HoraIda, RT_HoraIda), '+
          '  RT_HoraVolta = :RT_HoraVolta, '+
          '  DataVolta = Null '+
          'WHERE idProgramacaoExecutante = :ID '+
          '  AND (RT_Transbordo = False OR RT_Transbordo IS NULL)';

        Q.Parameters.ParamByName('RecolherPara').Value := Trim(ARecolherPara);
        Q.Parameters.ParamByName('RT_HoraVolta').Value := Trim(AHoraVolta);
      end;
    end
    else
    begin
      Q.SQL.Text :=
        'UPDATE tblProgramacaoExecutante SET '+
        '  booleanRecolhimento = False, '+
        '  RecolherPara = Null, '+
        '  RT_HoraVolta = Null, '+
        '  DataVolta = Null '+
        'WHERE idProgramacaoExecutante = :ID '+
        '  AND (RT_Classe IS NULL OR RT_Classe <> ''TR'')';
    end;

    Q.Parameters.ParamByName('ID').Value := AIdProgramacaoExecutante;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.CarregarDetalheFrequencia;
begin
  if FrmDataModule.ADOQueryFrequenciaResumo.IsEmpty then
    Exit;

  FrmDataModule.ADOQueryFrequenciaDetalhe.Close;
  FrmDataModule.ADOQueryFrequenciaDetalhe.Parameters.ParamByName('pCodigoSAP').Value :=
    FrmDataModule.ADOQueryFrequenciaResumo.FieldByName('CodigoSAP').AsString;
  FrmDataModule.ADOQueryFrequenciaDetalhe.Open;
end;

procedure TFrmConsultaExecutantesProgramados.AplicarRegrasRecolhimentoBanco(
  const ADataIni, ADataFim: TDateTime;
  const AForcarTudo: Boolean = False);
var
  Q: TADOQuery;
  Regras: TList<TRegraRecolhimentoRT>;
  Regra: TRegraRecolhimentoRT;
  DtIni, DtFimMais1: TDateTime;
  DataProgramacao: TDateTime;
  Origem, Destino, RecolherPara: string;
  IdExec: Integer;
  PlataformaDestino: TDadosPlataforma;
  Horario: THorario;
  AchouRegra: Boolean;
  I: Integer;
begin
  Regras := TList<TRegraRecolhimentoRT>.Create;
  Q := TADOQuery.Create(nil);
  try
    CarregarRegrasRecolhimentoRT(Regras);

    DtIni := Trunc(ADataIni);
    DtFimMais1 := Trunc(ADataFim) + 1;

    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'SELECT pd.DataProgramacao, pd.txtDestino, pe.* '+
      'FROM tblProgramacaoDiaria pd '+
      'INNER JOIN tblProgramacaoExecutante pe '+
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria '+
      'WHERE pd.DataProgramacao >= :DtIni '+
      '  AND pd.DataProgramacao < :DtFimMais1 '+
      '  AND (pe.RT_Transbordo = False OR pe.RT_Transbordo IS NULL) '+
      'ORDER BY pd.DataProgramacao, pe.NomeExecutante, pe.idProgramacaoExecutante';

    Q.Parameters.ParamByName('DtIni').Value := DtIni;
    Q.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    Q.Open;

    FrmPrincipal.ProgressBarIncializa(
      Q.RecordCount,
      'Aplicando regras de recolhimento do banco...'
    );

    while not Q.Eof do
    begin
      if not PrecisaReprocessarRecolhimentoOuClasse(Q, AForcarTudo) then
      begin
        Q.Next;
        FrmPrincipal.ProgressBarIncremento(1);
        Continue;
      end;

      IdExec := Q.FieldByName('idProgramacaoExecutante').AsInteger;
      DataProgramacao := Q.FieldByName('DataProgramacao').AsDateTime;
      Origem := Trim(Q.FieldByName('Origem').AsString);
      Destino := Trim(Q.FieldByName('txtDestino').AsString);

      PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(Destino);

      AchouRegra := False;
      RecolherPara := '';

      for I := 0 to Regras.Count - 1 do
      begin
        Regra := Regras[I];

        if RegraRecolhimentoConfere(Regra, Origem, Destino, PlataformaDestino) then
        begin
          AchouRegra := True;
          RecolherPara := ResolverRecolherParaRegra(Regra, Origem, Destino);

          if Trim(RecolherPara) = '' then
            RecolherPara := Origem;

          Horario := HoraIdaHoraVolta(
            PlataformaDestino.booleanTurno2,
            DataProgramacao
          );

          AtualizarRecolhimentoExecutante(
            IdExec,
            True,
            RecolherPara,
            Horario.HoraIda,
            Horario.HoraVolta,
            Horario.DataVolta,
            True
          );

          Break;
        end;
      end;

      if not AchouRegra then
      begin
        AtualizarRecolhimentoExecutante(
          IdExec,
          False,
          '',
          '',
          '',
          0,
          False
        );
      end;

      Q.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

    FrmPrincipal.ProgressBarAtualizar;
  finally
    Q.Free;
    Regras.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.MontarDadosRTDoExecutanteAtual(
  DS: TDataSet; out Dados: TDadosRT): Boolean;
var
  PlataformaOrigem, PlataformaDestino, PlataformaRetorno: TDadosPlataforma;
  Horario: THorario;
  AOrigem, ADestino, RecolherParaLocal: string;
begin
  Result := False;
  Dados := Default(TDadosRT);

  if (DS = nil) or (not DS.Active) then
    Exit;

  if not FrmDataModule.ADOQueryConfigRT.Active then
    FrmDataModule.ADOQueryConfigRT.Active := True;

  AOrigem := Trim(DS.FieldByName('Origem').AsString);
  ADestino := Trim(DS.FieldByName('txtDestino').AsString);

  PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(AOrigem);
  PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(ADestino);

  if not DeveCriarRTAutomaticamente(AOrigem, ADestino, PlataformaOrigem, PlataformaDestino) then
    Exit;

  Dados.idProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger;
  Dados.Origem := Trim(PlataformaOrigem.NomeSAP);
  Dados.Destino := Trim(PlataformaDestino.NomeSAP);

  RecolherParaLocal := Trim(DS.FieldByName('RecolherPara').AsString);
  if RecolherParaLocal = '' then
    RecolherParaLocal := AOrigem;

  PlataformaRetorno := TProgramacaoRTUtils.DadosPlataforma_RT(RecolherParaLocal);
  Dados.Retorno := Trim(PlataformaRetorno.NomeSAP);

  Dados.MatriculaPax := NormalizarCodigoSAP((Trim(DS.FieldByName('CodigoSAP').AsString)));
  Dados.CPF := Trim(DS.FieldByName('Documento').AsString);
  Dados.Passaporte := Trim(DS.FieldByName('OutroDocumento').AsString);

  Dados.TipoRT := Trim(DS.FieldByName('RT_Tipo').AsString);
  if Trim(Dados.TipoRT) = '' then
    Dados.TipoRT := DeterminarTipoRTAutomatico(Dados.MatriculaPax);

  Dados.Modal := Trim(DS.FieldByName('RT_Modal').AsString);
  Dados.Classe := Trim(DS.FieldByName('RT_Classe').AsString);
  Dados.booleanRecolhimento := DS.FieldByName('booleanRecolhimento').AsBoolean;

  Dados.CentroPlan := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_CentroPlan').AsString);
  Dados.GrpPlan := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_GrpPlan').AsString);

  Dados.DataIda := DataSAP(DS.FieldByName('DataProgramacao').AsDateTime);
  Dados.HoraIda := Trim(DS.FieldByName('RT_HoraIda').AsString);
  Dados.HoraVolta := Trim(DS.FieldByName('RT_HoraVolta').AsString);

  if not VarIsNull(DS.FieldByName('DataVolta').Value) then
    Dados.DataVolta := DataSAP(DS.FieldByName('DataVolta').AsDateTime)
  else if Dados.booleanRecolhimento then
  begin
    Horario := HoraIdaHoraVolta(
      PlataformaDestino.booleanTurno2,
      DS.FieldByName('DataProgramacao').AsDateTime
    );
    Dados.DataVolta := DataSAP(Horario.DataVolta);
  end
  else
    Dados.DataVolta := '';

  Dados.ChavePassageiro := ChavePassageiro(
    Dados.MatriculaPax,
    IfThen(Trim(Dados.CPF) <> '', 'C',
      IfThen(Trim(Dados.Passaporte) <> '', 'P', '')),
    IfThen(Trim(Dados.CPF) <> '', Dados.CPF, Dados.Passaporte)
  );

  if Trim(Dados.TipoRT) = '' then
    Exit;

  if Trim(Dados.Modal) = '' then
    Exit;

  if Trim(Dados.Classe) = '' then
    Exit;

  if Trim(Dados.HoraIda) = '' then
    Exit;

  if Trim(Dados.CentroPlan) = '' then
    Exit;

  if Trim(Dados.GrpPlan) = '' then
    Exit;

  if Trim(Dados.ChavePassageiro) = '' then
    Exit;

  Dados.ChaveIda := ChaveRTIda(Dados);
  Dados.ChaveVolta := ChaveRTVolta(Dados);
  Dados.ChaveCompleta := ChaveRTCompleta(Dados);

  Result := Trim(Dados.ChaveIda) <> '';
end;

procedure TFrmConsultaExecutantesProgramados.MontarBaseSustentacaoRTAtual(
  const ADataIni, ADataFim: TDateTime;
  AIdsValidos, ARTsValidas, AChavesCompletas, AChavesIda, AChavesVolta: TStringList);
var
  Q: TADOQuery;
  DtIni, DtFimMais1: TDateTime;
  Dados: TDadosRT;
begin
  DtIni := Trunc(ADataIni);
  DtFimMais1 := Trunc(ADataFim) + 1;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'SELECT pd.DataProgramacao, pd.txtDestino, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      'ORDER BY pd.DataProgramacao, pd.txtDestino, pe.Origem, pe.idProgramacaoExecutante';

    Q.Parameters.ParamByName('DtIni').Value := DtIni;
    Q.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    Q.Open;

    while not Q.Eof do
    begin
      if ExecutanteSustentaRT(Q) then
      begin
        AddKeySet(AIdsValidos, IntToStr(Q.FieldByName('idProgramacaoExecutante').AsInteger));

        if Trim(Q.FieldByName('RT').AsString) <> '' then
          AddKeySet(ARTsValidas, Q.FieldByName('RT').AsString);

        if MontarDadosRTDoExecutanteAtual(Q, Dados) then
        begin
          AddKeySet(AChavesCompletas, Dados.ChaveCompleta);
          AddKeySet(AChavesIda, Dados.ChaveIda);
          AddKeySet(AChavesVolta, Dados.ChaveVolta);
        end;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarMarcacaoRTOrfa(
  const AIdProgramacaoRT: Integer;
  const AOrfa, APendenteCancelamento: Boolean; const AMotivo: string);
var
  Q: TADOQuery;
  StatusAtual, StatusNovo, MotivoSafe, StatusSafe: string;
  RTCanceladaAtual: Boolean;
begin
  if AIdProgramacaoRT <= 0 then
    Exit;

  StatusAtual := '';
  RTCanceladaAtual := False;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'SELECT RT_Status, RT_Cancelada ' +
      'FROM tblProgramacaoRT ' +
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.Parameters.ParamByName('idProgramacaoRT').Value := AIdProgramacaoRT;
    Q.Open;

    if not Q.Eof then
    begin
      StatusAtual := Trim(Q.FieldByName('RT_Status').AsString);
      RTCanceladaAtual := (Q.FindField('RT_Cancelada') <> nil) and
                          Q.FieldByName('RT_Cancelada').AsBoolean;
    end;
  finally
    Q.Free;
  end;

  if RTCanceladaAtual or SameText(StatusAtual, RT_STATUS_CANCELADA) then
    StatusNovo := RT_STATUS_CANCELADA
  else if APendenteCancelamento then
    StatusNovo := RT_STATUS_ORFA
  else
    StatusNovo := RT_STATUS_EMITIDA;

  MotivoSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT_MotivoConciliacao',
    AMotivo,
    255
  );

  StatusSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT_Status',
    StatusNovo,
    40
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT SET ' +
      '  RT_Orfa = :RT_Orfa, ' +
      '  RT_PendenteCancelamento = :RT_PendenteCancelamento, ' +
      '  RT_MotivoConciliacao = :RT_MotivoConciliacao, ' +
      '  RT_DataUltimaConciliacao = :RT_DataUltimaConciliacao, ' +
      '  RT_Status = :RT_Status ' +
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.Parameters.ParamByName('RT_Orfa').Value := AOrfa;
    Q.Parameters.ParamByName('RT_PendenteCancelamento').Value := APendenteCancelamento;
    Q.Parameters.ParamByName('RT_MotivoConciliacao').Value := MotivoSafe;
    Q.Parameters.ParamByName('RT_DataUltimaConciliacao').Value := Now;
    Q.Parameters.ParamByName('RT_Status').Value := StatusSafe;
    Q.Parameters.ParamByName('idProgramacaoRT').Value := AIdProgramacaoRT;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarMensagemRetornoSemRebaixarStatus(
  const AIdExec, AIdRT: Integer;
  const AMensagem: string);
var
  Q: TADOQuery;
  MsgExecSafe, MsgRTSafe: string;
begin
  MsgExecSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionColibri,
    'tblProgramacaoExecutante',
    'RT_Mensagem',
    Trim(AMensagem),
    255
  );

  MsgRTSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT_Mensagem',
    Trim(AMensagem),
    255
  );

  if AIdExec > 0 then
  begin
    Q := TADOQuery.Create(nil);
    try
      Q.Connection := FrmDataModule.ADOConnectionColibri;
      Q.ParamCheck := True;
      Q.SQL.Text :=
        'UPDATE tblProgramacaoExecutante ' +
        'SET RT_Mensagem = :pMsg ' +
        'WHERE idProgramacaoExecutante = :pId';
      Q.Parameters.ParamByName('pMsg').Value := MsgExecSafe;
      Q.Parameters.ParamByName('pId').Value := AIdExec;
      Q.ExecSQL;
    finally
      Q.Free;
    end;
  end;

  if AIdRT > 0 then
  begin
    Q := TADOQuery.Create(nil);
    try
      Q.Connection := FrmDataModule.ADOConnectionRT;
      Q.ParamCheck := True;
      Q.SQL.Text :=
        'UPDATE tblProgramacaoRT ' +
        'SET RT_Mensagem = :pMsg ' +
        'WHERE idProgramacaoRT = :pId';
      Q.Parameters.ParamByName('pMsg').Value := MsgRTSafe;
      Q.Parameters.ParamByName('pId').Value := AIdRT;
      Q.ExecSQL;
    finally
      Q.Free;
    end;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarMarcacaoImportacaoRTOrfa(
  const AIdRTSapImport: Integer;
  const AOrfa, APendenteCancelamento: Boolean;
  const AMotivo: string);
var
  Q: TADOQuery;
  MotivoSafe: string;
begin
  if AIdRTSapImport <= 0 then
    Exit;

  MotivoSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblRTSapImport',
    'RT_MotivoConciliacao',
    AMotivo,
    255
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'UPDATE tblRTSapImport SET ' +
      '  RT_Orfa = :RT_Orfa, ' +
      '  RT_PendenteCancelamento = :RT_PendenteCancelamento, ' +
      '  RT_MotivoConciliacao = :RT_MotivoConciliacao, ' +
      '  RT_DataUltimaConciliacao = :RT_DataUltimaConciliacao ' +
      'WHERE idRTSapImport = :idRTSapImport';

    Q.Parameters.ParamByName('RT_Orfa').Value := AOrfa;
    Q.Parameters.ParamByName('RT_PendenteCancelamento').Value := APendenteCancelamento;
    Q.Parameters.ParamByName('RT_MotivoConciliacao').Value := MotivoSafe;
    Q.Parameters.ParamByName('RT_DataUltimaConciliacao').Value := Now;
    Q.Parameters.ParamByName('idRTSapImport').Value := AIdRTSapImport;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;


function TFrmConsultaExecutantesProgramados.CriarEspelhoLocalRTOrfaViaImportacaoSAP(
  const DadosSAP: TRTSapImportDados): Integer;
var
  Dados: TDadosRT;
begin
  if not FrmDataModule.ADOQueryConfigRT.Active then
    FrmDataModule.ADOQueryConfigRT.Active := True;

  Dados := Default(TDadosRT);

  Dados.idProgramacaoExecutante := 0;
  Dados.TipoRT := Trim(DadosSAP.QMART);
  if Dados.TipoRT = '' then
    Dados.TipoRT := DeterminarTipoRTAutomatico(DadosSAP.PERNR);

  Dados.DescricaoRT := Copy(
    IfThen(Trim(DadosSAP.QMTXT) <> '',
      Trim(DadosSAP.QMTXT),
      Trim(DadosSAP.Origem) + ': ' + Trim(DadosSAP.Passageiro)),
    1, 40);

  Dados.Requisitante := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_Requisitante').AsString);
  Dados.PessoaContato := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_PessoaContato').AsString);
  Dados.TelContato := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_TelContato').AsString);

  Dados.CentroPlan := Trim(DadosSAP.IWERK);
  if Dados.CentroPlan = '' then
    Dados.CentroPlan := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_CentroPlan').AsString);

  Dados.GrpPlan := Trim(DadosSAP.INGRP);
  if Dados.GrpPlan = '' then
    Dados.GrpPlan := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_GrpPlan').AsString);

  Dados.MatriculaPax := Trim(DadosSAP.PERNR);

  if SameText(Trim(DadosSAP.TipoDoc), 'C') then
    Dados.CPF := Trim(DadosSAP.Documento)
  else
    Dados.Passaporte := Trim(DadosSAP.Documento);

  Dados.Modal := Trim(DadosSAP.RT_Modal);
  Dados.Classe := Trim(DadosSAP.RT_Classe);

  Dados.Origem := Trim(DadosSAP.Origem);
  Dados.Destino := Trim(DadosSAP.Destino);
  Dados.Retorno := '';
  Dados.booleanRecolhimento := False;

  if DadosSAP.DataViagem > 0 then
    Dados.DataIda := DataSAP(DadosSAP.DataViagem)
  else
    Dados.DataIda := '';

  Dados.HoraIda := Trim(DadosSAP.HoraViagem);
  Dados.HoraVolta := '';
  Dados.DataVolta := '';

  Result := CriarRegistroTabelaRT(Dados);

  if Result > 0 then
    AtualizarProgramacaoRTViaImportacaoSAP(
      Result,
      DadosSAP.QMNUM,
      DadosSAP.StatusItem,
      DadosSAP.StatusDescricao,
      'ESPELHO_CANCELAMENTO_IMPORTACAO'
    );
end;

procedure TFrmConsultaExecutantesProgramados.ConciliarRTsOrfasNoPeriodo(
  const ADataIni, ADataFim: TDateTime);
var
  IDsValidos, RTsValidas, ChavesCompletas, ChavesIda, ChavesVolta: TStringList;
  QRT, QImp: TADOQuery;
  DtIni, DtFimMais1: TDateTime;
  IdExec, IdRTLocal: Integer;
  NumeroRT, ChCompleta, ChIda, ChVolta: string;
  Sustenta: Boolean;
  Motivo: string;
  DadosSAP: TRTSapImportDados;
  QtRTLocal, QtRTLocalOrfa, QtImport, QtImportOrfa, QtEspelhosCriados: Integer;

  procedure InitSet(SL: TStringList);
  begin
    SL.Sorted := True;
    SL.Duplicates := dupIgnore;
    SL.CaseSensitive := False;
  end;

begin
  // garante que a programação do período esteja atualizada antes da conciliação
  PrepararProgramacaoExecutanteParaRT(ADataIni, ADataFim);

  IDsValidos := TStringList.Create;
  RTsValidas := TStringList.Create;
  ChavesCompletas := TStringList.Create;
  ChavesIda := TStringList.Create;
  ChavesVolta := TStringList.Create;
  QRT := TADOQuery.Create(nil);
  QImp := TADOQuery.Create(nil);
  try
    InitSet(IDsValidos);
    InitSet(RTsValidas);
    InitSet(ChavesCompletas);
    InitSet(ChavesIda);
    InitSet(ChavesVolta);

    MontarBaseSustentacaoRTAtual(
      ADataIni, ADataFim,
      IDsValidos, RTsValidas, ChavesCompletas, ChavesIda, ChavesVolta
    );

    DtIni := Trunc(ADataIni);
    DtFimMais1 := Trunc(ADataFim) + 1;

    QtRTLocal := 0;
    QtRTLocalOrfa := 0;
    QtImport := 0;
    QtImportOrfa := 0;
    QtEspelhosCriados := 0;

    // ============================================================
    // 1) Avalia tblProgramacaoRT
    // ============================================================
    QRT.Connection := FrmDataModule.ADOConnectionRT;
    QRT.SQL.Text :=
      'SELECT * ' +
      'FROM tblProgramacaoRT ' +
      'WHERE DataIda >= :DtIni ' +
      '  AND DataIda < :DtFimMais1 ' +
      '  AND [RT] IS NOT NULL ' +
      '  AND Trim([RT]) <> '''' ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      'ORDER BY DataIda, RT, idProgramacaoRT';

    QRT.Parameters.ParamByName('DtIni').Value := DtIni;
    QRT.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    QRT.Open;

    while not QRT.Eof do
    begin
      Inc(QtRTLocal);

      IdExec := 0;
      if not VarIsNull(QRT.FieldByName('idProgramacaoExecutante').Value) then
        IdExec := QRT.FieldByName('idProgramacaoExecutante').AsInteger;

      NumeroRT := Trim(QRT.FieldByName('RT').AsString);
      ChCompleta := Trim(QRT.FieldByName('ChaveCompleta').AsString);
      ChIda := Trim(QRT.FieldByName('ChaveIda').AsString);
      ChVolta := Trim(QRT.FieldByName('ChaveVolta').AsString);

      Sustenta :=
        HasKeySet(IDsValidos, IntToStr(IdExec)) or
        HasKeySet(RTsValidas, NumeroRT) or
        HasKeySet(ChavesCompletas, ChCompleta) or
        HasKeySet(ChavesIda, ChIda) or
        HasKeySet(ChavesVolta, ChVolta);

      if Sustenta then
      begin
        AtualizarMarcacaoRTOrfa(
          QRT.FieldByName('idProgramacaoRT').AsInteger,
          False,
          False,
          'RT sustentada pela programação vigente'
        );
      end
      else
      begin
        Motivo := 'RT ativa sem correspondência na programação vigente do período';
        AtualizarMarcacaoRTOrfa(
          QRT.FieldByName('idProgramacaoRT').AsInteger,
          True,
          True,
          Motivo
        );
        Inc(QtRTLocalOrfa);
      end;

      QRT.Next;
    end;

    // ============================================================
    // 2) Avalia tblRTSapImport
    // ============================================================
    QImp.Connection := FrmDataModule.ADOConnectionRT;
    QImp.SQL.Text :=
      'SELECT * ' +
      'FROM tblRTSapImport ' +
      'WHERE DataViagem >= :DtIni ' +
      '  AND DataViagem < :DtFimMais1 ' +
      '  AND [QMNUM] IS NOT NULL ' +
      '  AND Trim([QMNUM]) <> '''' ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      '  AND ([StatusItem] IS NULL OR Trim([StatusItem]) <> ''09'') ' +
      'ORDER BY DataViagem, QMNUM, idRTSapImport';

    QImp.Parameters.ParamByName('DtIni').Value := DtIni;
    QImp.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    QImp.Open;

    while not QImp.Eof do
    begin
      Inc(QtImport);

      FillChar(DadosSAP, SizeOf(DadosSAP), 0);

      DadosSAP.idRTSapImport := QImp.FieldByName('idRTSapImport').AsInteger;
      DadosSAP.QMNUM := Trim(QImp.FieldByName('QMNUM').AsString);
      DadosSAP.QMART := Trim(QImp.FieldByName('QMART').AsString);
      DadosSAP.IWERK := Trim(QImp.FieldByName('IWERK').AsString);
      DadosSAP.INGRP := Trim(QImp.FieldByName('INGRP').AsString);
      DadosSAP.Origem := Trim(QImp.FieldByName('Origem').AsString);
      DadosSAP.Destino := Trim(QImp.FieldByName('txtDestino').AsString);
      if not VarIsNull(QImp.FieldByName('DataViagem').Value) then
        DadosSAP.DataViagem := QImp.FieldByName('DataViagem').AsDateTime;
      DadosSAP.HoraViagem := Trim(QImp.FieldByName('HoraViagem').AsString);
      DadosSAP.PERNR := Trim(QImp.FieldByName('PERNR').AsString);
      DadosSAP.TipoDoc := Trim(QImp.FieldByName('TipoDoc').AsString);
      DadosSAP.Documento := Trim(QImp.FieldByName('Documento').AsString);
      DadosSAP.Passageiro := Trim(QImp.FieldByName('Passageiro').AsString);
      DadosSAP.QMTXT := Trim(QImp.FieldByName('QMTXT').AsString);
      DadosSAP.RT_Modal := Trim(QImp.FieldByName('RT_Modal').AsString);
      DadosSAP.RT_Classe := Trim(QImp.FieldByName('RT_Classe').AsString);
      DadosSAP.StatusItem := Trim(QImp.FieldByName('StatusItem').AsString);
      DadosSAP.StatusDescricao := Trim(QImp.FieldByName('StatusDescricao').AsString);

      IdExec := 0;
      if not VarIsNull(QImp.FieldByName('idProgramacaoExecutante').Value) then
        IdExec := QImp.FieldByName('idProgramacaoExecutante').AsInteger;

      IdRTLocal := 0;
      if not VarIsNull(QImp.FieldByName('idProgramacaoRT').Value) then
        IdRTLocal := QImp.FieldByName('idProgramacaoRT').AsInteger;

      NumeroRT := Trim(QImp.FieldByName('QMNUM').AsString);
      ChCompleta := Trim(QImp.FieldByName('ChaveCompleta').AsString);
      ChIda := Trim(QImp.FieldByName('ChaveIda').AsString);
      ChVolta := Trim(QImp.FieldByName('ChaveVolta').AsString);

      Sustenta :=
        HasKeySet(IDsValidos, IntToStr(IdExec)) or
        HasKeySet(RTsValidas, NumeroRT) or
        HasKeySet(ChavesCompletas, ChCompleta) or
        HasKeySet(ChavesIda, ChIda) or
        HasKeySet(ChavesVolta, ChVolta);

      if Sustenta then
      begin
        AtualizarMarcacaoImportacaoRTOrfa(
          DadosSAP.idRTSapImport,
          False,
          False,
          'RT importada sustentada pela programação vigente'
        );
      end
      else
      begin
        Motivo := 'RT importada ativa sem correspondência na programação vigente do período';

        AtualizarMarcacaoImportacaoRTOrfa(
          DadosSAP.idRTSapImport,
          True,
          True,
          Motivo
        );

        if IdRTLocal <= 0 then
        begin
          IdRTLocal := CriarEspelhoLocalRTOrfaViaImportacaoSAP(DadosSAP);
          if IdRTLocal > 0 then
          begin
            AtualizarVinculoRTSapImport(
              DadosSAP.idRTSapImport,
              IdRTLocal,
              0,
              True,
              'Espelho local criado para cancelamento de RT órfã'
            );

            AtualizarMarcacaoRTOrfa(
              IdRTLocal,
              True,
              True,
              Motivo
            );

            Inc(QtEspelhosCriados);
          end;
        end
        else
        begin
          AtualizarMarcacaoRTOrfa(
            IdRTLocal,
            True,
            True,
            Motivo
          );
        end;

        Inc(QtImportOrfa);
      end;

      QImp.Next;
    end;

    ShowMessage(
      'Conciliação de RTs órfãs concluída.' + sLineBreak +
      'RTs locais avaliadas: ' + IntToStr(QtRTLocal) + sLineBreak +
      'RTs locais órfãs: ' + IntToStr(QtRTLocalOrfa) + sLineBreak +
      'RTs importadas avaliadas: ' + IntToStr(QtImport) + sLineBreak +
      'RTs importadas órfãs: ' + IntToStr(QtImportOrfa) + sLineBreak +
      'Espelhos locais criados: ' + IntToStr(QtEspelhosCriados)
    );

  finally
    IDsValidos.Free;
    RTsValidas.Free;
    ChavesCompletas.Free;
    ChavesIda.Free;
    ChavesVolta.Free;
    QRT.Free;
    QImp.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AplicarTransbordoNoPeriodo(
  const ADataIni, ADataFim: TDateTime);
var
  D: TDateTime;
begin
  D := Trunc(ADataIni);
  while D <= Trunc(ADataFim) do
  begin
    TProgramacaoRTUtils.ProgramacaoTransbordosComHora(
      RLTemporario,
      DateToStr(D),
      True
    );
    D := IncDay(D, 1);
  end;
end;

function TFrmConsultaExecutantesProgramados.TryParseDataSAPBr(
  const S: string; out AData: TDateTime): Boolean;
var
  FS: TFormatSettings;
begin
  AData := 0;
  FS := TFormatSettings.Create;
  FS.DateSeparator := '.';
  FS.ShortDateFormat := 'dd.mm.yyyy';
  Result := TryStrToDate(Trim(S), AData, FS);
end;

function TFrmConsultaExecutantesProgramados.StripPrefixoLog(
  const LinhaLog: string): string;
var
  P: Integer;
begin
  Result := Trim(LinhaLog);

  P := Pos(' | ', Result);
  if P > 0 then
    Exit(Trim(Copy(Result, P + 3, MaxInt)));

  P := Pos('|', Result);
  if P > 0 then
    Exit(Trim(Copy(Result, P + 1, MaxInt)));
end;

procedure TFrmConsultaExecutantesProgramados.actCancelarRTForcadoExecute(Sender: TObject);
var
  EnderecoVbs, EnderecoLog: string;
  DS: TDataSet;
  Bmk: TBookmark;
  LogRetorno: TStringList;
  i: Integer;

  idProgramacaoRT: Integer;
  RTNumero, Origem, Destino, CodigoSAP: string;

  function VbsQuote(const S: string): string;
  begin
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"';
  end;

  procedure MemoAdd(const S: string);
  begin
    MemoSAP.Lines.Add(S);
  end;

begin
  if Application.MessageBox(PChar(
  'Deseja realmente cancelar todas as RTs dos registros filtrados?'),
  '.::ATENÇÃO::.',36) = 6 then
  begin
    EnderecoVbs := ExtractFilePath(ParamStr(0)) + 'RT_Cancelar.vbs';
    EnderecoLog := ExtractFilePath(ParamStr(0)) + 'sap_log_cancel.txt';

    MemoSAP.Lines.Clear;

    DS := FrmDataModule.DataSourceGestaoRT.DataSet;
    if (DS = nil) or (not DS.Active) or DS.IsEmpty then
    begin
      MessageBox(0,PChar('Não há registros para cancelar.'),
          'Colibri',MB_ICONEXCLAMATION);
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

    MemoAdd('Function MsgContratoObrigatorio(ByVal st)');
    MemoAdd('  Dim s');
    MemoAdd('  s = LCase(Trim(CStr(st)))');
    MemoAdd('  MsgContratoObrigatorio = (InStr(1, s, "contrat", vbTextCompare) > 0) And (InStr(1, s, "obrig", vbTextCompare) > 0)');
    MemoAdd('End Function');
    MemoAdd('');

    MemoAdd('Sub CancelarRT(ByVal idProgRT, ByVal rtNumero)');
    MemoAdd('  On Error Resume Next');
    MemoAdd('  Dim st, tp');
    MemoAdd('  LogLine idProgRT & "|CANCELAR|INICIO|RT=" & rtNumero');
    MemoAdd('');
    MemoAdd('  session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"');
    MemoAdd('  session.findById("wnd[0]").sendVKey 0');
    MemoAdd('  WScript.Sleep 100');
    MemoAdd('');
    MemoAdd('  session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero');
    MemoAdd('  session.findById("wnd[0]").sendVKey 0');
    MemoAdd('  WScript.Sleep 100');
    MemoAdd('');
    MemoAdd('  ' + ''' >>> ajuste aqui se seu caminho de menu for diferente <<<');
    MemoAdd('  session.findById("wnd[0]/mbar/menu[0]/menu[10]/menu[11]/menu[0]").select');
    MemoAdd('  WScript.Sleep 100');
    MemoAdd('  session.findById("wnd[0]/tbar[0]/btn[11]").press');
    MemoAdd('  WScript.Sleep 100');
    MemoAdd('');
    MemoAdd('  st = StatusText()');
    MemoAdd('  If Len(Trim(st)) = 0 Then');
    MemoAdd('    WScript.Sleep 300');
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

          //DataIda := DS.FieldByName('DataIda').AsDateTime;
          Origem     := DS.FieldByName('Origem').AsString;
          Destino    := DS.FieldByName('txtDestino').AsString; // atenção: aqui é txtDestino
          CodigoSAP  := RemoverZerosEsquerda(DS.FieldByName('CodigoSAP').AsString);

          // regra 1: se RT está vazia → só marca erro local / loga e pula
          if RTNumero = '' then
          begin
            MemoAdd('LogLine "' + IntToStr(idProgramacaoRT) + '|CANCELAR|ERRO|RT_VAZIA"');
            DS.Next;
            Continue;
          end;

          // cancelamento manual forçado dos registros filtrados
          MemoAdd('CancelarRT ' + IntToStr(idProgramacaoRT) + ', ' + VbsQuote(RTNumero));
          MemoAdd('WScript.Sleep 100');

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
        MessageBox(0,PChar('Cancelamento concluído. Verifique log e tabela.'),
          'Colibri',MB_ICONEXCLAMATION);
      end
      else
        MessageBox(0,PChar('Script terminou, mas não encontrei o log: ' + EnderecoLog),
          'Colibri',MB_ICONEXCLAMATION);
    end
    else
      MessageBox(0,PChar('Erro ao iniciar o script VBS.'),
          'Colibri',MB_ICONERROR);
  end;
  actProcurarProgramacaoRT.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.ConciliarRTsOrfasExecute(
  Sender: TObject);
begin
  if Application.MessageBox(PChar(
  'Deseja rodar a concliliação de RTs orfãs no período?'),
  '.::ATENÇÃO::.',36) = 6 then
    ConciliarRTsOrfasNoPeriodo(dataInicio.Date, dataFim.Date);
end;

procedure TFrmConsultaExecutantesProgramados.actConfigurarRTExecute(
  Sender: TObject);
begin
  if not Assigned(FrmConfigRT) then
    Application.CreateForm(TFrmConfigRT, FrmConfigRT);
  FrmConfigRT.Show;   // ou ShowModal conforme o caso
end;

procedure TFrmConsultaExecutantesProgramados.actProgramacaoRTxProgramacaoExecutanteExecute(Sender: TObject);
var
  Bmk: TBookmark;
  DS: TDataSet;
  HasBookmark: Boolean;
  idProgramacaoExecutante: Integer;
  MensagemRT, RT_Numero, RT_Status,
  RT_HoraIda, RT_HoraVolta, RT_Modal, RT_Classe, RecolherPara: string;
  booleanRecolhimento: Boolean;
  DataVolta: TDateTime;
begin
  DS := FrmDataModule.DataSourceGestaoRT.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
    Exit;

  HasBookmark := False;
  if not DS.IsEmpty then
  begin
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  FrmPrincipal.ProgressBarIncializa(DS.RecordCount, 'Atualizando dados da tblProgramacaoExecutante...');

  try
    DS.DisableControls;
    try
      DS.First;
      while not DS.Eof do
      begin
        idProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger;

        if idProgramacaoExecutante > 0 then
        begin
          MensagemRT := CampoAsString(DS, 'RT_Mensagem');
          RT_Numero := CampoAsString(DS, 'RT');
          RT_Status := CampoAsString(DS, 'RT_Status');
          RT_HoraIda := CampoAsString(DS, 'RT_HoraIda');
          RT_HoraVolta := CampoAsString(DS, 'RT_HoraVolta');
          RT_Modal := CampoAsString(DS, 'RT_Modal');
          RT_Classe := CampoAsString(DS, 'RT_Classe');
          RecolherPara := CampoAsString(DS, 'RecolherPara');
          booleanRecolhimento := (DS.FindField('booleanRecolhimento') <> nil) and
                                 DS.FieldByName('booleanRecolhimento').AsBoolean;

          if (DS.FindField('DataVolta') <> nil) and (not VarIsNull(DS.FieldByName('DataVolta').Value)) then
            DataVolta := DS.FieldByName('DataVolta').AsDateTime
          else
            DataVolta := 0;

          AtualizarProgramacaoExecutante(
            idProgramacaoExecutante,
            MensagemRT,
            RT_Numero,
            RT_Status,
            RT_HoraIda,
            RT_HoraVolta,
            RT_Modal,
            RT_Classe,
            RecolherPara,
            booleanRecolhimento,
            DataVolta
          );
        end;

        DS.Next;
        FrmPrincipal.ProgressBarIncremento(1);
      end;
    finally
      if HasBookmark then
      begin
        try
          if DS.BookmarkValid(Bmk) then
            DS.GotoBookmark(Bmk);
        finally
          DS.FreeBookmark(Bmk);
        end;
      end;

      DS.EnableControls;
      FrmPrincipal.ProgressBarAtualizar;
    end;
  except
    on E: Exception do
      MessageBox(0, PChar('Erro ao atualizar dados da RT na programação: ' + E.Message),
        'Colibri', MB_ICONERROR);
  end;
end;

function TFrmConsultaExecutantesProgramados.HoraIdaHoraVolta(booleanTurno2: Boolean;
  DataIda: TDateTime): THorario;
begin
  if booleanTurno2 then
  begin
    Result.HoraIda:= TRIM(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraInicio_Turno2').AsString);
    Result.HoraVolta:= TRIM(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraVolta_Turno2').AsString);
    Result.DataVolta := DataIda+1;
    Result.DataIda := DataIda;
  end
  else
  begin
    Result.HoraIda:= TRIM(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraInicio_Turno1').AsString);
    Result.HoraVolta:= TRIM(FrmDataModule.ADOQueryConfigRT.FieldByName('HoraVolta_Turno1').AsString);
    Result.DataVolta := DataIda;
    Result.DataIda := DataIda;
  end;
end;

function TFrmConsultaExecutantesProgramados.DetermineClasse(
  const ATipoRT, AOrigem, ADestino, ListaOrigens, Classe: string; Recolhimento: Boolean): string;
var
  ClasseInformada: string;
begin
  ClasseInformada := Trim(Classe);

  // preserva somente classe manual que NÃO seja TR
  // TR será definido exclusivamente pela rotina de transbordo
  if (ClasseInformada <> '') and (not SameText(ClasseInformada, 'TR')) then
    Exit(ClasseInformada);

  if SameText(ATipoRT, 'R7') then
    Result := 'COM'
  else if Recolhimento and SameText(AOrigem, 'TMIB') then
    Result := 'EV'
  else
    Result := 'TT';
end;

procedure TFrmConsultaExecutantesProgramados.RecalcularClassesRTNoPeriodo(
  const ADataIni, ADataFim: TDateTime;
  const AForcarTudo: Boolean = False);
var
  Q: TADOQuery;
  DtIni, DtFimMais1: TDateTime;
  TipoRT, Origem, ClasseNova: string;
  BooleanRecolhimento, EhTransbordo: Boolean;
begin
  DtIni := Trunc(ADataIni);
  DtFimMais1 := Trunc(ADataFim) + 1;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.CursorType := ctStatic;
    Q.LockType := ltOptimistic;

    Q.SQL.Text :=
      'SELECT pd.DataProgramacao, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      'ORDER BY pe.NomeExecutante, pe.idProgramacaoExecutante';

    Q.Parameters.ParamByName('DtIni').Value := DtIni;
    Q.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    Q.Open;

    while not Q.Eof do
    begin
      if not PrecisaReprocessarRecolhimentoOuClasse(Q, AForcarTudo) then
      begin
        Q.Next;
        Continue;
      end;

      TipoRT := Trim(Q.FieldByName('RT_Tipo').AsString);
      Origem := Trim(Q.FieldByName('Origem').AsString);

      if Q.FindField('booleanRecolhimento') <> nil then
        BooleanRecolhimento := Q.FieldByName('booleanRecolhimento').AsBoolean
      else
        BooleanRecolhimento := False;

      if Q.FindField('RT_Transbordo') <> nil then
        EhTransbordo := Q.FieldByName('RT_Transbordo').AsBoolean
      else
        EhTransbordo := False;

      if EhTransbordo then
      begin
        if SameText(TipoRT, 'R7') then
          ClasseNova := 'COM'
        else
          ClasseNova := 'TR';
      end
      else
      begin
        if SameText(TipoRT, 'R7') then
          ClasseNova := 'COM'
        else if BooleanRecolhimento and SameText(Origem, 'TMIB') then
          ClasseNova := 'EV'
        else
          ClasseNova := 'TT';
      end;

      if Trim(Q.FieldByName('RT_Classe').AsString) <> ClasseNova then
      begin
        Q.Edit;
        Q.FieldByName('RT_Classe').AsString := ClasseNova;
        Q.Post;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.TryParseDataChave(
  const S: string; out AData: TDateTime): Boolean;
var
  FS: TFormatSettings;
  V: string;
  Y, M, D: Word;
begin
  Result := False;
  AData := 0;
  V := Trim(S);

  if V = '' then
    Exit;

  if TryStrToDate(V, AData) then
    Exit(True);

  FS := TFormatSettings.Create;
  FS.DateSeparator := '.';
  FS.ShortDateFormat := 'dd.mm.yyyy';
  if TryStrToDate(V, AData, FS) then
    Exit(True);

  FS := TFormatSettings.Create;
  FS.DateSeparator := '/';
  FS.ShortDateFormat := 'dd/mm/yyyy';
  if TryStrToDate(V, AData, FS) then
    Exit(True);

  FS := TFormatSettings.Create;
  FS.DateSeparator := '-';
  FS.ShortDateFormat := 'yyyy-mm-dd';
  if TryStrToDate(V, AData, FS) then
    Exit(True);

  V := StringReplace(V, '.', '', [rfReplaceAll]);
  V := StringReplace(V, '/', '', [rfReplaceAll]);
  V := StringReplace(V, '-', '', [rfReplaceAll]);

  if Length(V) = 8 then
  begin
    // tenta YYYYMMDD
    Y := StrToIntDef(Copy(V, 1, 4), 0);
    M := StrToIntDef(Copy(V, 5, 2), 0);
    D := StrToIntDef(Copy(V, 7, 2), 0);
    if TryEncodeDate(Y, M, D, AData) then
      Exit(True);

    // tenta DDMMYYYY
    D := StrToIntDef(Copy(V, 1, 2), 0);
    M := StrToIntDef(Copy(V, 3, 2), 0);
    Y := StrToIntDef(Copy(V, 5, 4), 0);
    if TryEncodeDate(Y, M, D, AData) then
      Exit(True);
  end;
end;

function TFrmConsultaExecutantesProgramados.TamanhoCampoTexto(
  AConn: TADOConnection; const ATabela, ACampo: string;
  const APadrao: Integer = 255): Integer;
var
  Q: TADOQuery;
begin
  Result := APadrao;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := AConn;
    Q.SQL.Text := 'SELECT TOP 1 [' + ACampo + '] FROM [' + ATabela + ']';
    Q.Open;

    if (Q.FindField(ACampo) <> nil) and
       (Q.FieldByName(ACampo).DataType in [ftString, ftWideString]) and
       (Q.FieldByName(ACampo).Size > 0) then
      Result := Q.FieldByName(ACampo).Size;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.AjustarTextoParaCampo(
  AConn: TADOConnection; const ATabela, ACampo, AValor: string;
  const APadrao: Integer = 255): string;
begin
  Result := Copy(Trim(AValor), 1,
    TamanhoCampoTexto(AConn, ATabela, ACampo, APadrao));
end;

function TFrmConsultaExecutantesProgramados.NormalizarDataChave(
  const AData: TDateTime): string;
begin
  if AData > 0 then
    Result := FormatDateTime('yyyymmdd', Trunc(AData))
  else
    Result := '';
end;

function TFrmConsultaExecutantesProgramados.NormalizarDataChave(
  const S: string): string;
var
  D: TDateTime;
begin
  if TryParseDataChave(S, D) then
    Exit(NormalizarDataChave(D));

  Result := UpperCase(Trim(S));
  Result := StringReplace(Result, '.', '', [rfReplaceAll]);
  Result := StringReplace(Result, '/', '', [rfReplaceAll]);
  Result := StringReplace(Result, '-', '', [rfReplaceAll]);
end;

function TFrmConsultaExecutantesProgramados.DataSAP(const AData: TDateTime): string;
begin
  Result := FrmPrincipal.dataSAP(FormatDateTime('dd/mm/yyyy', Trunc(AData)));
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarRetornoExecutante(
  const AIdExec: Integer;
  const AMensagem, ANumeroRT, AStatus: string);
var
  Q: TADOQuery;
  SQL, StatusSafe, MensagemSafe, RTSafe: string;
begin
  if AIdExec <= 0 then
    Exit;

  StatusSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionColibri,
    'tblProgramacaoExecutante',
    'RT_Status',
    StatusRTNormalizado(AStatus),
    30
  );

  MensagemSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionColibri,
    'tblProgramacaoExecutante',
    'RT_Mensagem',
    Trim(AMensagem),
    255
  );

  RTSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionColibri,
    'tblProgramacaoExecutante',
    'RT',
    Trim(ANumeroRT),
    50
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;

    SQL :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  RT_Status = :RT_Status, ' +
      '  RT_Mensagem = :RT_Mensagem ';

    if Trim(ANumeroRT) <> '' then
      SQL := SQL + ', RT = :RT ';

    SQL := SQL + 'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    Q.SQL.Text := SQL;
    Q.Parameters.ParamByName('RT_Status').Value := StatusSafe;
    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemSafe;

    if Trim(ANumeroRT) <> '' then
      Q.Parameters.ParamByName('RT').Value := RTSafe;

    Q.Parameters.ParamByName('idProgramacaoExecutante').Value := AIdExec;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AddKeySet(
  ALst: TStringList; const AValor: string);
var
  S: string;
begin
  if ALst = nil then
    Exit;

  S := UpperCase(Trim(AValor));
  if S = '' then
    Exit;

  if ALst.IndexOf(S) < 0 then
    ALst.Add(S);
end;

function TFrmConsultaExecutantesProgramados.HasKeySet(
  ALst: TStringList; const AValor: string): Boolean;
var
  S: string;
begin
  Result := False;

  if ALst = nil then
    Exit;

  S := UpperCase(Trim(AValor));
  if S = '' then
    Exit;

  Result := ALst.IndexOf(S) >= 0;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarRetornoTabelaRT(
  const AIdRT: Integer;
  const AMensagem, ANumeroRT, AStatus: string);
var
  Q: TADOQuery;
  SQL, StatusSafe, MensagemSafe, RTSafe: string;
begin
  if AIdRT <= 0 then
    Exit;

  StatusSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT_Status',
    StatusRTNormalizado(AStatus),
    30
  );

  MensagemSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT_Mensagem',
    Trim(AMensagem),
    255
  );


  RTSafe := AjustarTextoParaCampo(
    FrmDataModule.ADOConnectionRT,
    'tblProgramacaoRT',
    'RT',
    Trim(ANumeroRT),
    50
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;

    SQL :=
      'UPDATE tblProgramacaoRT SET ' +
      '  RT_Status = :RT_Status, ' +
      '  RT_Mensagem = :RT_Mensagem ';

    if Trim(ANumeroRT) <> '' then
      SQL := SQL + ', RT = :RT ';

    SQL := SQL + 'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.SQL.Text := SQL;

    Q.Parameters.ParamByName('RT_Status').Value := StatusSafe;
    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemSafe;

    if Trim(ANumeroRT) <> '' then
      Q.Parameters.ParamByName('RT').Value := RTSafe;

    Q.Parameters.ParamByName('idProgramacaoRT').Value := AIdRT;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ToolButton1Click(Sender: TObject);
begin
  DebugCamposDataset(FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet);
end;

procedure TFrmConsultaExecutantesProgramados.DebugCamposDataset(DS: TDataSet);
var
  I: Integer;
  S: string;
begin
  if DS = nil then
  begin
    ShowMessage('DS=nil');
    Exit;
  end;

  S := '';
  for I := 0 to DS.FieldCount - 1 do
    S := S + DS.Fields[I].FieldName + sLineBreak;

  ShowMessage(S);
end;

function TFrmConsultaExecutantesProgramados.RateioValido(
  const ACentroCusto, ADiagramaRede, AOperRede, AElementoPEP: string): Boolean;
begin
  Result :=
    (Trim(ACentroCusto) <> '') or
    ((Trim(ADiagramaRede) <> '') and (Trim(AOperRede) <> '')) or
    (Trim(AElementoPEP) <> '');
end;

function TFrmConsultaExecutantesProgramados.ObterRateioExecutanteComFallback(
  const AIdProgramacaoExecutante: Integer;
  const ACodigoSAPAtual, ANomeExecutante, AEmpresa: string;
  out ACodigoSAPUsado, ACentroCusto, ADiagramaRede, AOperRede, AElementoPEP: string): Boolean;
var
  Q: TADOQuery;
  CodigoEncontrado: string;
  AchouPorCodigoSAP: Boolean;
begin
  Result := False;
  ACodigoSAPUsado := Trim(ACodigoSAPAtual);
  ACentroCusto := '';
  ADiagramaRede := '';
  AOperRede := '';
  AElementoPEP := '';

  AchouPorCodigoSAP := False;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionConsulta;
    Q.ParamCheck := True;

    { 1. Tenta primeiro pelo CodigoSAP atual }
    if Trim(ACodigoSAPAtual) <> '' then
    begin
      Q.SQL.Text :=
        'SELECT TOP 1 CodigoSAP, CentroCusto, DiagramaRede, OperRede, ElementoPEP ' +
        'FROM tblExecutante ' +
        'WHERE Trim(CodigoSAP) = :pCodigoSAP ' +
        'ORDER BY idExecutante DESC';

      Q.Parameters.ParamByName('pCodigoSAP').Value := Trim(ACodigoSAPAtual);
      Q.Open;

      if not Q.Eof then
      begin
        CodigoEncontrado := Trim(Q.FieldByName('CodigoSAP').AsString);
        ACodigoSAPUsado := CodigoEncontrado;
        ACentroCusto := Trim(Q.FieldByName('CentroCusto').AsString);
        ADiagramaRede := Trim(Q.FieldByName('DiagramaRede').AsString);
        AOperRede := Trim(Q.FieldByName('OperRede').AsString);
        AElementoPEP := Trim(Q.FieldByName('ElementoPEP').AsString);

        Result := RateioValido(ACentroCusto, ADiagramaRede, AOperRede, AElementoPEP);
        Exit;
      end;

      Q.Close;
      Q.SQL.Clear;
      Q.Parameters.Clear;
    end;

    { 2. Só se não achou o CodigoSAP, procura por NomeExecutante + txtEmpresa }
    if (not AchouPorCodigoSAP) and
       (Trim(ANomeExecutante) <> '') and
       (Trim(AEmpresa) <> '') then
    begin
      Q.SQL.Text :=
        'SELECT TOP 1 CodigoSAP, CentroCusto, DiagramaRede, OperRede, ElementoPEP ' +
        'FROM tblExecutante ' +
        'WHERE UCase(Trim(NomeExecutante)) = :pNome ' +
        '  AND UCase(Trim(txtEmpresa)) = :pEmpresa ' +
        'ORDER BY idExecutante DESC';

      Q.Parameters.ParamByName('pNome').Value := UpperCase(Trim(ANomeExecutante));
      Q.Parameters.ParamByName('pEmpresa').Value := UpperCase(Trim(AEmpresa));
      Q.Open;

      if not Q.Eof then
      begin
        CodigoEncontrado := Trim(Q.FieldByName('CodigoSAP').AsString);
        ACodigoSAPUsado := CodigoEncontrado;
        ACentroCusto := Trim(Q.FieldByName('CentroCusto').AsString);
        ADiagramaRede := Trim(Q.FieldByName('DiagramaRede').AsString);
        AOperRede := Trim(Q.FieldByName('OperRede').AsString);
        AElementoPEP := Trim(Q.FieldByName('ElementoPEP').AsString);

        Result := RateioValido(ACentroCusto, ADiagramaRede, AOperRede, AElementoPEP);

        if Result and (Trim(CodigoEncontrado) <> '') and
           (Trim(CodigoEncontrado) <> Trim(ACodigoSAPAtual)) then
          AtualizarCodigoSAPProgramacaoExecutante(AIdProgramacaoExecutante, CodigoEncontrado);

        Exit;
      end;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.ObterInfoTblPlataforma(
  const APlataforma: string;
  out ABooleanProntidao: Boolean;
  out ACentroCusto, ADiagramaRede, AOperRede, AElementoPEP: string): Boolean;
var
  Q: TADOQuery;
begin
  Result := False;
  ABooleanProntidao := False;
  ACentroCusto := '';
  ADiagramaRede := '';
  AOperRede := '';
  AElementoPEP := '';

  if Trim(APlataforma) = '' then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionConsulta;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'SELECT TOP 1 booleanProntidao, CentroCusto, DiagramaRede, OperRede, ElementoPEP ' +
      'FROM tblPlataforma ' +
      'WHERE UCase(Trim(Plataforma)) = :pPlataforma ' +
      '   OR UCase(Trim(NomeSAP)) = :pPlataforma';

    Q.Parameters.ParamByName('pPlataforma').Value := UpperCase(Trim(APlataforma));
    Q.Open;

    if not Q.Eof then
    begin
      ABooleanProntidao := Q.FieldByName('booleanProntidao').AsBoolean;
      ACentroCusto := Trim(Q.FieldByName('CentroCusto').AsString);
      ADiagramaRede := Trim(Q.FieldByName('DiagramaRede').AsString);
      AOperRede := Trim(Q.FieldByName('OperRede').AsString);
      AElementoPEP := Trim(Q.FieldByName('ElementoPEP').AsString);
      Result := True;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.DescricaoStatusRTSAP(
  const AStatus: string): string;
begin
  case StrToIntDef(Trim(AStatus), -1) of
    1:  Result := 'Em Coleta';
    2:  Result := 'Unitizado';
    3:  Result := 'Liberado para Programação';
    4:  Result := 'Programado';
    5:  Result := 'Desprogramado';
    6:  Result := 'Liberado via Desdobramento';
    7:  Result := 'Atendido';
    8:  Result := 'Não Atendido';
    9:  Result := 'Cancelado';
    10: Result := 'Faltoso';
    11: Result := 'Desdobrado';
    12: Result := 'Solicitado';
    13: Result := 'Suspenso';
    14: Result := 'Desmembrado';
    15: Result := 'Solicitado via Desmembramento';
    16: Result := 'Solicitado via Desdobramento';
    17: Result := 'Detalhado';
    18: Result := 'Suspenso Fat. Externo';
    19: Result := 'Excluído';
    20: Result := 'Embarcado';
  else
    Result := '';
  end;
end;

function TFrmConsultaExecutantesProgramados.NormalizarTextoChave(
  const S: string): string;
begin
  Result := Trim(UpperCase(S));
end;

function TFrmConsultaExecutantesProgramados.NormalizarHoraChave(
  const S: string): string;
var
  H: string;
begin
  H := Trim(S);

  if H = '' then
    Exit('');

  // mantém só HH:MM
  if Length(H) >= 5 then
    H := Copy(H, 1, 5);

  Result := H;
end;

function TFrmConsultaExecutantesProgramados.NormalizarDocumentoChave(
  const S: string): string;
var
  Doc: string;
begin
  Doc := Trim(S);

  if Doc = '' then
    Exit('');

  // Para CPF remove máscara; para passaporte também remove separadores simples
  Doc := FrmPrincipal.SomenteNumero(Doc);

  if Doc <> '' then
    Result := Doc
  else
    Result := NormalizarTextoChave(S);
end;

function TFrmConsultaExecutantesProgramados.ChavePassageiro(
  const CodigoSAP, TipoDoc, Documento: string): string;
var
  Cod, TDoc, Doc: string;
begin
  Cod  := Trim(CodigoSAP);
  TDoc := NormalizarTextoChave(TipoDoc);
  Doc  := NormalizarDocumentoChave(Documento);

  if Cod <> '' then
    Exit('SAP=' + Cod);

  if (TDoc <> '') and (Doc <> '') then
    Exit(TDoc + '=' + Doc);

  if Doc <> '' then
    Exit('DOC=' + Doc);

  Result := '';
end;

function TFrmConsultaExecutantesProgramados.ChavePassageiroSAPImport(
  const PERNR, TipoDoc, Documento: string): string;
var
  VPernr, VTipoDoc, VDoc: string;
begin
  VPernr   := Trim(PERNR);
  VTipoDoc := NormalizarTextoChave(TipoDoc);
  VDoc     := NormalizarDocumentoChave(Documento);

  if VPernr <> '' then
    Exit('SAP=' + VPernr);

  if (VTipoDoc <> '') and (VDoc <> '') then
    Exit(VTipoDoc + '=' + VDoc);

  if VDoc <> '' then
    Exit('DOC=' + VDoc);

  Result := '';
end;
function TFrmConsultaExecutantesProgramados.ChaveRTIdaSAPImport(
  const DadosSAP: TRTSapImportDados): string;
var
  ChPass: string;
begin
  ChPass := ChavePassageiroSAPImport(
    DadosSAP.PERNR,
    DadosSAP.TipoDoc,
    DadosSAP.Documento
  );

  Result :=
    ChPass + '|' +
    'DT='   + NormalizarDataChave(DadosSAP.DataViagem) + '|' +
    'HR='   + NormalizarHoraChave(DadosSAP.HoraViagem) + '|' +
    'ORG='  + NormalizarTextoChave(DadosSAP.Origem) + '|' +
    'DST='  + NormalizarTextoChave(DadosSAP.Destino) + '|' +
    'TIPO=' + NormalizarTextoChave(DadosSAP.QMART) + '|' +
    'CENT=' + NormalizarTextoChave(DadosSAP.IWERK) + '|' +
    'GRP='  + NormalizarTextoChave(DadosSAP.INGRP) + '|' +
    'MOD='  + NormalizarTextoChave(DadosSAP.RT_Modal) + '|' +
    'CLS='  + NormalizarTextoChave(DadosSAP.RT_Classe);
end;
procedure TFrmConsultaExecutantesProgramados.ImportarRTsSAPPeriodo;
begin
  ImportarRTsSAPPeriodoPorTipo('R3', dataInicio.Date, dataFim.Date);
  ImportarRTsSAPPeriodoPorTipo('R7', dataInicio.Date, dataFim.Date);

  VincularRTSapImportComProgramacao(dataInicio.Date, dataFim.Date);

  actProcurarSAPImport.Execute;
  actProcurarProgramacaoRT.Execute;
  actProcurarProgramacaoExecutante.Execute;

  MessageBox(0,PChar('Importação das RTs do SAP e vinculação concluídas.'),
        'Colibri',MB_ICONINFORMATION);
end;

procedure TFrmConsultaExecutantesProgramados.InserirOuAtualizarRTSapImport(
  const DadosSAP: TRTSapImportDados);
var
  QSel, QExec: TADOQuery;
  IdExistente: Integer;
  ChPassageiro, ChaveIda, ChaveVolta, ChaveCompleta: string;
  RTCancelada: Boolean;

  function DateOrNull(const AValue: TDateTime): Variant;
  begin
    if AValue > 0 then
      Result := AValue
    else
      Result := Null;
  end;

begin
  IdExistente := 0;

  ChPassageiro := ChavePassageiroSAPImport(
    DadosSAP.PERNR,
    DadosSAP.TipoDoc,
    DadosSAP.Documento
  );

  ChaveIda := ChaveRTIdaSAPImport(DadosSAP);

  // Por enquanto, na importação do grid do SAP estamos tratando cada linha
  // como um trecho/registro individual. Portanto, ChaveVolta fica vazia.
  ChaveVolta := '';
  ChaveCompleta := ChaveIda;

  RTCancelada :=
    (Trim(DadosSAP.StatusItem) = '09') or
    ContainsText(DadosSAP.StatusDescricao, 'CANCEL');

  QSel := TADOQuery.Create(nil);
  QExec := TADOQuery.Create(nil);
  try
    QSel.Connection := FrmDataModule.ADOConnectionRT;
    QExec.Connection := FrmDataModule.ADOConnectionRT;

    // Procura registro já importado da mesma RT/trecho
    QSel.SQL.Text :=
      'SELECT TOP 1 idRTSapImport '+
      'FROM tblRTSapImport '+
      'WHERE QMNUM = :QMNUM '+
      '  AND Origem = :Origem '+
      '  AND txtDestino = :txtDestino '+
      '  AND DataViagem = :DataViagem '+
      '  AND HoraViagem = :HoraViagem '+
      '  AND ChavePassageiro = :ChavePassageiro '+
      'ORDER BY idRTSapImport DESC';

    QSel.Parameters.ParamByName('QMNUM').Value := Trim(DadosSAP.QMNUM);
    QSel.Parameters.ParamByName('Origem').Value := Trim(DadosSAP.Origem);
    QSel.Parameters.ParamByName('txtDestino').Value := Trim(DadosSAP.Destino);
    QSel.Parameters.ParamByName('DataViagem').Value := DateOrNull(DadosSAP.DataViagem);
    QSel.Parameters.ParamByName('HoraViagem').Value := NormalizarHoraChave(DadosSAP.HoraViagem);
    QSel.Parameters.ParamByName('ChavePassageiro').Value := ChPassageiro;
    QSel.Open;

    if not QSel.Eof then
      IdExistente := QSel.FieldByName('idRTSapImport').AsInteger;

    QExec.SQL.Clear;

    if IdExistente > 0 then
    begin
      QExec.SQL.Text :=
        'UPDATE tblRTSapImport SET '+
        '  DataImportacao = :DataImportacao, '+
        '  OrigemConsulta = :OrigemConsulta, '+
        '  PeriodoInicio = :PeriodoInicio, '+
        '  PeriodoFim = :PeriodoFim, '+
        '  QMNUM = :QMNUM, '+
        '  QMART = :QMART, '+
        '  IWERK = :IWERK, '+
        '  INGRP = :INGRP, '+
        '  Origem = :Origem, '+
        '  txtDestino = :txtDestino, '+
        '  DataViagem = :DataViagem, '+
        '  HoraViagem = :HoraViagem, '+
        '  PERNR = :PERNR, '+
        '  TipoDoc = :TipoDoc, '+
        '  Documento = :Documento, '+
        '  Passageiro = :Passageiro, '+
        '  QMTXT = :QMTXT, '+
        '  RT_Modal = :RT_Modal, '+
        '  RT_Classe = :RT_Classe, '+
        '  StatusItem = :StatusItem, '+
        '  StatusDescricao = :StatusDescricao, '+
        '  RT_Cancelada = :RT_Cancelada, '+
        '  ChavePassageiro = :ChavePassageiro, '+
        '  ChaveIda = :ChaveIda, '+
        '  ChaveVolta = :ChaveVolta, '+
        '  ChaveCompleta = :ChaveCompleta, '+
        '  idProgramacaoRT = :idProgramacaoRT, '+
        '  idProgramacaoExecutante = :idProgramacaoExecutante, '+
        '  ImportadoColibri = :ImportadoColibri, '+
        '  Observacao = :Observacao '+
        'WHERE idRTSapImport = :idRTSapImport';
    end
    else
    begin
      QExec.SQL.Text :=
        'INSERT INTO tblRTSapImport ('+
        '  DataImportacao, OrigemConsulta, PeriodoInicio, PeriodoFim, '+
        '  QMNUM, QMART, IWERK, INGRP, '+
        '  Origem, txtDestino, DataViagem, HoraViagem, '+
        '  PERNR, TipoDoc, Documento, Passageiro, QMTXT, '+
        '  RT_Modal, RT_Classe, '+
        '  StatusItem, StatusDescricao, RT_Cancelada, '+
        '  ChavePassageiro, ChaveIda, ChaveVolta, ChaveCompleta, '+
        '  idProgramacaoRT, idProgramacaoExecutante, ImportadoColibri, Observacao '+
        ') VALUES ('+
        '  :DataImportacao, :OrigemConsulta, :PeriodoInicio, :PeriodoFim, '+
        '  :QMNUM, :QMART, :IWERK, :INGRP, '+
        '  :Origem, :txtDestino, :DataViagem, :HoraViagem, '+
        '  :PERNR, :TipoDoc, :Documento, :Passageiro, :QMTXT, '+
        '  :RT_Modal, :RT_Classe, '+
        '  :StatusItem, :StatusDescricao, :RT_Cancelada, '+
        '  :ChavePassageiro, :ChaveIda, :ChaveVolta, :ChaveCompleta, '+
        '  :idProgramacaoRT, :idProgramacaoExecutante, :ImportadoColibri, :Observacao '+
        ')';
    end;

    QExec.Parameters.ParamByName('DataImportacao').Value :=
      DateOrNull(DadosSAP.DataImportacao);
    QExec.Parameters.ParamByName('OrigemConsulta').Value :=
      Trim(DadosSAP.OrigemConsulta);
    QExec.Parameters.ParamByName('PeriodoInicio').Value :=
      DateOrNull(DadosSAP.PeriodoInicio);
    QExec.Parameters.ParamByName('PeriodoFim').Value :=
      DateOrNull(DadosSAP.PeriodoFim);

    QExec.Parameters.ParamByName('QMNUM').Value := Trim(DadosSAP.QMNUM);
    QExec.Parameters.ParamByName('QMART').Value := Trim(DadosSAP.QMART);
    QExec.Parameters.ParamByName('IWERK').Value := Trim(DadosSAP.IWERK);
    QExec.Parameters.ParamByName('INGRP').Value := Trim(DadosSAP.INGRP);

    QExec.Parameters.ParamByName('Origem').Value := Trim(DadosSAP.Origem);
    QExec.Parameters.ParamByName('txtDestino').Value := Trim(DadosSAP.Destino);
    QExec.Parameters.ParamByName('DataViagem').Value := DateOrNull(DadosSAP.DataViagem);
    QExec.Parameters.ParamByName('HoraViagem').Value :=
      NormalizarHoraChave(DadosSAP.HoraViagem);

    QExec.Parameters.ParamByName('PERNR').Value := Trim(DadosSAP.PERNR);
    QExec.Parameters.ParamByName('TipoDoc').Value := Trim(DadosSAP.TipoDoc);
    QExec.Parameters.ParamByName('Documento').Value := Trim(DadosSAP.Documento);
    QExec.Parameters.ParamByName('Passageiro').Value := Trim(DadosSAP.Passageiro);
    QExec.Parameters.ParamByName('QMTXT').Value := Trim(DadosSAP.QMTXT);

    QExec.Parameters.ParamByName('RT_Modal').Value := Trim(DadosSAP.RT_Modal);
    QExec.Parameters.ParamByName('RT_Classe').Value := Trim(DadosSAP.RT_Classe);

    QExec.Parameters.ParamByName('StatusItem').Value := Trim(DadosSAP.StatusItem);
    QExec.Parameters.ParamByName('StatusDescricao').Value :=
      Trim(DadosSAP.StatusDescricao);
    QExec.Parameters.ParamByName('RT_Cancelada').Value := RTCancelada;

    QExec.Parameters.ParamByName('ChavePassageiro').Value := ChPassageiro;
    QExec.Parameters.ParamByName('ChaveIda').Value := ChaveIda;
    QExec.Parameters.ParamByName('ChaveVolta').Value := ChaveVolta;
    QExec.Parameters.ParamByName('ChaveCompleta').Value := ChaveCompleta;

    QExec.Parameters.ParamByName('idProgramacaoRT').Value :=
      DadosSAP.idProgramacaoRT;
    QExec.Parameters.ParamByName('idProgramacaoExecutante').Value :=
      DadosSAP.idProgramacaoExecutante;
    QExec.Parameters.ParamByName('ImportadoColibri').Value :=
      DadosSAP.ImportadoColibri;
    QExec.Parameters.ParamByName('Observacao').Value :=
      Trim(DadosSAP.Observacao);

    if IdExistente > 0 then
      QExec.Parameters.ParamByName('idRTSapImport').Value := IdExistente;

    QExec.ExecSQL;

  finally
    QSel.Free;
    QExec.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.ChaveRTIda(
  const Dados: TDadosRT): string;
var
  ChPass: string;
begin
  ChPass := ChavePassageiro(
    Dados.MatriculaPax,
    IfThen(Trim(Dados.CPF) <> '', 'C',
      IfThen(Trim(Dados.Passaporte) <> '', 'P', '')),
    IfThen(Trim(Dados.CPF) <> '', Dados.CPF, Dados.Passaporte)
  );

  Result :=
    ChPass + '|' +
    'DT='   + NormalizarDataChave(Dados.DataIda) + '|' +
    'HR='   + NormalizarHoraChave(Dados.HoraIda) + '|' +
    'ORG='  + NormalizarTextoChave(Dados.Origem) + '|' +
    'DST='  + NormalizarTextoChave(Dados.Destino) + '|' +
    'TIPO=' + NormalizarTextoChave(Dados.TipoRT) + '|' +
    'CENT=' + NormalizarTextoChave(Dados.CentroPlan) + '|' +
    'GRP='  + NormalizarTextoChave(Dados.GrpPlan) + '|' +
    'MOD='  + NormalizarTextoChave(Dados.Modal) + '|' +
    'CLS='  + NormalizarTextoChave(Dados.Classe);
end;

function TFrmConsultaExecutantesProgramados.ChaveRTVolta(
  const Dados: TDadosRT): string;
var
  ChPass: string;
begin
  if not Dados.booleanRecolhimento then
    Exit('');

  if (Trim(Dados.DataVolta) = '') or
     (Trim(Dados.HoraVolta) = '') or
     (Trim(Dados.Destino) = '') or
     (Trim(Dados.Retorno) = '') then
    Exit('');

  ChPass := ChavePassageiro(
    Dados.MatriculaPax,
    IfThen(Trim(Dados.CPF) <> '', 'C',
      IfThen(Trim(Dados.Passaporte) <> '', 'P', '')),
    IfThen(Trim(Dados.CPF) <> '', Dados.CPF, Dados.Passaporte)
  );

  Result :=
    ChPass + '|' +
    'DT='   + NormalizarDataChave(Dados.DataVolta) + '|' +
    'HR='   + NormalizarHoraChave(Dados.HoraVolta) + '|' +
    'ORG='  + NormalizarTextoChave(Dados.Destino) + '|' +
    'DST='  + NormalizarTextoChave(Dados.Retorno) + '|' +
    'TIPO=' + NormalizarTextoChave(Dados.TipoRT) + '|' +
    'CENT=' + NormalizarTextoChave(Dados.CentroPlan) + '|' +
    'GRP='  + NormalizarTextoChave(Dados.GrpPlan) + '|' +
    'MOD='  + NormalizarTextoChave(Dados.Modal) + '|' +
    'CLS='  + NormalizarTextoChave(Dados.Classe);
end;

function TFrmConsultaExecutantesProgramados.ChaveRTCompleta(
  const Dados: TDadosRT): string;
var
  Ida, Volta: string;
begin
  Ida := ChaveRTIda(Dados);
  Volta := ChaveRTVolta(Dados);

  if Volta <> '' then
    Result := Ida + '||VOLTA=' + Volta
  else
    Result := Ida;
end;

procedure TFrmConsultaExecutantesProgramados.GerarMultiplasRTsArray;
type
  TRTItem = record
    IdExec: string;
    IdRT: string;
    Linha: string;
    Acao: string;
    RtNumero: string;
    TipoRT: string;
    Pernr: string;
    ContratoFix: string;

    TipoDoc: string;
    Documento: string;
    Passageiro: string;

    QmTxt: string;
    Requis: string;
    Pessoa: string;
    Fone: string;
    Iwerk: string;
    Ingrp: string;
    CodTrasg: string;
    Classe: string;
    Origem: string;
    Destino: string;
    Retorno: string;
    DtIda: string;
    HIda: string;
    DtVolta: string;
    HVolta: string;
    CCusto: string;
    DiagramaRede: string;
    OperRede: string;
    ElementoPEP: string;
    Cobertura: string;
    Recolhimetno: string;
  end;

var
  EnderecoVbs, EnderecoLog, OrigemPlataformas: string;
  DS, DSRT: TDataSet;
  Bmk: TBookmark;
  LogRetorno: TStringList;
  Itens: array of TRTItem;
  ItemCount: Integer;
  HasBookmark: Boolean;

  QtBloqueados, QtProntos: Integer;
  MotivoBloqueio: string;

  Dados: TDadosRT;
  Horario: THorario;
  PlataformaOrigem, PlataformaDestino, PlataformaRetorno: TDadosPlataforma;
  TesteRT: TRT;
  CodicoCusto: TCodicoCusto;
  i, seq: Integer;
  A: TRTItem;
  HoraAtual, HoraCobertura: TDateTime;
  CodigoSAPAtual, CodigoSAPFinal, TipoRTFinal: string;

  function VbsQuote(const S: string): string;
  begin
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"';
  end;

  procedure MemoAdd(const S: string);
  begin
    MemoSAP.Lines.Add(S);
  end;

  function BoolTo01(const B: Boolean): string;
  begin
    if B then
      Result := '1'
    else
      Result := '0';
  end;

  procedure AddItem(const AItem: TRTItem);
  begin
    SetLength(Itens, ItemCount + 1);
    Itens[ItemCount] := AItem;
    Inc(ItemCount);
  end;

  function BuildItemLine(const A: TRTItem): string;
  begin
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

        VbsQuote(A.TipoDoc) + ',' +
        VbsQuote(A.Documento) + ',' +
        VbsQuote(A.Passageiro) + ',' +

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
        VbsQuote(A.Retorno) + ',' +

        VbsQuote(A.DtIda) + ',' +
        VbsQuote(A.HIda) + ',' +
        VbsQuote(A.DtVolta) + ',' +
        VbsQuote(A.HVolta) + ',' +

        VbsQuote(A.CCusto) + ',' +
        VbsQuote(A.DiagramaRede) + ',' +
        VbsQuote(A.OperRede) + ',' +
        VbsQuote(A.ElementoPEP) + ',' +
        VbsQuote(A.Cobertura) + ',' +
        VbsQuote(A.Recolhimetno) +
      ')';
  end;

  function CriarRegistroTabelaRT_Obrigatorio(const ADados: TDadosRT): Integer;
  begin
    Result := CriarRegistroTabelaRT(ADados);
  end;

begin
  EnderecoVbs := ExtractFilePath(ParamStr(0)) + 'RT_R3_R7.vbs';
  EnderecoLog := ExtractFilePath(ParamStr(0)) + 'sap_log_rts.txt';
  OrigemPlataformas := FrmPrincipal.OrigemPlataformas;
  QtProntos:= 0;
  QtBloqueados:= 0;
  MemoSAP.Lines.Clear;

  FrmDataModule.ADOQueryConfigRT.Active := False;
  FrmDataModule.ADOQueryConfigRT.Active := True;

  DS := FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet;
  DSRT := FrmDataModule.DataSourceConfigRT.DataSet;

  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
  begin
    MessageBox(0,PChar('Não há registros para emitir RT.'),
        'Colibri',MB_ICONINFORMATION);
    Exit;
  end;

  ItemCount := 0;
  SetLength(Itens, 0);
  seq := 0;

  HasBookmark := False;
  if DS.Active and (not DS.IsEmpty) then
  begin
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  try
    DS.DisableControls;
    try
      DS.First;
      while not DS.Eof do
      begin
        MotivoBloqueio := '';
        if not PodeGerarRTNoRegistro(DS, MotivoBloqueio) then
        begin
          Inc(QtBloqueados);
          DS.Next;
          Continue;
        end;
        //-----------------------------------------------
        Inc(QtProntos);
        Dados.CentroPlan    := Trim(DSRT.FieldByName('RT_CentroPlan').AsString);
        Dados.GrpPlan       := Trim(DSRT.FieldByName('RT_GrpPlan').AsString);
        Dados.Requisitante  := Trim(DSRT.FieldByName('RT_Requisitante').AsString);
        Dados.TelContato    := Trim(DSRT.FieldByName('RT_TelContato').AsString);
        Dados.PessoaContato := Trim(DSRT.FieldByName('RT_PessoaContato').AsString);
        Dados.HoraCobertura := DSRT.FieldByName('RT_HoraCobertura').AsDateTime;
        //-----------------------------------------------
        HoraAtual     := Frac(Now);
        HoraCobertura := Frac(Dados.HoraCobertura);
        Dados.RTCobertura := (HoraAtual > HoraCobertura);
        //-----------------------------------------------
        Dados.idProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger;
        //-----------------------------------------------
        PlataformaOrigem   := TProgramacaoRTUtils.DadosPlataforma_RT(DS.FieldByName('Origem').AsString);
        PlataformaDestino  := TProgramacaoRTUtils.DadosPlataforma_RT(DS.FieldByName('txtDestino').AsString);
        PlataformaRetorno  := TProgramacaoRTUtils.DadosPlataforma_RT(DS.FieldByName('RecolherPara').AsString);
        //-----------------------------------------------
        Dados.Origem       := PlataformaOrigem.NomeSAP;
        Dados.Destino      := PlataformaDestino.NomeSAP;
        Dados.Retorno      := PlataformaRetorno.NomeSAP;
        //-----------------------------------------------
        CodigoSAPAtual := Trim(DS.FieldByName('CodigoSAP').AsString);

        // remove zero à esquerda, persiste em tblProgramacaoExecutante e tblExecutante
        PersistirCodigoSAPNormalizado(
          DS.FieldByName('idProgramacaoExecutante').AsInteger,
          DS.FieldByName('NomeExecutante').AsString,
          DS.FieldByName('Empresa').AsString,
          CodigoSAPAtual,
          CodigoSAPAtual
        );

        // tenta reclassificar para R3 usando o valor já normalizado
        TentarReclassificarExecutanteParaR3(
          DS.FieldByName('idProgramacaoExecutante').AsInteger,
          DS.FieldByName('NomeExecutante').AsString,
          DS.FieldByName('Empresa').AsString,
          CodigoSAPAtual,
          CodigoSAPFinal,
          TipoRTFinal
        );

        if Trim(CodigoSAPFinal) = '' then
          CodigoSAPFinal := NormalizarCodigoSAP(CodigoSAPAtual);

        Dados.MatriculaPax := CodigoSAPFinal;
        Dados.Modal        := Trim(DS.FieldByName('RT_Modal').AsString);

        if TipoRTFinal <> '' then
          Dados.TipoRT := TipoRTFinal
        else
          Dados.TipoRT := DeterminarTipoRTAutomatico(CodigoSAPFinal);
        //-----------------------------------------------
        Dados.Classe       := Trim(DS.FieldByName('RT_Classe').AsString);
        Dados.CPF          := Trim(DS.FieldByName('Documento').AsString);
        Dados.Passaporte   := Trim(DS.FieldByName('OutroDocumento').AsString);
        Dados.booleanRecolhimento := DS.FieldByName('booleanRecolhimento').AsBoolean;
        //-----------------------------------------------
        if Trim(Dados.Retorno) = '' then
          Dados.Retorno := Dados.Origem;
        //-----------------------------------------------
        Dados.DataIda := DataSAP(DS.FieldByName('DataProgramacao').AsDateTime);
        Dados.DescricaoRT := Dados.Origem + ': ' + DS.FieldByName('NomeExecutante').AsString;
        //-----------------------------------------------
        //Preencher data de Volta
        //-----------------------------------------------
        Horario := HoraIdaHoraVolta(
          PlataformaDestino.booleanTurno2,
          DS.FieldByName('DataProgramacao').AsDateTime
        );

        // valores calculados
        Dados.HoraIda   := Trim(Horario.HoraIda);
        Dados.HoraVolta := Trim(Horario.HoraVolta);
        Dados.DataVolta := DataSAP(Horario.DataVolta);

        // respeita preenchimento manual da programação
        if Trim(DS.FieldByName('RT_HoraIda').AsString) <> '' then
          Dados.HoraIda := Trim(DS.FieldByName('RT_HoraIda').AsString);

        if Trim(DS.FieldByName('RT_HoraVolta').AsString) <> '' then
          Dados.HoraVolta := Trim(DS.FieldByName('RT_HoraVolta').AsString);

        if not VarIsNull(DS.FieldByName('DataVolta').Value) then
          Dados.DataVolta := DataSAP(DS.FieldByName('DataVolta').AsDateTime);

        // fallback final
        if Trim(Dados.HoraIda) = '' then
          Dados.HoraIda := '07:00';

        if Dados.booleanRecolhimento and (Trim(Dados.HoraVolta) = '' ) then
          Dados.HoraVolta := '16:30';

        if Dados.booleanRecolhimento and (Trim(Dados.DataVolta) = '') then
          Dados.DataVolta := Dados.DataIda;
        //----------------------------------------------------------


        if PlataformaDestino.booleanProntidao then
        begin
          Dados.CentroCusto  := PlataformaDestino.CentroCusto;
          Dados.DiagramaRede := PlataformaDestino.DiagramaRede;
          Dados.OperRede     := PlataformaDestino.OperRede;
          Dados.ElementoPEP  := PlataformaDestino.ElementoPEP;
        end
        else
        begin
          CodicoCusto := CustoExecutante(Dados.MatriculaPax, Dados.CPF, Dados.Passaporte);

          if (Trim(CodicoCusto.CentroCusto) = '') and
             (Trim(CodicoCusto.ElementoPEP) = '') and
             ((Trim(CodicoCusto.DiagramaRede) = '') or (Trim(CodicoCusto.OperRede) = '')) then
          begin
            Dados.CentroCusto  := PlataformaDestino.CentroCusto;
            Dados.DiagramaRede := PlataformaDestino.DiagramaRede;
            Dados.OperRede     := PlataformaDestino.OperRede;
            Dados.ElementoPEP  := PlataformaDestino.ElementoPEP;
          end
          else
          begin
            Dados.CentroCusto  := CodicoCusto.CentroCusto;
            Dados.DiagramaRede := CodicoCusto.DiagramaRede;
            Dados.OperRede     := CodicoCusto.OperRede;
            Dados.ElementoPEP  := CodicoCusto.ElementoPEP;
          end;
        end;

        if (Dados.Origem <> Dados.Destino) and
           (not PlataformaDestino.booleanNaoCriarRT) and
           (not PlataformaOrigem.booleanNaoCriarRT) then
        begin
          Dados.ChavePassageiro := ChavePassageiro(
              Dados.MatriculaPax,
              IfThen(Trim(Dados.CPF) <> '', 'C',
                IfThen(Trim(Dados.Passaporte) <> '', 'P', '')),
              IfThen(Trim(Dados.CPF) <> '', Dados.CPF, Dados.Passaporte)
            );

            Dados.ChaveIda := ChaveRTIda(Dados);
            Dados.ChaveVolta := ChaveRTVolta(Dados);
            Dados.ChaveCompleta := ChaveRTCompleta(Dados);

          TesteRT := Existe_RT(Dados);
          Inc(seq);

          A := Default(TRTItem);
          A.IdExec         := IntToStr(Dados.idProgramacaoExecutante);
          A.IdRT           := '';
          A.Linha          := IntToStr(seq);
          A.Acao           := 'CRIAR';
          A.RtNumero       := '';
          A.TipoRT         := Dados.TipoRT;
          A.Pernr          := Dados.MatriculaPax;
          A.ContratoFix    := '';
          A.QmTxt          := Copy(Dados.DescricaoRT, 1, 40);
          A.Requis         := Dados.Requisitante;
          A.Pessoa         := Dados.PessoaContato;
          A.Fone           := Dados.TelContato;
          A.Iwerk          := Dados.CentroPlan;
          A.Ingrp          := Dados.GrpPlan;
          A.CodTrasg       := Dados.Modal;
          A.Classe         := Dados.Classe;
          A.Origem         := Dados.Origem;
          A.Destino        := Dados.Destino;
          A.Retorno        := Dados.Retorno;
          A.DtIda          := Dados.DataIda;
          A.HIda           := Dados.HoraIda;
          A.DtVolta        := Dados.DataVolta;
          A.HVolta         := Dados.HoraVolta;
          A.CCusto         := Trim(Dados.CentroCusto);
          A.DiagramaRede   := Trim(Dados.DiagramaRede);
          A.OperRede       := Trim(Dados.OperRede);
          A.ElementoPEP    := Trim(Dados.ElementoPEP);
          A.Cobertura      := BoolTo01(Dados.RTCobertura);
          A.Recolhimetno   := BoolTo01(Dados.booleanRecolhimento);

          if Trim(Dados.CPF) <> '' then
          begin
            A.TipoDoc   := 'C';
            A.Documento := Dados.CPF;
          end
          else
          begin
            A.TipoDoc   := 'P';
            A.Documento := Dados.Passaporte;
          end;

          A.Passageiro := Trim(DS.FieldByName('NomeExecutante').AsString);

          if TesteRT.booRTExiste then
          begin
            if Trim(TesteRT.RT_Numero) <> '' then
            begin
              // já existe RT válida; não recria
            end
            else
            begin
              // já existe registro local pendente; reaproveita
              A.IdRT := IntToStr(TesteRT.idProgramacaoRT);
              AddItem(A);
            end;
          end
          else
          begin
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
    if HasBookmark then
    begin
      try
        if DS.BookmarkValid(Bmk) then
          DS.GotoBookmark(Bmk);
      finally
        DS.FreeBookmark(Bmk);
      end;
    end;
  end;

    if ItemCount = 0 then
  begin
    MessageBox(0,
      PChar('Nenhuma RT apta para processamento.' + sLineBreak +
            'Registros prontos: ' + IntToStr(QtProntos) + sLineBreak +
            'Registros bloqueados: ' + IntToStr(QtBloqueados)),
      'Colibri',
      MB_ICONEXCLAMATION);
    Exit;
  end;

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
  //-----------------------------------------------
  MemoAdd('Function ExtrairNumeroRTConflito(ByVal txt)');
  MemoAdd('  Dim re, ms');
  MemoAdd('  ExtrairNumeroRTConflito = ""');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set re = CreateObject("VBScript.RegExp")');
  MemoAdd('  re.Global = False');
  MemoAdd('  re.IgnoreCase = True');
  MemoAdd('  re.Pattern = "nota\s+(\d{6,})"');
  MemoAdd('  Set ms = re.Execute(CStr(txt))');
  MemoAdd('  If ms.Count > 0 Then');
  MemoAdd('    ExtrairNumeroRTConflito = ms(0).SubMatches(0)');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Function ExtrairPrimeiroNumeroLongo(ByVal txt)');
  MemoAdd('  Dim re, ms');
  MemoAdd('  ExtrairPrimeiroNumeroLongo = ""');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set re = CreateObject("VBScript.RegExp")');
  MemoAdd('  re.Global = False');
  MemoAdd('  re.IgnoreCase = True');
  MemoAdd('  re.Pattern = "(\d{6,})"');
  MemoAdd('  Set ms = re.Execute(CStr(txt))');
  MemoAdd('  If ms.Count > 0 Then');
  MemoAdd('    ExtrairPrimeiroNumeroLongo = ms(0).SubMatches(0)');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Function ExtrairNumeroNotaGravada(ByVal txt)');
  MemoAdd('  Dim re, ms');
  MemoAdd('  ExtrairNumeroNotaGravada = ""');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set re = CreateObject("VBScript.RegExp")');
  MemoAdd('  re.Global = False');
  MemoAdd('  re.IgnoreCase = True');
  MemoAdd('  re.Pattern = "nota\s+(\d{6,})\s+gravada"');
  MemoAdd('  Set ms = re.Execute(CStr(txt))');
  MemoAdd('  If ms.Count > 0 Then');
  MemoAdd('    ExtrairNumeroNotaGravada = ms(0).SubMatches(0)');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Function CapturarRTCriadaRapido(ByVal sess)');
  MemoAdd('  Dim st, num, i');
  MemoAdd('  CapturarRTCriadaRapido = ""');
  MemoAdd('  st = Trim(CStr(GetStatusText(sess)))');
  MemoAdd('  num = ExtrairNumeroNotaGravada(st)');
  MemoAdd('  If Len(num) > 0 Then');
  MemoAdd('    CapturarRTCriadaRapido = num');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  For i = 1 To 5');
  MemoAdd('    WScript.Sleep 120');
  MemoAdd('    st = Trim(CStr(GetStatusText(sess)))');
  MemoAdd('    num = ExtrairNumeroNotaGravada(st)');
  MemoAdd('    If Len(num) > 0 Then');
  MemoAdd('      CapturarRTCriadaRapido = num');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  Next');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Function GetPopupMessageText(ByVal sess)');
  MemoAdd('  Dim txt');
  MemoAdd('  GetPopupMessageText = ""');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  If sess.Children.Count > 1 Then');
  MemoAdd('    txt = Trim(CStr(sess.findById("wnd[1]/usr/txtMESSTXT2").Text))');
  MemoAdd('    If Len(txt) > 0 Then GetPopupMessageText = txt');
  MemoAdd('    If Len(GetPopupMessageText) = 0 Then');
  MemoAdd('      txt = Trim(CStr(sess.findById("wnd[1]/usr/txtMESSTXT1").Text))');
  MemoAdd('      If Len(txt) > 0 Then GetPopupMessageText = txt');
  MemoAdd('    End If');
  MemoAdd('    If Len(GetPopupMessageText) = 0 Then');
  MemoAdd('      txt = Trim(CStr(sess.findById("wnd[1]/usr/txtMESSTXT3").Text))');
  MemoAdd('      If Len(txt) > 0 Then GetPopupMessageText = txt');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('  If Len(GetPopupMessageText) = 0 Then');
  MemoAdd('    GetPopupMessageText = Trim(CStr(GetStatusText(sess)))');
  MemoAdd('  End If');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Function MsgConflitoRTExistente(ByVal st)');
  MemoAdd('  Dim s');
  MemoAdd('  s = LCase(Trim(CStr(st)))');
  MemoAdd('  MsgConflitoRTExistente = _');
  MemoAdd('    ((InStr(1, s, "já cadastrado para mesma data", vbTextCompare) > 0) Or _');
  MemoAdd('     (InStr(1, s, "ja cadastrado para mesma data", vbTextCompare) > 0) Or _');
  MemoAdd('     (InStr(1, s, "mesma data e hora de viagem", vbTextCompare) > 0)) And _');
  MemoAdd('    (InStr(1, s, "nota", vbTextCompare) > 0)');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------
  MemoAdd('Sub AbrirTelaTransporte(ByVal cobertura)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  session.findById("wnd[0]/shellcont/shell").selectItem "010","Column01"');
  MemoAdd('  session.findById("wnd[0]/shellcont/shell").ensureVisibleHorizontalItem "010","Column01"');
  MemoAdd('  session.findById("wnd[0]/shellcont/shell").clickLink "010","Column01"');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  If (Trim(cobertura) = "1") Or (UCase(Trim(cobertura)) = "TRUE") Then');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  //-------------------------------------------------------
  MemoAdd('Function GetPopupIndex(ByVal sess)');
  MemoAdd('  Dim idx');
  MemoAdd('  idx = sess.Children.Count - 1');
  MemoAdd('  If idx < 1 Then idx = 1');
  MemoAdd('  GetPopupIndex = idx');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function SafeFocus(ByVal sess, ByVal id)');
  MemoAdd('  SafeFocus = False');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  sess.findById(id).setFocus');
  MemoAdd('  If Err.Number = 0 Then SafeFocus = True');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function MsgCampoContratoObrigatorio(ByVal st)');
  MemoAdd('  Dim s');
  MemoAdd('  s = LCase(Trim(CStr(st)))');
  MemoAdd('  MsgCampoContratoObrigatorio = (InStr(1, s, "contrato", vbTextCompare) > 0) And (InStr(1, s, "obrig", vbTextCompare) > 0)');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function MsgCampoModalObrigatorio(ByVal st)');
  MemoAdd('  Dim s');
  MemoAdd('  s = LCase(Trim(CStr(st)))');
  MemoAdd('  MsgCampoModalObrigatorio = (InStr(1, s, "modal", vbTextCompare) > 0) And (InStr(1, s, "obrig", vbTextCompare) > 0)');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Sub CancelarTodosPopups(sess)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim w');
  MemoAdd('  Do While sess.Children.Count > 1');
  MemoAdd('    w = sess.Children.Count - 1');
  MemoAdd('    sess.findById("wnd[" & w & "]").sendVKey 12');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    If sess.Children.Count > 1 Then');
  MemoAdd('      sess.findById("wnd[" & w & "]/tbar[0]/btn[12]").press');
  MemoAdd('      WScript.Sleep 100');
  MemoAdd('    End If');
  MemoAdd('  Loop');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Sub ResetSAP(sess)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Call CancelarTodosPopups(sess)');
  MemoAdd('  sess.findById("wnd[0]/tbar[0]/okcd").text = "/n"');
  MemoAdd('  sess.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function SafeRunItem(ByVal a, ByVal salvar)');
  MemoAdd('  SafeRunItem = False');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim ok');
  MemoAdd('  ok = ProcessarRT(a(0),a(1),a(2),a(3),a(4),a(5),a(6),a(7), _');
  MemoAdd('                   a(8),a(9),a(10), _');
  MemoAdd('                   a(11),a(12),a(13),a(14),a(15),a(16),a(17),a(18), _');
  MemoAdd('                   a(19),a(20),a(21),a(22),a(23),a(24), _');
  MemoAdd('                   a(25),a(26),a(27),a(28),a(29),a(30),a(31), salvar)');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    LogLine a(0) & "|" & a(1) & "|" & a(2) & "|EXCECAO_VBS|" & Err.Number & "|" & Err.Description');
  MemoAdd('    Err.Clear');
  MemoAdd('    Call ResetSAP(session)');
  MemoAdd('    SafeRunItem = False');
  MemoAdd('  Else');
  MemoAdd('    SafeRunItem = ok');
  MemoAdd('    If ok = False Then Call ResetSAP(session)');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function ResolverContratoObrigatorio(ByVal idExec, ByVal idRT, ByVal linha, ByVal contratoFix, ByVal rowIdx)');
  MemoAdd('  ResolverContratoObrigatorio = False');
  MemoAdd('  Dim st, stType, pathContrato, tbl, startRow, i, w, idLbl, t, filled');
  MemoAdd('  pathContrato = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                 "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CONTRATO[19," & rowIdx & "]"');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set tbl = session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/tblSAPLYSCS_INFADREQTRANSTC_RT3")');
  MemoAdd('  If Not tbl Is Nothing Then');
  MemoAdd('    tbl.VerticalScrollbar.Position = rowIdx');
  MemoAdd('    tbl.CurrentCellRow = rowIdx');
  MemoAdd('    tbl.CurrentCellColumn = "YSCSRTITPAX-CONTRATO"');
  MemoAdd('  End If');
  MemoAdd('  Err.Clear');
  MemoAdd('  session.findById(pathContrato).setFocus');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('  If Len(Trim(contratoFix)) > 0 Then');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    session.findById(pathContrato).text = contratoFix');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    Err.Clear');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('  Else');
  MemoAdd('    startRow = 40');
  MemoAdd('    Do While startRow >= 4');
  MemoAdd('      On Error Resume Next');
  MemoAdd('      session.findById(pathContrato).setFocus');
  MemoAdd('      session.findById("wnd[0]").sendVKey 4');
  MemoAdd('      WScript.Sleep 100');
  MemoAdd('      Err.Clear');
  MemoAdd('      On Error GoTo 0');
  MemoAdd('      w = GetPopupIndex(session)');
  MemoAdd('      filled = False');
  MemoAdd('      For i = startRow To 4 Step -1');
  MemoAdd('        idLbl = "wnd[" & w & "]/usr/lbl[1," & i & "]"');
  MemoAdd('        If SafeFocus(session, idLbl) Then');
  MemoAdd('          On Error Resume Next');
  MemoAdd('          t = Trim(CStr(session.findById(idLbl).Text))');
  MemoAdd('          Err.Clear');
  MemoAdd('          On Error GoTo 0');
  MemoAdd('          If Len(t) > 0 Then');
  MemoAdd('            On Error Resume Next');
  MemoAdd('            session.findById("wnd[" & w & "]").sendVKey 2');
  MemoAdd('            WScript.Sleep 100');
  MemoAdd('            Err.Clear');
  MemoAdd('            On Error GoTo 0');
  MemoAdd('            On Error Resume Next');
  MemoAdd('            filled = (Len(Trim(CStr(session.findById(pathContrato).Text))) > 0)');
  MemoAdd('            Err.Clear');
  MemoAdd('            On Error GoTo 0');
  MemoAdd('            If filled Then Exit For');
  MemoAdd('            startRow = i - 1');
  MemoAdd('            Exit For');
  MemoAdd('          End If');
  MemoAdd('        End If');
  MemoAdd('      Next');
  MemoAdd('      If filled Then Exit Do');
  MemoAdd('      If startRow < 4 Then Exit Do');
  MemoAdd('    Loop');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    If Len(Trim(CStr(session.findById(pathContrato).Text))) = 0 Then');
  MemoAdd('      ResolverContratoObrigatorio = False');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('    Err.Clear');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    Err.Clear');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('  End If');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  ResolverContratoObrigatorio = (stType <> "E")');
  MemoAdd('End Function');
  MemoAdd('');
  //------------------------------------------------------
  MemoAdd('Function MensagemIndicaFalhaNegocio(ByVal txt)');
  MemoAdd('  Dim s');
  MemoAdd('  s = UCase(Trim(CStr(txt)))');
  MemoAdd('  MensagemIndicaFalhaNegocio = False');
  MemoAdd('  If (InStr(1, s, "PASSAGEIRO", vbTextCompare) > 0) And (InStr(1, s, "ATIVO", vbTextCompare) > 0) Then');
  MemoAdd('    If (InStr(1, s, "NAO", vbTextCompare) > 0) Or (InStr(1, s, "N O", vbTextCompare) > 0) Or (InStr(1, s, "INATIVO", vbTextCompare) > 0) Then');
  MemoAdd('      MensagemIndicaFalhaNegocio = True');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('  If InStr(1, s, "NAO ESTA ATIVO", vbTextCompare) > 0 Then');
  MemoAdd('    MensagemIndicaFalhaNegocio = True');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  If InStr(1, s, "N O ESTA ATIVO", vbTextCompare) > 0 Then');
  MemoAdd('    MensagemIndicaFalhaNegocio = True');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  If InStr(1, s, "N O EST ATIVO", vbTextCompare) > 0 Then');
  MemoAdd('    MensagemIndicaFalhaNegocio = True');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('Sub PreencherRateioRT3(ByVal row, ByVal ccusto, ByVal ElementoPEP, ByVal DiagramaRede, ByVal OperRede)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim base, c, pep, r, o');
  MemoAdd('  base = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('         "tblSAPLYSCS_INFADREQTRANSTC_RT3/"');
  MemoAdd('  c   = Trim(CStr(ccusto))');
  MemoAdd('  pep = Trim(CStr(ElementoPEP))');
  MemoAdd('  r   = Trim(CStr(DiagramaRede))');
  MemoAdd('  o   = Trim(CStr(OperRede))');
  MemoAdd('  If Len(c) > 0 Then');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-CCUSTO[63," & row & "]").Text = c');
  MemoAdd('  ElseIf Len(pep) > 0 Then');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-ELEM_PEP[64," & row & "]").Text = pep');
  MemoAdd('  ElseIf (Len(r) > 0) And (Len(o) > 0) Then');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-CCUSTO[63," & row & "]").Text = ""');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-ELEM_PEP[64," & row & "]").Text = ""');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-REDE[65," & row & "]").Text = r');
  MemoAdd('    session.findById(base & "ctxtYSCSRTITPAX-OPER_REDE[66," & row & "]").Text = o');
  MemoAdd('  End If');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function PreencherQmTxt(ByVal sess, ByVal qmtxt, ByVal idExec, ByVal idRT, ByVal linha)');
  MemoAdd('  PreencherQmTxt = False');
  MemoAdd('  Dim pathQmTxt, outTxt, v');
  MemoAdd('  pathQmTxt = "wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/subNOTIF_TYPE:SAPLIQS0:1061/txtVIQMEL-QMTXT"');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  v = Trim(CStr(qmtxt))');
  MemoAdd('  If Len(v) > 40 Then v = Left(v, 40)');
  MemoAdd('  Err.Clear');
  MemoAdd('  sess.findById(pathQmTxt).setFocus');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  Err.Clear');
  MemoAdd('  sess.findById(pathQmTxt).Text = ""');
  MemoAdd('  Err.Clear');
  MemoAdd('  sess.findById(pathQmTxt).Text = v');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  sess.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  Err.Clear');
  MemoAdd('  outTxt = Trim(CStr(sess.findById(pathQmTxt).Text))');
  MemoAdd('  If Len(outTxt) = 0 Then');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    Err.Clear');
  MemoAdd('    outTxt = Trim(CStr(sess.findById(pathQmTxt).Text))');
  MemoAdd('  End If');
  MemoAdd('  If Len(outTxt) = 0 Then Exit Function');
  MemoAdd('  PreencherQmTxt = True');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function PreencherPernrRT3(ByVal sess, ByVal rowIdx, ByVal pernr, ByVal idExec, ByVal idRT, ByVal linha)');
  MemoAdd('  PreencherPernrRT3 = False');
  MemoAdd('  Dim tbl, pathPernr, outPernr');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set tbl = sess.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/tblSAPLYSCS_INFADREQTRANSTC_RT3")');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  If Not tbl Is Nothing Then');
  MemoAdd('    tbl.VerticalScrollbar.Position = rowIdx');
  MemoAdd('    tbl.CurrentCellRow = rowIdx');
  MemoAdd('    tbl.CurrentCellColumn = "YSCSRTITPAX-PERNR"');
  MemoAdd('  End If');
  MemoAdd('  Err.Clear');
  MemoAdd('  pathPernr = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('              "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-PERNR[3," & rowIdx & "]"');
  MemoAdd('  sess.findById(pathPernr).setFocus');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  sess.findById(pathPernr).Text = Trim(CStr(pernr))');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  sess.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  outPernr = Trim(CStr(sess.findById(pathPernr).Text))');
  MemoAdd('  If Len(outPernr) = 0 Then');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    outPernr = Trim(CStr(sess.findById(pathPernr).Text))');
  MemoAdd('  End If');
  MemoAdd('  If Len(outPernr) = 0 Then Exit Function');
  MemoAdd('  PreencherPernrRT3 = True');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------

  MemoAdd('Function ProcessarRT3(ByVal idExec, ByVal idRT, ByVal linha, ByVal acao, ByVal rtNumero, ByVal tipoRT, _');
  MemoAdd('                     ByVal pernr, ByVal contratoFix, ByVal tipoDoc, ByVal documento, ByVal passageiro, _');
  MemoAdd('                     ByVal qmtxt, ByVal requis, ByVal pessoaContato, ByVal foneContato, _');
  MemoAdd('                     ByVal iwerk, ByVal ingrp, ByVal codTrasg, ByVal classePas, _');
  MemoAdd('                     ByVal origem, ByVal destino, ByVal retorno, ByVal dtIda, ByVal hIda, ByVal dtVolta, ByVal hVolta, _');
  MemoAdd('                     ByVal ccusto, ByVal DiagramaRede, ByVal OperRede, ByVal ElementoPEP, _');
  MemoAdd('                     ByVal cobertura, ByVal Recolhimetno, ByVal salvar)');
  MemoAdd('  ProcessarRT3 = False');
  MemoAdd('  On Error Resume Next');
  //---------------------------------------------------------
  MemoAdd('  Dim st, stType, pCod, outCod, stCod, tyCod, rtCriada, popupMsg, popupMsgUltima, rtExistente');
  //---------------------------------------------------------
  MemoAdd('  rtCriada = ""');
  MemoAdd('  popupMsgUltima = ""');
  MemoAdd('  rtExistente = ""');
  MemoAdd('  If UCase(acao) = "MODIFICAR" Then');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  Else');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw51"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = tipoRT');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  If Not PreencherQmTxt(session, qmtxt, idExec, idRT, linha) Then Exit Function');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  If (Trim(cobertura) = "1") Or (UCase(Trim(cobertura)) = "TRUE") Then');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnANWENDERSTATUS").press');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True');
  MemoAdd('    session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('  Call AbrirTelaTransporte(cobertura)');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtYSCSRTCB-REQUIS").text = requis');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-PESSOA_CONTATO").text = pessoaContato');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-TELEFONE_CONTATO").text = foneContato');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-IWERK").text = iwerk');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-INGRP").text = ingrp');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  If Not PreencherPernrRT3(session, 0, pernr, idExec, idRT, linha) Then Exit Function');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  If Len(Trim(st)) = 0 Then');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('  End If');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If MsgCampoContratoObrigatorio(st) Then');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix, 0) Then');
  MemoAdd('      Call ResetSAP(session)');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('    On Error Resume Next');
  MemoAdd('  End If');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    If InStr(1, st, "não está ativo", vbTextCompare) > 0 Then Exit Function');
  MemoAdd('    If InStr(1, st, "Contrato obrigatório", vbTextCompare) > 0 Then');
  MemoAdd('      On Error GoTo 0');
  MemoAdd('      If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix, 0) Then Exit Function');
  MemoAdd('      On Error Resume Next');
  MemoAdd('    Else');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  End If');

  MemoAdd('  pCod = "wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('         "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-CODTRASG[8,0]"');
  MemoAdd('  session.findById(pCod).setFocus');
  MemoAdd('  session.findById(pCod).text = Trim(CStr(codTrasg))');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  outCod = session.findById(pCod).text');
  MemoAdd('  stCod = GetStatusText(session)');
  MemoAdd('  tyCod = GetStatusType(session)');
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
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-DTVIAGEM[52,0]").text = dtIda');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-HVIAGEM[53,0]").text = hIda');
  MemoAdd('  Call PreencherRateioRT3(0, ccusto, ElementoPEP, DiagramaRede, OperRede)');
  MemoAdd('  If (Trim(Recolhimetno) = "1") Or (UCase(Trim(Recolhimetno)) = "TRUE") Then');
  MemoAdd('    If Not PreencherPernrRT3(session, 1, pernr, idExec, idRT, linha) Then Exit Function');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    If Len(Trim(st)) = 0 Then');
  MemoAdd('      WScript.Sleep 100');
  MemoAdd('      st = GetStatusText(session)');
  MemoAdd('    End If');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If MsgCampoContratoObrigatorio(st) Then');
  MemoAdd('      On Error GoTo 0');
  MemoAdd('      If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix, 1) Then');
  MemoAdd('        Call ResetSAP(session)');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      On Error Resume Next');
  MemoAdd('      WScript.Sleep 100');
  MemoAdd('    End If');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-CODTRASG[8,1]").text = codTrasg');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTITPAX-CLASSEPAS[9,1]").text = classePas');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_ORIGEM[28,1]").text = destino');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-LOCAL_DESTINO[38,1]").text = retorno');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-DTVIAGEM[52,1]").text = dtVolta');
  MemoAdd('    session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0103/" & _');
  MemoAdd('                     "tblSAPLYSCS_INFADREQTRANSTC_RT3/ctxtYSCSRTSITPAX-HVIAGEM[53,1]").text = hVolta');
  MemoAdd('    Call PreencherRateioRT3(1, ccusto, ElementoPEP, DiagramaRede, OperRede)');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If MsgCampoContratoObrigatorio(st) Then');
  MemoAdd('      On Error GoTo 0');
  MemoAdd('      If Not ResolverContratoObrigatorio(idExec, idRT, linha, contratoFix, 1) Then');
  MemoAdd('        Call ResetSAP(session)');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      On Error Resume Next');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('');
  //INICIO SALVAR ProcessarRT3
    MemoAdd('  If salvar Then');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If Len(Trim(popupMsg)) > 0 Then popupMsgUltima = popupMsg');
  MemoAdd('');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        rtExistente = ExtrairNumeroRTConflito(popupMsg)');
  MemoAdd('        If Len(Trim(rtExistente)) > 0 Then');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|Passageiro já cadastrado para mesma data/hora na RT " & rtExistente & ". " & popupMsg');
  MemoAdd('        Else');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|" & popupMsg');
  MemoAdd('        End If');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('');
  MemoAdd('      session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    st = Trim(CStr(GetStatusText(session)))');
  MemoAdd('    If Len(st) = 0 Then st = popupMsgUltima');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('');
  MemoAdd('    If stType = "E" Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_1_SAVE_RT3|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If Len(Trim(popupMsg)) > 0 Then popupMsgUltima = popupMsg');
  MemoAdd('');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        rtExistente = ExtrairNumeroRTConflito(popupMsg)');
  MemoAdd('        If Len(Trim(rtExistente)) > 0 Then');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|Passageiro já cadastrado para mesma data/hora na RT " & rtExistente & ". " & popupMsg');
  MemoAdd('        Else');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|" & popupMsg');
  MemoAdd('        End If');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('');
  MemoAdd('      session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    rtCriada = CapturarRTCriadaRapido(session)');
  MemoAdd('    If Len(Trim(rtCriada)) > 0 Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|RT_CRIADA|" & rtCriada');
  MemoAdd('    Else');
  MemoAdd('      st = Trim(CStr(GetStatusText(session)))');
  MemoAdd('      If Len(st) = 0 Then st = popupMsgUltima');
  MemoAdd('');
  MemoAdd('      If MsgConflitoRTExistente(st) Then');
  MemoAdd('        rtExistente = ExtrairNumeroRTConflito(st)');
  MemoAdd('        If Len(Trim(rtExistente)) > 0 Then');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|Passageiro já cadastrado para mesma data/hora na RT " & rtExistente & ". " & st');
  MemoAdd('        Else');
  MemoAdd('          LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|" & st');
  MemoAdd('        End If');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('');
  MemoAdd('      If Len(Trim(st)) = 0 Then');
  MemoAdd('        st = "RT não confirmada no SAP"');
  MemoAdd('      End If');
  MemoAdd('');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|" & st');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  //FIM SALVAR ProcessarRT3-----------------------------------------
  MemoAdd('');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('    FechaPopupSeExistir session');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  LogLine idExec & "|" & idRT & "|" & linha & "|OK|" & st');
  MemoAdd('  ProcessarRT3 = True');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  //FIM ProcessarRT3-----------------------------------------
  MemoAdd('Function TrySetText(ByVal sess, ByVal id, ByVal value)');
  MemoAdd('  TrySetText = False');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  sess.findById(id).setFocus');
  MemoAdd('  sess.findById(id).Text = value');
  MemoAdd('  sess.findById("wnd[0]").sendVKey 0');
  MemoAdd('  If Err.Number = 0 Then TrySetText = True');
  MemoAdd('  Err.Clear');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('Function TrySetTextByIds(ByVal sess, ByVal ids, ByVal value)');
  MemoAdd('  TrySetTextByIds = False');
  MemoAdd('  Dim i');
  MemoAdd('  For i = 0 To UBound(ids)');
  MemoAdd('    If TrySetText(sess, CStr(ids(i)), value) Then');
  MemoAdd('      TrySetTextByIds = True');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  Next');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('Sub PreencherRateioAprovacaoRT7(ByVal idExec, ByVal idRT, ByVal linha, ByVal ccusto, ByVal ElementoPEP, ByVal DiagramaRede, ByVal OperRede)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim c, pep, r, o, ok, ids');
  MemoAdd('  c   = Trim(CStr(ccusto))');
  MemoAdd('  pep = Trim(CStr(ElementoPEP))');
  MemoAdd('  r   = Trim(CStr(DiagramaRede))');
  MemoAdd('  o   = Trim(CStr(OperRede))');
  MemoAdd('  ok = False');
  MemoAdd('  If Len(c) > 0 Then');
  MemoAdd('    ok = TrySetText(session, "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-CCUSTO[3,0]", c)');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('  If Len(pep) > 0 Then');
  MemoAdd('    ids = Array( _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-ELEM_PEP[4,0]", _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-PEP[4,0]", _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-ELEM_PEP[3,0]", _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-PEP[3,0]" _');
  MemoAdd('    )');
  MemoAdd('    ok = TrySetTextByIds(session, ids, pep)');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('  If (Len(r) > 0) And (Len(o) > 0) Then');
  MemoAdd('    ids = Array( _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-REDE[4,0]", _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-REDE[3,0]" _');
  MemoAdd('    )');
  MemoAdd('    ok = TrySetTextByIds(session, ids, r)');
  MemoAdd('    ids = Array( _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-OPER_REDE[5,0]", _');
  MemoAdd('      "wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/ctxtYSCSOBJAPRO-OPER_REDE[4,0]" _');
  MemoAdd('    )');
  MemoAdd('    ok = ok And TrySetTextByIds(session, ids, o)');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');
  //-----------------------------------------------------------------------



  MemoAdd('Function ProcessarRT7(ByVal idExec, ByVal idRT, ByVal linha, ByVal acao, ByVal rtNumero, ByVal tipoRT, _');
  MemoAdd('                     ByVal pernr, ByVal contratoFix, ByVal tipoDoc, ByVal documento, ByVal passageiro, _');
  MemoAdd('                     ByVal qmtxt, ByVal requis, ByVal pessoaContato, ByVal foneContato, _');
  MemoAdd('                     ByVal iwerk, ByVal ingrp, ByVal codTrasg, ByVal classePas, _');
  MemoAdd('                     ByVal origem, ByVal destino, ByVal retorno, ByVal dtIda, ByVal hIda, ByVal dtVolta, ByVal hVolta, _');
  MemoAdd('                     ByVal ccusto, ByVal DiagramaRede, ByVal OperRede, ByVal ElementoPEP, _');
  MemoAdd('                     ByVal cobertura, ByVal Recolhimetno, ByVal salvar)');
  MemoAdd('  ProcessarRT7 = False');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim st, stType, hIda2, hVolta2, rtCriada, temRecolhimento, popupMsg');
  MemoAdd('  rtCriada = ""');
  MemoAdd('  popupMsg = ""');
  MemoAdd('  hIda2 = Trim(CStr(hIda))');
  MemoAdd('  hVolta2 = Trim(CStr(hVolta))');
  MemoAdd('  If Len(hIda2) = 5 Then hIda2 = hIda2 & ":00"');
  MemoAdd('  If Len(hVolta2) = 5 Then hVolta2 = hVolta2 & ":00"');
  MemoAdd('  temRecolhimento = ((Trim(CStr(Recolhimetno)) = "1") Or (UCase(Trim(CStr(Recolhimetno))) = "TRUE"))');
  MemoAdd('');
  MemoAdd('  If UCase(acao) = "MODIFICAR" Then');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw52"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = rtNumero');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  Else');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/okcd").text = "/niw51"');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[0]/usr/ctxtRIWO00-QMART").text = tipoRT');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('  If Not PreencherQmTxt(session, qmtxt, idExec, idRT, linha) Then Exit Function');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  If (Trim(cobertura) = "1") Or (UCase(Trim(cobertura)) = "TRUE") Then');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    session.findById("wnd[0]/usr/subSCREEN_1:SAPLIQS0:1060/btnANWENDERSTATUS").press');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    session.findById("wnd[1]/usr/tblSAPLBSVATC_E/radJ_STMAINT-ANWS[0,2]").selected = True');
  MemoAdd('    session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    session.findById("wnd[0]").sendVKey 0');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  Call AbrirTelaTransporte(cobertura)');
  MemoAdd('');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtYSCSRTCB-REQUIS").text = requis');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-PESSOA_CONTATO").text = pessoaContato');
  MemoAdd('  session.findById("wnd[0]/usr/txtYSCSRTCB-TELEFONE_CONTATO").text = foneContato');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-IWERK").text = iwerk');
  MemoAdd('  session.findById("wnd[0]/usr/ctxtVIQMEL-INGRP").text = ingrp');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  '' ===== Linha 0 - Ida =====');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/radWA_BOTOES-CAT_PAX_EXT[4,0]").selected = True');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/cmbYSCSRTITPAX-TIPO_DOC[6,0]").key = tipoDoc');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/txtYSCSRTITPAX-DOCUMENTO[7,0]").text = documento');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/txtYSCSRTITPAX-PASSAGEIRO[8,0]").text = passageiro');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTSITPAX-CODTRASG[16,0]").text = codTrasg');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTITPAX-CLASSEPAS[17,0]").text = classePas');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTSITPAX-LOCAL_ORIGEM[36,0]").text = origem');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTSITPAX-LOCAL_DESTINO[46,0]").text = destino');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTSITPAX-DTVIAGEM[59,0]").text = dtIda');
  MemoAdd('  session.findById("wnd[0]/usr/subSUBTELA_001:SAPLYSCS_INFADREQTRANS:0107/" & _');
  MemoAdd('                   "tblSAPLYSCS_INFADREQTRANSTC_RT7/ctxtYSCSRTSITPAX-HVIAGEM[60,0]").text = hIda2');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  If session.Children.Count > 1 Then');
  MemoAdd('    session.findById("wnd[1]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_LINHA_IDA_RT7|" & st');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  '' ===== Aprovação / Rateio =====');
  MemoAdd('  session.findById("wnd[0]/tbar[1]/btn[6]").press');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  session.findById("wnd[0]/usr/tblSAPLYSCS_INFADREQTRANSTC_OBJ_APR_RT/txtYSCSOBJAPRO-PERCENTUAL[2,0]").text = "100"');
  MemoAdd('  Call PreencherRateioAprovacaoRT7(idExec, idRT, linha, ccusto, ElementoPEP, DiagramaRede, OperRede)');
  MemoAdd('  session.findById("wnd[0]").sendVKey 0');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('');
  MemoAdd('  If session.Children.Count > 1 Then');
  MemoAdd('    session.findById("wnd[1]").sendVKey 0');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_RATEIO_RT7|" & st');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');




    MemoAdd('  If salvar Then');
  MemoAdd('    '' ===== Save intermediário =====');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        LogLine idExec & "|" & idRT & "|" & linha & "|CONFLITO_RT_EXISTENTE|" & popupMsg');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      session.findById("wnd[1]").sendVKey 0');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If stType = "E" Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_SAVE_INTERMEDIARIO_RT7|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    If MensagemIndicaFalhaNegocio(st) Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    '' ===== Save final 1 =====');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        LogLine idExec & "|" & idRT & "|" & linha & "|CONFLITO_RT_EXISTENTE|" & popupMsg');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      session.findById("wnd[1]").sendVKey 0');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If stType = "E" Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_SAVE_FINAL_1_RT7|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    If MensagemIndicaFalhaNegocio(st) Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    '' ===== Save final 2 =====');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 150');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        LogLine idExec & "|" & idRT & "|" & linha & "|CONFLITO_RT_EXISTENTE|" & popupMsg');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      session.findById("wnd[1]").sendVKey 0');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If stType = "E" Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_SAVE_FINAL_2_RT7|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    If MensagemIndicaFalhaNegocio(st) Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    '' ===== Save final 3 - necessário no R7 =====');
  MemoAdd('    session.findById("wnd[0]/tbar[0]/btn[11]").press');
  MemoAdd('    WScript.Sleep 200');
  MemoAdd('');
  MemoAdd('    If session.Children.Count > 1 Then');
  MemoAdd('      popupMsg = GetPopupMessageText(session)');
  MemoAdd('      If MsgConflitoRTExistente(popupMsg) Then');
  MemoAdd('        LogLine idExec & "|" & idRT & "|" & linha & "|CONFLITO_RT_EXISTENTE|" & popupMsg');
  MemoAdd('        session.findById("wnd[1]/tbar[0]/btn[0]").press');
  MemoAdd('        WScript.Sleep 100');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      session.findById("wnd[1]").sendVKey 0');
  MemoAdd('      WScript.Sleep 150');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    st = GetStatusText(session)');
  MemoAdd('    stType = GetStatusType(session)');
  MemoAdd('    If stType = "E" Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO_APOS_SAVE_FINAL_3_RT7|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    If MensagemIndicaFalhaNegocio(st) Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('      FechaPopupSeExistir session');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    rtCriada = CapturarRTCriadaRapido(session)');
  MemoAdd('    If Len(Trim(rtCriada)) > 0 Then');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|RT_CRIADA|" & rtCriada');
  MemoAdd('    Else');
  MemoAdd('      st = GetStatusText(session)');
  MemoAdd('      If MsgConflitoRTExistente(st) Then');
  MemoAdd('        LogLine idExec & "|" & idRT & "|" & linha & "|CONFLITO_RT_EXISTENTE|" & st');
  MemoAdd('        Exit Function');
  MemoAdd('      End If');
  MemoAdd('      LogLine idExec & "|" & idRT & "|" & linha & "|RT_NAO_CONFIRMADA|" & st');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('  End If');

  MemoAdd('');
  MemoAdd('  st = GetStatusText(session)');
  MemoAdd('  stType = GetStatusType(session)');
  MemoAdd('  If stType = "E" Then');
  MemoAdd('    LogLine idExec & "|" & idRT & "|" & linha & "|ERRO|" & st');
  MemoAdd('    FechaPopupSeExistir session');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  LogLine idExec & "|" & idRT & "|" & linha & "|OK|" & st');
  MemoAdd('  ProcessarRT7 = True');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');
  //-----------------------------------------------------------------------
  MemoAdd('Function ProcessarRT(ByVal idExec, ByVal idRT, ByVal linha, ByVal acao, ByVal rtNumero, ByVal tipoRT, _');
  MemoAdd('                    ByVal pernr, ByVal contratoFix, ByVal tipoDoc, ByVal documento, ByVal passageiro, _');
  MemoAdd('                    ByVal qmtxt, ByVal requis, ByVal pessoaContato, ByVal foneContato, _');
  MemoAdd('                    ByVal iwerk, ByVal ingrp, ByVal codTrasg, ByVal classePas, _');
  MemoAdd('                    ByVal origem, ByVal destino, ByVal retorno, ByVal dtIda, ByVal hIda, ByVal dtVolta, ByVal hVolta, _');
  MemoAdd('                    ByVal ccusto, ByVal DiagramaRede, ByVal OperRede, ByVal ElementoPEP, _');
  MemoAdd('                    ByVal cobertura, ByVal Recolhimetno, ByVal salvar)');
  MemoAdd('  If LCase(Trim(tipoRT)) = "r7" Then');
  MemoAdd('    ProcessarRT = ProcessarRT7(idExec,idRT,linha,acao,rtNumero,tipoRT,pernr,contratoFix,tipoDoc,documento,passageiro, _');
  MemoAdd('                              qmtxt,requis,pessoaContato,foneContato,iwerk,ingrp,codTrasg,classePas,origem,destino,retorno,dtIda,hIda,dtVolta,hVolta, _');
  MemoAdd('                              ccusto,DiagramaRede,OperRede,ElementoPEP,cobertura,Recolhimetno,salvar)');
  MemoAdd('  Else');
  MemoAdd('    ProcessarRT = ProcessarRT3(idExec,idRT,linha,acao,rtNumero,tipoRT,pernr,contratoFix,tipoDoc,documento,passageiro, _');
  MemoAdd('                              qmtxt,requis,pessoaContato,foneContato,iwerk,ingrp,codTrasg,classePas,origem,destino,retorno,dtIda,hIda,dtVolta,hVolta, _');
  MemoAdd('                              ccusto,DiagramaRede,OperRede,ElementoPEP,cobertura,Recolhimetno,salvar)');
  MemoAdd('  End If');
  MemoAdd('End Function');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('LogLine "INICIO_LOTE"');
  MemoAdd('LogLine "LOG_USADO|" & LOGFILE');
  MemoAdd('');
  //-------------------------------------------------------
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
  MemoAdd('On Error Resume Next');
  MemoAdd('session.findById("wnd[0]").maximize');
  MemoAdd('On Error GoTo 0');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('Dim itens, i, okCount, errCount, salvar');

  if FrmDataModule.ADOQueryConfigRT.FieldByName('RT_Salvar').AsBoolean then
    MemoAdd('salvar = True')
  else
    MemoAdd('salvar = False');
  MemoAdd('');
  //-------------------------------------------------------
  MemoAdd('itens = Array( _');
  for i := 0 to ItemCount - 1 do
  begin
    if i < ItemCount - 1 then
      MemoAdd(BuildItemLine(Itens[i]) + ', _')
    else
      MemoAdd(BuildItemLine(Itens[i]) + ')');
  end;
  //-------------------------------------------------------
  MemoAdd('');
  MemoAdd('okCount = 0');
  MemoAdd('errCount = 0');
  MemoAdd('');
  MemoAdd('For i = 0 To UBound(itens)');
  MemoAdd('  Dim a, n');
  MemoAdd('  a = itens(i)');
  MemoAdd('  n = UBound(a)');
  MemoAdd('  If n < 31 Then');
  MemoAdd('    LogLine "ERRO_ITEM_TAMANHO|" & i & "|UBOUND=" & n');
  MemoAdd('    errCount = errCount + 1');
  MemoAdd('  Else');
  MemoAdd('    If SafeRunItem(a, salvar) Then');
  MemoAdd('      okCount = okCount + 1');
  MemoAdd('    Else');
  MemoAdd('      errCount = errCount + 1');
  MemoAdd('    End If');
  MemoAdd('  End If');
  MemoAdd('  WScript.Sleep 100');
  MemoAdd('Next');
  MemoAdd('');
  MemoAdd('LogLine "FIM_LOTE|OK=" & okCount & "|ERRO=" & errCount');

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
      MessageBox(0,
        PChar('Processamento SAP concluído!' + sLineBreak +
              'Itens enviados ao SAP: ' + IntToStr(ItemCount) + sLineBreak +
              'Registros prontos analisados: ' + IntToStr(QtProntos) + sLineBreak +
              'Registros bloqueados: ' + IntToStr(QtBloqueados)),
        'Colibri',
        MB_ICONINFORMATION);
    end
    else
      MessageBox(0,PChar('Script terminou, mas não encontrei o log: ' + EnderecoLog),
        'Colibri',MB_ICONWARNING);
  end
  else
    MessageBox(0,PChar('Erro ao iniciar o script VBS.'),
        'Colibri',MB_ICONERROR);
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
    Q.Connection := FrmDataModule.ADOConnectionRT;

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
    Q.FieldByName('RecolherPara').AsString    := Trim(Dados.Retorno);

    Q.FieldByName('DataIda').AsDateTime    := FrmPrincipal.DataSAP_ToDate(Dados.DataIda);

    // ✅ ida/volta + flag
    Q.FieldByName('RT_HoraIda').AsString      := Trim(Dados.HoraIda);
    Q.FieldByName('RT_HoraVolta').AsString       := Trim(Dados.HoraVolta);
    Q.FieldByName('booleanRecolhimento').AsBoolean := Dados.booleanRecolhimento;

    Q.FieldByName('RT_CentroCusto').AsString  := Trim(Dados.CentroCusto);
    Q.FieldByName('RT_DiagramaRede').AsString  := Trim(Dados.DiagramaRede);
    Q.FieldByName('RT_OperRede').AsString  := Trim(Dados.OperRede);
    Q.FieldByName('RT_ElementoPEP').AsString  := Trim(Dados.ElementoPEP);
    Q.FieldByName('RT_Cobertura').AsBoolean   := Dados.RTCobertura;

    Q.FieldByName('ChavePassageiro').AsString :=
    ChavePassageiro(
      Dados.MatriculaPax,
      IfThen(Trim(Dados.CPF) <> '', 'C',
           IfThen(Trim(Dados.Passaporte) <> '', 'P', '')),
      IfThen(Trim(Dados.CPF) <> '', Dados.CPF, Dados.Passaporte)
      );

    Q.FieldByName('ChaveIda').AsString := ChaveRTIda(Dados);
    Q.FieldByName('ChaveVolta').AsString := ChaveRTVolta(Dados);
    Q.FieldByName('ChaveCompleta').AsString := ChaveRTCompleta(Dados);

    Q.FieldByName('idProgramacaoExecutante').AsInteger := Dados.idProgramacaoExecutante;

    if Trim(Dados.DataVolta) <> '' then
      Q.FieldByName('DataVolta').AsDateTime := FrmPrincipal.DataSAP_ToDate(Dados.DataVolta)
    else
      Q.FieldByName('DataVolta').Clear;

    //-------------------------------------------------------
    if Q.FindField('RT_Status') <> nil then
    begin
      if Trim(Dados.RTNumero) <> '' then
        Q.FieldByName('RT_Status').AsString := RT_STATUS_EMITIDA
      else
        Q.FieldByName('RT_Status').AsString := RT_STATUS_PRONTO_EMITIR;
    end;

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
        QID.Connection := FrmDataModule.ADOConnectionRT;
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
  const CodigoSAP, CPF, Passaporte: String): TCodicoCusto;
var
  Q: TADOQuery;
begin
  Result := Default(TCodicoCusto);

  Q := FrmDataModule.ADOQueryTemporarioDBConsulta1;
  Q.Close;
  Q.SQL.Clear;
  Q.Parameters.Clear;

  if Trim(CodigoSAP) <> '' then
  begin
    Q.SQL.Text :=
      'SELECT CentroCusto, DiagramaRede, OperRede, ElementoPEP '+
      'FROM tblExecutante '+
      'WHERE CodigoSAP = :CodigoSAP';
    Q.Parameters.ParamByName('CodigoSAP').Value := Trim(CodigoSAP);
  end
  else if Trim(CPF) <> '' then
  begin
    Q.SQL.Text :=
      'SELECT CentroCusto, DiagramaRede, OperRede, ElementoPEP '+
      'FROM tblExecutante '+
      'WHERE CPF = :CPF OR CPF = :CPF2';
    Q.Parameters.ParamByName('CPF').Value  := Trim(CPF);
    Q.Parameters.ParamByName('CPF2').Value := FrmPrincipal.FormatarCPF(FrmPrincipal.SomenteNumero(CPF));
  end
  else if Trim(Passaporte) <> '' then
  begin
    Q.SQL.Text :=
      'SELECT CentroCusto, DiagramaRede, OperRede, ElementoPEP '+
      'FROM tblExecutante '+
      'WHERE OutroDocumento = :Passaporte';
    Q.Parameters.ParamByName('Passaporte').Value := Trim(Passaporte);
  end
  else
    Exit; // sem chave de busca

  Q.Open;

  if not Q.Eof then
  begin
    Result.CentroCusto  := Q.FieldByName('CentroCusto').AsString;
    Result.DiagramaRede := Q.FieldByName('DiagramaRede').AsString;
    Result.OperRede     := Q.FieldByName('OperRede').AsString;
    Result.ElementoPEP  := Q.FieldByName('ElementoPEP').AsString;
  end;
end;

function TFrmConsultaExecutantesProgramados.StatusUnicoPorEventoSAP(
  const AEvento, AValor, ARTNumero: string): string;
var
  EventoUp, ValorUp, RTNum: string;
begin
  EventoUp := UpperCase(Trim(AEvento));
  ValorUp := UpperCase(Trim(NormalizarMensagemLog(AValor)));
  RTNum := Trim(ARTNumero);

  if EventoUp = 'RT_CRIADA' then
    Exit(RT_STATUS_EMITIDA);

  if EventoUp = 'JA_EXISTE' then
    Exit(RT_STATUS_JA_EXISTE);

  if EventoUp = 'CONFLITO_RT_EXISTENTE' then
    Exit(RT_STATUS_ERRO_CONFLITO_RT);

  if EventoUp = 'RT_NAO_CONFIRMADA' then
  begin
    if Pos('NAO ESTA ATIVO', ValorUp) > 0 then
      Exit(RT_STATUS_ERRO_NAO_ATIVO)
    else
      Exit(RT_STATUS_ERRO_EMISSAO);
  end;

  if EventoUp = 'ERRO' then
    Exit(RT_STATUS_ERRO_EMISSAO);

  if (RTNum <> '') and (RTNum <> 'SEM RT') then
    Exit(RT_STATUS_EMITIDA);

  Result := RT_STATUS_ERRO_EMISSAO;
end;


procedure TFrmConsultaExecutantesProgramados.ProcessarRetornoSAP(const LinhaLog: string);
var
  Partes: TArray<string>;
  IdExec, IdRT: Integer;
  Evento, Valor, NumeroRT, StatusUnico, ValorNorm: string;
begin
  if Trim(LinhaLog) = '' then
    Exit;

  if Pos('|INICIO_LOTE', UpperCase(LinhaLog)) > 0 then
    Exit;

  if Pos('|FIM_LOTE', UpperCase(LinhaLog)) > 0 then
    Exit;

  if Pos('|LOG_USADO|', UpperCase(LinhaLog)) > 0 then
    Exit;

  Partes := LinhaLog.Split(['|']);

  if Length(Partes) < 6 then
    Exit;

  IdExec := StrToIntDef(Trim(Partes[1]), 0);
  IdRT := StrToIntDef(Trim(Partes[2]), 0);
  Evento := UpperCase(Trim(Partes[4]));
  Valor := Trim(Partes[5]);
  ValorNorm := NormalizarMensagemLog(Valor);

  NumeroRT := '';

  if Evento = 'RT_CRIADA' then
  begin
    NumeroRT := Trim(Valor);

    AtualizarRetornoTabelaRT(
      IdRT,
      MSG_RT_EMITIDA,
      NumeroRT,
      RT_STATUS_EMITIDA
    );

    AtualizarRetornoExecutante(
      IdExec,
      MSG_RT_EMITIDA,
      NumeroRT,
      RT_STATUS_EMITIDA
    );
    Exit;
  end;

  if Evento = 'JA_EXISTE' then
  begin
    AtualizarRetornoTabelaRT(
      IdRT,
      ValorNorm,
      '',
      RT_STATUS_JA_EXISTE
    );

    AtualizarRetornoExecutante(
      IdExec,
      ValorNorm,
      '',
      RT_STATUS_JA_EXISTE
    );
    Exit;
  end;

  if Evento = 'CONFLITO_RT_EXISTENTE' then
  begin
    AtualizarRetornoTabelaRT(
      IdRT,
      ValorNorm,
      '',
      RT_STATUS_ERRO_CONFLITO_RT
    );

    AtualizarRetornoExecutante(
      IdExec,
      ValorNorm,
      '',
      RT_STATUS_ERRO_CONFLITO_RT
    );
    Exit;
  end;

  if Evento = 'RT_NAO_CONFIRMADA' then
  begin
    StatusUnico := StatusUnicoPorEventoSAP(Evento, ValorNorm, '');

    AtualizarRetornoTabelaRT(
      IdRT,
      ValorNorm,
      '',
      StatusUnico
    );

    AtualizarRetornoExecutante(
      IdExec,
      ValorNorm,
      '',
      StatusUnico
    );
    Exit;
  end;

  if Evento = 'ERRO' then
  begin
    AtualizarRetornoTabelaRT(
      IdRT,
      ValorNorm,
      '',
      RT_STATUS_ERRO_EMISSAO
    );

    AtualizarRetornoExecutante(
      IdExec,
      ValorNorm,
      '',
      RT_STATUS_ERRO_EMISSAO
    );
    Exit;
  end;

  if Evento = 'OK' then
  begin
    if Trim(ValorNorm) <> '' then
      AtualizarMensagemRetornoSemRebaixarStatus(IdExec, IdRT, ValorNorm);

    Exit;
  end;
end;

function TFrmConsultaExecutantesProgramados.RTLocalIndicaCancelamentoReal(
  const AStatus: string;
  const ARTCancelada: Boolean): Boolean;
var
  SProc: string;
begin
  SProc := UpperCase(Trim(AStatus));

  Result :=
    ARTCancelada or
    (SProc = 'CANCELADA') or
    (SProc = 'CANCELAMENTO_ERRO');
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoExecutante(
  const idProgramacaoExecutante: Integer;
  const MensagemRT, RT_Numero, RT_Status, RT_HoraIda, RT_HoraVolta, RT_Modal,
  RT_Classe, RecolherPara: string; booleanRecolhimento: Boolean;
  DataVolta: TDateTime);
var
  Q: TADOQuery;
  MensagemSafe, StatusSafe,
  HoraIdaSafe, HoraVoltaSafe, ModalSafe, ClasseSafe,
  RecolherParaSafe, RTSafe: string;
  StatusSeguro: string;
begin
  if idProgramacaoExecutante <= 0 then
    Exit;

  StatusSeguro := Trim(RT_Status);

  // tabela soberana não deve ficar como CANCELADA
  if SameText(StatusSeguro, RT_STATUS_CANCELADA) then
    StatusSeguro := RT_STATUS_JA_EXISTE;

  if Trim(StatusSeguro) = '' then
    StatusSeguro := RT_STATUS_PENDENTE;

  StatusSeguro := StatusRTNormalizado(StatusSeguro);

  if Trim(StatusSeguro) = '' then
    StatusSeguro := RT_STATUS_PENDENTE;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;

    MensagemSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_Mensagem',
      MensagemRT,
      100
    );

    RTSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT',
      Trim(RT_Numero),
      15
    );

    StatusSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_Status',
      StatusSeguro,
      40
    );

    HoraIdaSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_HoraIda',
      RT_HoraIda,
      5
    );

    HoraVoltaSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_HoraVolta',
      RT_HoraVolta,
      5
    );

    ModalSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_Modal',
      RT_Modal,
      5
    );

    ClasseSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RT_Classe',
      RT_Classe,
      5
    );

    RecolherParaSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionColibri,
      'tblProgramacaoExecutante',
      'RecolherPara',
      RecolherPara,
      20
    );

    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  RT_Mensagem = :RT_Mensagem, ' +
      '  [RT] = :RT, ' +
      '  RT_Status = :RT_Status, ' +
      '  RT_HoraIda = :RT_HoraIda, ' +
      '  RT_HoraVolta = :RT_HoraVolta, ' +
      '  DataVolta = :DataVolta, ' +
      '  booleanRecolhimento = :booleanRecolhimento, ' +
      '  RT_Modal = :RT_Modal, ' +
      '  RT_Classe = :RT_Classe, ' +
      '  RecolherPara = :RecolherPara ' +
      'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemSafe;
    Q.Parameters.ParamByName('RT').Value := RTSafe;
    Q.Parameters.ParamByName('RT_Status').Value := StatusSafe;
    Q.Parameters.ParamByName('RT_HoraIda').Value := HoraIdaSafe;
    Q.Parameters.ParamByName('RT_HoraVolta').Value := HoraVoltaSafe;

    if DataVolta > 0 then
      Q.Parameters.ParamByName('DataVolta').Value := DataVolta
    else
      Q.Parameters.ParamByName('DataVolta').Value := Null;

    Q.Parameters.ParamByName('booleanRecolhimento').Value := booleanRecolhimento;
    Q.Parameters.ParamByName('RT_Modal').Value := ModalSafe;
    Q.Parameters.ParamByName('RT_Classe').Value := ClasseSafe;
    Q.Parameters.ParamByName('RecolherPara').Value := RecolherParaSafe;
    Q.Parameters.ParamByName('idProgramacaoExecutante').Value := idProgramacaoExecutante;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actGerarMultiplasRTsExecute(
  Sender: TObject);
var
  HoraComecou, HoraTerminou: TDateTime;
  DuracaoMinutos: Integer;
begin
  if Application.MessageBox(PChar(
  'Deseja realmente criar as RTs no SAP?'),
  '.::ATENÇÃO::.',36) = 6 then
  begin
    HoraComecou := Now;
    GerarMultiplasRTsArray;

    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active := False;
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active := True;
    FrmDataModule.ADOQueryGestaoRT.Active := False;
    FrmDataModule.ADOQueryGestaoRT.Active := True;

    HoraTerminou := Now;
    DuracaoMinutos := MinutesBetween(HoraTerminou, HoraComecou);

    MessageBox(0,PChar(
      'Processamento concluído.' + sLineBreak +
      'Início: ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', HoraComecou) + sLineBreak +
      'Término: ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', HoraTerminou) + sLineBreak +
      'Duração: ' + IntToStr(DuracaoMinutos) + ' minutos'),
      'Tempo de Processamento',MB_ICONINFORMATION);
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarRTSapImportPorChave(
  const Dados: TDadosRT;
  out ADadosSAP: TRTSapImportDados): Boolean;
var
  Q: TADOQuery;

  function BuscarPorCampo(const ACampo, AValor: string): Boolean;
  begin
    Result := False;

    if Trim(AValor) = '' then
      Exit;

    Q.Close;
    Q.SQL.Clear;
    Q.SQL.Text :=
      'SELECT TOP 1 * ' +
      'FROM tblRTSapImport ' +
      'WHERE ( ([' + ACampo + '] = :Valor) ' +
      '        OR ([' + ACampo + '] IS NOT NULL AND Trim([' + ACampo + ']) = :Valor) ) ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      '  AND ([QMNUM] IS NOT NULL AND Trim([QMNUM]) <> '''') ' +
      '  AND ([StatusItem] IS NULL OR Trim([StatusItem]) <> ''09'') ' +
      'ORDER BY DataImportacao DESC, idRTSapImport DESC';

    Q.Parameters.ParamByName('Valor').Value := Trim(AValor);
    Q.Open;

    Result := not Q.IsEmpty;
  end;

begin
  Result := False;
  ADadosSAP := Default(TRTSapImportDados);

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;

    if BuscarPorCampo('ChaveCompleta', Dados.ChaveCompleta) or
       BuscarPorCampo('ChaveIda', Dados.ChaveIda) or
       ((Trim(Dados.ChaveVolta) <> '') and BuscarPorCampo('ChaveVolta', Dados.ChaveVolta)) then
    begin
      ADadosSAP.idRTSapImport := Q.FieldByName('idRTSapImport').AsInteger;
      ADadosSAP.DataImportacao := Q.FieldByName('DataImportacao').AsDateTime;
      ADadosSAP.OrigemConsulta := Trim(Q.FieldByName('OrigemConsulta').AsString);

      if not VarIsNull(Q.FieldByName('PeriodoInicio').Value) then
        ADadosSAP.PeriodoInicio := Q.FieldByName('PeriodoInicio').AsDateTime;

      if not VarIsNull(Q.FieldByName('PeriodoFim').Value) then
        ADadosSAP.PeriodoFim := Q.FieldByName('PeriodoFim').AsDateTime;

      ADadosSAP.QMNUM := Trim(Q.FieldByName('QMNUM').AsString);
      ADadosSAP.QMART := Trim(Q.FieldByName('QMART').AsString);
      ADadosSAP.IWERK := Trim(Q.FieldByName('IWERK').AsString);
      ADadosSAP.INGRP := Trim(Q.FieldByName('INGRP').AsString);
      ADadosSAP.Origem := Trim(Q.FieldByName('Origem').AsString);
      ADadosSAP.Destino := Trim(Q.FieldByName('txtDestino').AsString);

      if not VarIsNull(Q.FieldByName('DataViagem').Value) then
        ADadosSAP.DataViagem := Q.FieldByName('DataViagem').AsDateTime;

      ADadosSAP.HoraViagem := Trim(Q.FieldByName('HoraViagem').AsString);
      ADadosSAP.PERNR := Trim(Q.FieldByName('PERNR').AsString);
      ADadosSAP.TipoDoc := Trim(Q.FieldByName('TipoDoc').AsString);
      ADadosSAP.Documento := Trim(Q.FieldByName('Documento').AsString);
      ADadosSAP.Passageiro := Trim(Q.FieldByName('Passageiro').AsString);
      ADadosSAP.QMTXT := Trim(Q.FieldByName('QMTXT').AsString);
      ADadosSAP.RT_Modal := Trim(Q.FieldByName('RT_Modal').AsString);
      ADadosSAP.RT_Classe := Trim(Q.FieldByName('RT_Classe').AsString);
      ADadosSAP.StatusItem := Trim(Q.FieldByName('StatusItem').AsString);
      ADadosSAP.StatusDescricao := Trim(Q.FieldByName('StatusDescricao').AsString);

      ADadosSAP.idProgramacaoRT := StrToIntDef(Trim(Q.FieldByName('idProgramacaoRT').AsString), 0);
      ADadosSAP.idProgramacaoExecutante := StrToIntDef(Trim(Q.FieldByName('idProgramacaoExecutante').AsString), 0);
      ADadosSAP.ImportadoColibri := Q.FieldByName('ImportadoColibri').AsBoolean;
      ADadosSAP.Observacao := Trim(Q.FieldByName('Observacao').AsString);

      Result := True;
    end;
  finally
    Q.Free;
  end;
end;


procedure TFrmConsultaExecutantesProgramados.actPrepararDadosRTExecute(
  Sender: TObject);
var
  Resp: Integer;
  ForcarTudo: Boolean;
  ModoPreparacao: string;
begin
  Resp := Application.MessageBox(
    PChar(
      'Como deseja preparar a programação para criação das RTs?' + sLineBreak + sLineBreak +
      'SIM = Preparação incremental' + sLineBreak +
      '     Reprocessa somente registros pendentes, com erro ou inconsistentes.' + sLineBreak + sLineBreak +
      'NÃO = Preparação total' + sLineBreak +
      '     Reprocessa todos os registros do período.' + sLineBreak + sLineBreak +
      'CANCELAR = Não executar.'
    ),
    '.::PREPARAÇÃO DE RT::.',
    MB_YESNOCANCEL or MB_ICONQUESTION
  );

  case Resp of
    IDYES:
      begin
        ForcarTudo := False;
        ModoPreparacao := 'incremental';
      end;

    IDNO:
      begin
        ForcarTudo := True;
        ModoPreparacao := 'total';
      end;

  else
    Exit;
  end;

  Screen.Cursor := crSQLWait;
  try
    actProcurarProgramacaoExecutante.Execute;
    actProcurarProgramacaoRT.Execute;
    actProcurarSAPImport.Execute;

    ReclassificarExecutantesParaR3NoPeriodo(
      dataInicio.Date,
      dataFim.Date,
      ForcarTudo
    );

    actProcurarProgramacaoExecutante.Execute;

    PrepararProgramacaoExecutanteParaRT(
      dataInicio.Date,
      dataFim.Date,
      ForcarTudo
    );

    actProcurarProgramacaoExecutante.Execute;
    actProcurarProgramacaoRT.Execute;
    actProcurarSAPImport.Execute;
  finally
    Screen.Cursor := crDefault;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ReclassificarExecutantesParaR3NoPeriodo(
  const ADataIni, ADataFim: TDateTime; const AForcarTudo: Boolean);
var
  QLeitura, QUpdate: TADOQuery;
  DtIni, DtFimMais1: TDateTime;
  IdExec: Integer;
  CodigoSAPAtual, CodigoSAPNormalizado, CodigoSAPFinal, TipoRTFinal: string;
begin
  DtIni := Trunc(ADataIni);
  DtFimMais1 := Trunc(ADataFim) + 1;

  QLeitura := TADOQuery.Create(nil);
  QUpdate := TADOQuery.Create(nil);
  try
    QLeitura.Connection := FrmDataModule.ADOConnectionColibri;
    QLeitura.CursorType := ctStatic;
    QLeitura.LockType := ltReadOnly;
    QLeitura.ParamCheck := True;

    QLeitura.SQL.Text :=
      'SELECT pd.DataProgramacao, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      'ORDER BY pe.NomeExecutante, pe.idProgramacaoExecutante';

    QLeitura.Parameters.ParamByName('DtIni').Value := DtIni;
    QLeitura.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    QLeitura.Open;

    QUpdate.Connection := FrmDataModule.ADOConnectionColibri;
    QUpdate.ParamCheck := True;
    QUpdate.SQL.Text :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  RT_Tipo = :RT_Tipo ' +
      'WHERE idProgramacaoExecutante = :ID';

    FrmPrincipal.ProgressBarIncializa(
      QLeitura.RecordCount,
      'Reclassificando executantes para R3...'
    );

    while not QLeitura.Eof do
    begin
      if PrecisaPrepararRegistro(QLeitura, AForcarTudo) then
      begin
        IdExec := QLeitura.FieldByName('idProgramacaoExecutante').AsInteger;
        CodigoSAPAtual := Trim(QLeitura.FieldByName('CodigoSAP').AsString);
        CodigoSAPNormalizado := CodigoSAPAtual;

        PersistirCodigoSAPNormalizado(
          IdExec,
          QLeitura.FieldByName('NomeExecutante').AsString,
          QLeitura.FieldByName('Empresa').AsString,
          CodigoSAPAtual,
          CodigoSAPNormalizado
        );

        CodigoSAPFinal := '';
        TipoRTFinal := '';

        TentarReclassificarExecutanteParaR3(
          IdExec,
          QLeitura.FieldByName('NomeExecutante').AsString,
          QLeitura.FieldByName('Empresa').AsString,
          CodigoSAPNormalizado,
          CodigoSAPFinal,
          TipoRTFinal
        );

        if Trim(TipoRTFinal) <> '' then
        begin
          QUpdate.Parameters.ParamByName('RT_Tipo').Value := TipoRTFinal;
          QUpdate.Parameters.ParamByName('ID').Value := IdExec;
          QUpdate.ExecSQL;
        end;
      end;

      QLeitura.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

    FrmPrincipal.ProgressBarAtualizar;
  finally
    QLeitura.Free;
    QUpdate.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.PrecisaPrepararRegistro(
  DS: TDataSet; const AForcarTudo: Boolean): Boolean;
var
  Status, RTNumero, MsgAtual: string;
begin
  if AForcarTudo then
    Exit(True);

  if (DS = nil) or (not DS.Active) then
    Exit(False);

  Status := StatusRTNormalizado(CampoAsString(DS, 'RT_Status'));
  RTNumero := Trim(CampoAsString(DS, 'RT'));
  MsgAtual := UpperCase(Trim(CampoAsString(DS, 'RT_Mensagem')));

  { se já tem RT, ainda precisa passar uma vez caso o status esteja vazio,
    para consolidar como EMITIDA ou JA_EXISTE }
  if (RTNumero <> '') and (RTNumero <> 'SEM RT') then
  begin
    if Status = RT_STATUS_CANCELADA then
      Exit(True);

    if Status = '' then
      Exit(True);

    { se veio da importação SAP e por algum motivo ainda não consolidou }
    if (Pos('IMPORTADO DO SAP POR', MsgAtual) > 0) and
       (Status <> RT_STATUS_JA_EXISTE) then
      Exit(True);

    Exit(False);
  end;

  if Status = '' then
    Exit(True);

  if Status = RT_STATUS_PENDENTE then
    Exit(True);

  if Status = RT_STATUS_ERRO_NAO_ATIVO then
    Exit(True);

  if Status = RT_STATUS_ERRO_CONFLITO_RT then
    Exit(True);

  if Status = RT_STATUS_ERRO_EMISSAO then
    Exit(True);

  if Status = RT_STATUS_PRONTO_EMITIR then
    Exit(False);

  if Status = RT_STATUS_NAO_CRIAR then
    Exit(False);

  if Status = RT_STATUS_EMITIDA then
    Exit(False);

  if Status = RT_STATUS_JA_EXISTE then
    Exit(False);

  if Status = RT_STATUS_ORFA then
    Exit(False);

  if Status = RT_STATUS_CANCELADA then
    Exit(False);

  Result := True;
end;

procedure TFrmConsultaExecutantesProgramados.PrepararProgramacaoExecutanteParaRT(
  const ADataIni, ADataFim: TDateTime;
  const AForcarTudo: Boolean = False);
begin
  SincronizarProgramacaoExecutanteComRTExistente(
    ADataIni,
    ADataFim
  );

  actProcurarProgramacaoExecutante.Execute;

  AplicarRegrasAutomaticasRT(ADataIni, ADataFim, AForcarTudo);

  actProcurarProgramacaoExecutante.Execute;

  AplicarRegrasRecolhimentoBanco(ADataIni, ADataFim, AForcarTudo);

  actProcurarProgramacaoExecutante.Execute;

  RecalcularClassesRTNoPeriodo(ADataIni, ADataFim, AForcarTudo);

  actProcurarProgramacaoExecutante.Execute;

  AvaliarNecessidadeCriacaoRT(ADataIni, ADataFim, AForcarTudo);
end;

function TFrmConsultaExecutantesProgramados.PrecisaReprocessarRecolhimentoOuClasse(
  DS: TDataSet; const AForcarTudo: Boolean): Boolean;
var
  Status, TipoRT, ClasseAtual, ClasseEsperada: string;
  Origem, RecolherPara, HoraVolta: string;
  BooleanRecolhimento, EhTransbordo, TemDataVolta: Boolean;
  RTNumero: string;
begin
  if AForcarTudo then
    Exit(True);

  if (DS = nil) or (not DS.Active) then
    Exit(False);

  Status := StatusRTNormalizado(CampoAsString(DS, 'RT_Status'));
  RTNumero := Trim(CampoAsString(DS, 'RT'));

  { se já existe RT válida, não reprocessa recolhimento/classe }
  if (RTNumero <> '') and (RTNumero <> 'SEM RT') then
  begin
    if Status = RT_STATUS_CANCELADA then
      Exit(True);

    Exit(False);
  end;

  { ainda não estabilizou na preparação }
  if (Status = '') or
     (Status = RT_STATUS_PENDENTE) or
     (Status = RT_STATUS_ERRO_NAO_ATIVO) or
     (Status = RT_STATUS_ERRO_CONFLITO_RT) or
     (Status = RT_STATUS_ERRO_EMISSAO) then
    Exit(True);

  { estados já consolidados não precisam reprocessar }
  if (Status = RT_STATUS_NAO_CRIAR) or
     (Status = RT_STATUS_EMITIDA) or
     (Status = RT_STATUS_JA_EXISTE) or
     (Status = RT_STATUS_ORFA) then
    Exit(False);

  TipoRT := UpperCase(Trim(CampoAsString(DS, 'RT_Tipo')));
  if TipoRT = '' then
    Exit(True);

  ClasseAtual := UpperCase(Trim(CampoAsString(DS, 'RT_Classe')));
  Origem := UpperCase(Trim(CampoAsString(DS, 'Origem')));
  RecolherPara := Trim(CampoAsString(DS, 'RecolherPara'));
  HoraVolta := Trim(CampoAsString(DS, 'RT_HoraVolta'));

  if DS.FindField('booleanRecolhimento') <> nil then
    BooleanRecolhimento := DS.FieldByName('booleanRecolhimento').AsBoolean
  else
    BooleanRecolhimento := False;

  if DS.FindField('RT_Transbordo') <> nil then
    EhTransbordo := DS.FieldByName('RT_Transbordo').AsBoolean
  else
    EhTransbordo := False;

  TemDataVolta := False;
  if DS.FindField('DataVolta') <> nil then
    TemDataVolta := not VarIsNull(DS.FieldByName('DataVolta').Value);

  { inconsistência nos campos de recolhimento }
  if BooleanRecolhimento then
  begin
    if (RecolherPara = '') or (HoraVolta = '') or (not TemDataVolta) then
      Exit(True);
  end
  else
  begin
    { limpa sobra de recolhimento somente para não-transbordo }
    if (not EhTransbordo) and
       ((RecolherPara <> '') or (HoraVolta <> '') or TemDataVolta) then
      Exit(True);
  end;

  { classe esperada }
  if EhTransbordo then
  begin
    if SameText(TipoRT, 'R7') then
      ClasseEsperada := 'COM'
    else
      ClasseEsperada := 'TR';
  end
  else
  begin
    if SameText(TipoRT, 'R7') then
      ClasseEsperada := 'COM'
    else if BooleanRecolhimento and SameText(Origem, 'TMIB') then
      ClasseEsperada := 'EV'
    else
      ClasseEsperada := 'TT';
  end;

  Result := ClasseAtual <> ClasseEsperada;
end;

function TFrmConsultaExecutantesProgramados.StatusImportacaoSAP(
  const ARTNumero, AStatusItem, AStatusDescricao: string): string;
var
  SItem, SDesc, RTNum: string;
begin
  SItem := Trim(AStatusItem);
  SDesc := UpperCase(Trim(AStatusDescricao));
  RTNum := Trim(ARTNumero);

  if (SItem = '09') or (Pos('CANCEL', SDesc) > 0) then
    Exit(RT_STATUS_CANCELADA);

  if RTNum <> '' then
    Exit(RT_STATUS_JA_EXISTE);

  Result := RT_STATUS_PENDENTE;
end;

procedure TFrmConsultaExecutantesProgramados.SincronizarProgramacaoExecutanteComRTExistente(
  const ADataIni, ADataFim: TDateTime);
var
  QRT: TADOQuery;
  DtIni, DtFimMais1: TDateTime;

  IdExecAtual, IdExecCorreto, IdRT: Integer;

  MensagemRT, RT_Numero, RT_Status: string;
  RT_HoraIda, RT_HoraVolta, RT_Modal, RT_Classe, RecolherPara: string;
  BooleanRecolhimento: Boolean;
  DataVolta: TDateTime;
  RTCanceladaLocal: Boolean;

  QtSincronizados, QtSemMatch, QtRelink, QtOrfas: Integer;
  MotivoOrfandade: string;
begin
  DtIni := Trunc(ADataIni);
  DtFimMais1 := Trunc(ADataFim) + 1;

  QRT := TADOQuery.Create(nil);
  try
    QRT.Connection := FrmDataModule.ADOConnectionRT;
    QRT.SQL.Text :=
      'SELECT * ' +
      'FROM tblProgramacaoRT ' +
      'WHERE DataIda >= :DtIni ' +
      '  AND DataIda < :DtFimMais1 ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      'ORDER BY DataIda, txtDestino, Origem, idProgramacaoRT';

    QRT.Parameters.ParamByName('DtIni').Value := DtIni;
    QRT.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    QRT.Open;

    QtSincronizados := 0;
    QtSemMatch := 0;
    QtRelink := 0;
    QtOrfas := 0;

    FrmPrincipal.ProgressBarIncializa(
      QRT.RecordCount,
      'Sincronizando programação com RT existente (soberania da programação)...'
    );

    while not QRT.Eof do
    begin
      IdRT := QRT.FieldByName('idProgramacaoRT').AsInteger;

      IdExecAtual := 0;
      if not VarIsNull(QRT.FieldByName('idProgramacaoExecutante').Value) then
        IdExecAtual := QRT.FieldByName('idProgramacaoExecutante').AsInteger;

      IdExecCorreto := BuscarProgramacaoExecutanteParaSincronizarRT(QRT);

      if (IdExecCorreto > 0) and (not ExecutanteIdAceitaRT(IdExecCorreto)) then
        IdExecCorreto := 0;

      if (IdExecAtual > 0) and (IdExecAtual <> IdExecCorreto) then
      begin
        LimparProgramacaoExecutanteRT(
          IdExecAtual,
          'RT desvinculada por inconsistência entre RT local e programação vigente'
        );
      end;

      if IdExecCorreto > 0 then
      begin
        if IdExecAtual <> IdExecCorreto then
        begin
          AtualizarVinculoProgramacaoRTExecutante(IdRT, IdExecCorreto);
          Inc(QtRelink);
        end;

        MensagemRT := QRT.FieldByName('RT_Mensagem').AsString;
        RT_Numero := QRT.FieldByName('RT').AsString;
        RT_Status := QRT.FieldByName('RT_Status').AsString;
        RT_HoraIda := QRT.FieldByName('RT_HoraIda').AsString;
        RT_HoraVolta := QRT.FieldByName('RT_HoraVolta').AsString;
        RT_Modal := QRT.FieldByName('RT_Modal').AsString;
        RT_Classe := QRT.FieldByName('RT_Classe').AsString;
        RecolherPara := QRT.FieldByName('RecolherPara').AsString;
        BooleanRecolhimento := QRT.FieldByName('booleanRecolhimento').AsBoolean;

        if QRT.FindField('RT_Cancelada') <> nil then
          RTCanceladaLocal := QRT.FieldByName('RT_Cancelada').AsBoolean
        else
          RTCanceladaLocal := False;

        if not VarIsNull(QRT.FieldByName('DataVolta').Value) then
          DataVolta := QRT.FieldByName('DataVolta').AsDateTime
        else
          DataVolta := 0;

        // se a RT está sustentada pela programação, tira marcação de órfã
        if SameText(Trim(RT_Status), RT_STATUS_PENDENTE) then
        begin
          AtualizarMarcacaoRTOrfa(
            IdRT,
            False,
            False,
            'RT sustentada pela programação vigente'
          );

          if Trim(RT_Numero) <> '' then
            RT_Status := RT_STATUS_EMITIDA
          else
            RT_Status := RT_STATUS_PRONTO_EMITIR;

          if Trim(MensagemRT) = '' then
            MensagemRT := 'RT sincronizada com a programação vigente';
        end;

        // se houve cancelamento real no SAP / espelho local,
        // a programação continua soberana e apta para regeração
        if RTLocalIndicaCancelamentoReal(
             RT_Status,
             RTCanceladaLocal) then
        begin
          RT_Numero := '';
          RT_Status := RT_STATUS_PRONTO_EMITIR;
          MensagemRT :=
            'RT cancelada no SAP/espelho local. ' +
            'Programação soberana mantida para possível nova geração.';
        end;

        AtualizarProgramacaoExecutante(
          IdExecCorreto,
          MensagemRT,
          RT_Numero,
          RT_Status,
          RT_HoraIda,
          RT_HoraVolta,
          RT_Modal,
          RT_Classe,
          RecolherPara,
          BooleanRecolhimento,
          DataVolta
        );

        Inc(QtSincronizados);
      end
      else
      begin
        AtualizarVinculoProgramacaoRTExecutante(IdRT, 0);

        MotivoOrfandade :=
          'RT sem correspondência válida na tblProgramacaoExecutante ' +
          '(origem/destino/passageiro/data inconsistentes ou registro NAO_CRIAR)';

        AtualizarMarcacaoRTOrfa(
          IdRT,
          True,
          True,
          MotivoOrfandade
        );

        Inc(QtOrfas);
        Inc(QtSemMatch);
      end;

      QRT.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

    FrmPrincipal.ProgressBarAtualizar;

    ShowMessage(
      'Sincronização concluída.' + sLineBreak +
      'Sincronizados: ' + IntToStr(QtSincronizados) + sLineBreak +
      'Relink por chave: ' + IntToStr(QtRelink) + sLineBreak +
      'RTs órfãs / pendente cancelamento: ' + IntToStr(QtOrfas) + sLineBreak +
      'Sem correspondência: ' + IntToStr(QtSemMatch)
    );
  finally
    QRT.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AplicarRegrasAutomaticasRT(
  const ADataIni, ADataFim: TDateTime;
  const AForcarTudo: Boolean = False);
var
  DS: TDataSet;
  Bmk: TBookmark;
  HasBookmark: Boolean;
  SelRegistro: Integer;

  AOrigem, ADestino: string;
  CodigoSAPOriginal, CodigoSAPNormalizado, TipoRTNormalizado: string;
  Modal, TipoRT, Classe: string;

  PlataformaDestino, PlataformaOrigem: TDadosPlataforma;
  Horario: THorario;
  DataProgramacao: TDateTime;
  BooleanRecolhimentoAtual: Boolean;
begin
  DS := FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
    Exit;

  SelRegistro := 0;
  HasBookmark := False;

  if not DS.IsEmpty then
  begin
    SelRegistro := DS.RecNo;
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  FrmPrincipal.ProgressBarIncializa(
    DS.RecordCount,
    'Aplicando regras automáticas de RT...'
  );

  try
    DS.DisableControls;
    try
      DS.First;

      while not DS.Eof do
      begin
        if not PrecisaPrepararRegistro(DS, AForcarTudo) then
        begin
          DS.Next;
          FrmPrincipal.ProgressBarIncremento(1);
          Continue;
        end;

        DataProgramacao := DS.FieldByName('DataProgramacao').AsDateTime;
        AOrigem := Trim(DS.FieldByName('Origem').AsString);
        ADestino := Trim(DS.FieldByName('txtDestino').AsString);

        CodigoSAPOriginal := Trim(DS.FieldByName('CodigoSAP').AsString);
        CodigoSAPNormalizado := NormalizarCodigoSAP(CodigoSAPOriginal);

        if Trim(CodigoSAPNormalizado) <> '' then
          TipoRTNormalizado := DeterminarTipoRTAutomatico(CodigoSAPNormalizado)
        else
          TipoRTNormalizado := '';

        if DS.FindField('booleanRecolhimento') <> nil then
          BooleanRecolhimentoAtual := DS.FieldByName('booleanRecolhimento').AsBoolean
        else
          BooleanRecolhimentoAtual := False;

        PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(AOrigem);
        PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(ADestino);

        Horario := HoraIdaHoraVolta(
          PlataformaDestino.booleanTurno2,
          DataProgramacao
        );

        Modal := DeterminarModalAutomatico(AOrigem, ADestino);

        if Trim(CodigoSAPNormalizado) <> '' then
          TipoRT := DeterminarTipoRTAutomatico(CodigoSAPNormalizado)
        else
          TipoRT := DeterminarTipoRTAutomatico(CodigoSAPOriginal);

        Classe := Trim(DS.FieldByName('RT_Classe').AsString);
        Classe := DetermineClasse(
          TipoRT,
          AOrigem,
          ADestino,
          FrmPrincipal.OrigemPlataformas,
          Classe,
          BooleanRecolhimentoAtual
        );

        DS.Edit;
        try
          if (Trim(CodigoSAPNormalizado) <> '') and
             (not SameText(Trim(DS.FieldByName('CodigoSAP').AsString), Trim(CodigoSAPNormalizado))) then
            DS.FieldByName('CodigoSAP').AsString := CodigoSAPNormalizado;

          if (Trim(TipoRTNormalizado) <> '') and
             (Trim(DS.FieldByName('RT_Tipo').AsString) = '') then
            DS.FieldByName('RT_Tipo').AsString := TipoRTNormalizado;

          if (Trim(AOrigem) <> '') and (Trim(ADestino) <> '') then
          begin
            PreencherCamposAutomaticosRTNoRegistro(
              DS,
              TipoRT,
              Modal,
              Classe,
              Horario
            );

            AplicarRecolhimentoNoRegistro(
              DS,
              Horario,
              AOrigem,
              BooleanRecolhimentoAtual
            );
          end;

          if not DeveCriarRTAutomaticamente(
            AOrigem,
            ADestino,
            PlataformaOrigem,
            PlataformaDestino
          ) then
            DS.FieldByName('RT_Mensagem').AsString := 'NÃO CRIAR RT';

          DS.Post;
        except
          DS.Cancel;
          raise;
        end;

        if (Trim(CodigoSAPNormalizado) <> '') and
           (not SameText(Trim(CodigoSAPOriginal), Trim(CodigoSAPNormalizado))) then
        begin
          AtualizarCodigoSAPTblExecutante(
            DS.FieldByName('NomeExecutante').AsString,
            DS.FieldByName('Empresa').AsString,
            CodigoSAPNormalizado
          );
        end;

        DS.Next;
        FrmPrincipal.ProgressBarIncremento(1);
      end;
    finally
      if HasBookmark then
      begin
        try
          if DS.BookmarkValid(Bmk) then
            DS.GotoBookmark(Bmk);
        finally
          DS.FreeBookmark(Bmk);
        end;
      end;

      DS.EnableControls;
      FrmPrincipal.ProgressBarAtualizar;
    end;
  finally
    if (SelRegistro > 0) and (SelRegistro <= DS.RecordCount) then
      DS.RecNo := SelRegistro;
  end;

  AplicarTransbordoNoPeriodo(ADataIni, ADataFim);
end;

procedure TFrmConsultaExecutantesProgramados.GarantirTabelaRTRegraRecolhimento;
var
  Q: TADOQuery;

  function ExisteRegra(
    const AOrigem, ADestino, ADestinoNaoCriarRT,
    ARecolherParaTipo, ARecolherParaValor: string): Boolean;
  begin
    Q.Close;
    Q.SQL.Clear;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'SELECT TOP 1 idRegraRecolhimento '+
      'FROM tblRTRegraRecolhimento '+
      'WHERE Trim(Origem) = :Origem '+
      '  AND Trim(txtDestino) = :txtDestino '+
      '  AND Trim(DestinoNaoCriarRT) = :DestinoNaoCriarRT '+
      '  AND Trim(RecolherParaTipo) = :RecolherParaTipo '+
      '  AND Trim(RecolherParaValor) = :RecolherParaValor';

    Q.Parameters.ParamByName('Origem').Value := Trim(AOrigem);
    Q.Parameters.ParamByName('txtDestino').Value := Trim(ADestino);
    Q.Parameters.ParamByName('DestinoNaoCriarRT').Value := Trim(ADestinoNaoCriarRT);
    Q.Parameters.ParamByName('RecolherParaTipo').Value := Trim(ARecolherParaTipo);
    Q.Parameters.ParamByName('RecolherParaValor').Value := Trim(ARecolherParaValor);

    Q.Open;
    Result := not Q.Eof;
  end;

  procedure InserirRegra(
    const AAtiva: Boolean;
    const APrioridade: Integer;
    const ADescricao, AOrigem, ADestino, ADestinoNaoCriarRT,
    ARecolherParaTipo, ARecolherParaValor, AObservacao: string);
  begin
    if ExisteRegra(
         AOrigem,
         ADestino,
         ADestinoNaoCriarRT,
         ARecolherParaTipo,
         ARecolherParaValor) then
      Exit;

    Q.Close;
    Q.SQL.Clear;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'INSERT INTO tblRTRegraRecolhimento ('+
      '  Ativa, Prioridade, Descricao, '+
      '  Origem, txtDestino, DestinoNaoCriarRT, '+
      '  RecolherParaTipo, RecolherParaValor, Observacao '+
      ') VALUES ('+
      '  :Ativa, :Prioridade, :Descricao, '+
      '  :Origem, :txtDestino, :DestinoNaoCriarRT, '+
      '  :RecolherParaTipo, :RecolherParaValor, :Observacao'+
      ')';

    Q.Parameters.ParamByName('Ativa').Value := AAtiva;
    Q.Parameters.ParamByName('Prioridade').Value := APrioridade;
    Q.Parameters.ParamByName('Descricao').Value := Copy(Trim(ADescricao), 1, 100);

    Q.Parameters.ParamByName('Origem').Value := Copy(Trim(AOrigem), 1, 20);
    Q.Parameters.ParamByName('txtDestino').Value := Copy(Trim(ADestino), 1, 20);
    Q.Parameters.ParamByName('DestinoNaoCriarRT').Value := Copy(Trim(ADestinoNaoCriarRT), 1, 3);

    Q.Parameters.ParamByName('RecolherParaTipo').Value := Copy(Trim(ARecolherParaTipo), 1, 10);
    Q.Parameters.ParamByName('RecolherParaValor').Value := Copy(Trim(ARecolherParaValor), 1, 20);
    Q.Parameters.ParamByName('Observacao').Value := Copy(Trim(AObservacao), 1, 255);

    Q.ExecSQL;
  end;

begin
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
    '[DestinoNaoCriarRT] VARCHAR(3), '+

    '[RecolherParaTipo] VARCHAR(10), '+
    '[RecolherParaValor] VARCHAR(20), '+

    '[Observacao] VARCHAR(255)',

    FrmDataModule.ADOConnectionRT
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;

    // Regra atual: TMIB e destino que cria RT => recolher para origem
    InserirRegra(
      True,
      10,
      'Origem TMIB com recolhimento para origem',
      'TMIB',
      '',
      'NAO',
      'ORIGEM',
      '',
      'Regra padrão inicial'
    );

    // Regra atual: PCM-09 e destino que cria RT => recolher para origem
    InserirRegra(
      True,
      20,
      'Origem PCM-09 com recolhimento para origem',
      'PCM-09',
      '',
      'NAO',
      'ORIGEM',
      '',
      'Regra padrão inicial'
    );

  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarProgramacaoExecutanteParaRecuperacao(
  const DadosSAP: TRTSapImportDados): Integer;
var
  Q: TADOQuery;
  OrigemLocal, DestinoLocal: string;
  DocumentoLimpo, DocumentoFmt: string;
  CodigoSAPRaw, CodigoSAPNorm: string;
  DtIni, DtFim: TDateTime;

  procedure PrepararBase;
  begin
    Q.Close;
    Q.SQL.Clear;
    Q.Parameters.Clear;
    Q.ParamCheck := True;

    Q.SQL.Add('SELECT TOP 1 pe.idProgramacaoExecutante ');
    Q.SQL.Add('FROM tblProgramacaoDiaria pd ');
    Q.SQL.Add('INNER JOIN tblProgramacaoExecutante pe ');
    Q.SQL.Add('  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ');
    Q.SQL.Add('WHERE pd.DataProgramacao >= :DataIni ');
    Q.SQL.Add('  AND pd.DataProgramacao < :DataFim ');
    Q.SQL.Add('  AND Trim(pe.Origem) = :Origem ');
    Q.SQL.Add('  AND Trim(pd.txtDestino) = :Destino ');

    with Q.Parameters.ParamByName('DataIni') do
    begin
      DataType := ftDateTime;
      Value := DtIni;
    end;

    with Q.Parameters.ParamByName('DataFim') do
    begin
      DataType := ftDateTime;
      Value := DtFim;
    end;

    with Q.Parameters.ParamByName('Origem') do
    begin
      DataType := ftString;
      Size := 50;
      Value := OrigemLocal;
    end;

    with Q.Parameters.ParamByName('Destino') do
    begin
      DataType := ftString;
      Size := 50;
      Value := DestinoLocal;
    end;
  end;

begin
  Result := 0;

  if DadosSAP.DataViagem <= 0 then
    Exit;

  OrigemLocal := PlataformaPorNomeSAP(DadosSAP.Origem);
  DestinoLocal := PlataformaPorNomeSAP(DadosSAP.Destino);

  CodigoSAPRaw := Trim(DadosSAP.PERNR);
  CodigoSAPNorm := NormalizarCodigoSAP(CodigoSAPRaw);

  DocumentoLimpo := FrmPrincipal.SomenteNumero(Trim(DadosSAP.Documento));
  DocumentoFmt := FrmPrincipal.FormatarCPF(DocumentoLimpo);

  DtIni := Trunc(DadosSAP.DataViagem);
  DtFim := DtIni + 1;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;

    // 1) tenta por Código SAP, aceitando valor bruto e normalizado
    if CodigoSAPNorm <> '' then
    begin
      PrepararBase;

      Q.SQL.Add('  AND (Trim(pe.CodigoSAP) = :CodigoSAP1 ');

      if SameText(CodigoSAPRaw, CodigoSAPNorm) then
        Q.SQL.Add('       OR Trim(pe.CodigoSAP) = :CodigoSAP1) ')
      else
        Q.SQL.Add('       OR Trim(pe.CodigoSAP) = :CodigoSAP2) ');

      Q.SQL.Add('ORDER BY pe.idProgramacaoExecutante DESC');

      with Q.Parameters.ParamByName('CodigoSAP1') do
      begin
        DataType := ftString;
        Size := 30;
        Value := CodigoSAPNorm;
      end;

      if not SameText(CodigoSAPRaw, CodigoSAPNorm) then
      begin
        with Q.Parameters.ParamByName('CodigoSAP2') do
        begin
          DataType := ftString;
          Size := 30;
          Value := CodigoSAPRaw;
        end;
      end;

      Q.Open;
      if not Q.Eof then
        Exit(Q.FieldByName('idProgramacaoExecutante').AsInteger);
    end;

    // 2) tenta por documento
    if DocumentoLimpo <> '' then
    begin
      PrepararBase;

      Q.SQL.Add('  AND (Trim(pe.Documento) = :Documento1 ');
      Q.SQL.Add('       OR Trim(pe.Documento) = :Documento2 ');
      Q.SQL.Add('       OR Trim(pe.OutroDocumento) = :Documento3) ');
      Q.SQL.Add('ORDER BY pe.idProgramacaoExecutante DESC');

      with Q.Parameters.ParamByName('Documento1') do
      begin
        DataType := ftString;
        Size := 30;
        Value := DocumentoLimpo;
      end;

      with Q.Parameters.ParamByName('Documento2') do
      begin
        DataType := ftString;
        Size := 30;
        Value := DocumentoFmt;
      end;

      with Q.Parameters.ParamByName('Documento3') do
      begin
        DataType := ftString;
        Size := 50;
        Value := Trim(DadosSAP.Documento);
      end;

      Q.Open;
      if not Q.Eof then
        Exit(Q.FieldByName('idProgramacaoExecutante').AsInteger);
    end;

    // 3) tenta por nome
    if Trim(DadosSAP.Passageiro) <> '' then
    begin
      PrepararBase;

      Q.SQL.Add('  AND Trim(pe.NomeExecutante) = :NomeExecutante ');
      Q.SQL.Add('ORDER BY pe.idProgramacaoExecutante DESC');

      with Q.Parameters.ParamByName('NomeExecutante') do
      begin
        DataType := ftString;
        Size := 120;
        Value := Trim(DadosSAP.Passageiro);
      end;

      Q.Open;
      if not Q.Eof then
        Exit(Q.FieldByName('idProgramacaoExecutante').AsInteger);
    end;

  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.NormalizarCodigoSAP(const ACodigo: string): string;
var
  I: Integer;
  ApenasNumeros: Boolean;
begin
  /// <summary>
  /// Normaliza o C�digo SAP removendo zeros � esquerda se for num�rico.
  /// Regra de neg�cio cr�tica para evitar problemas na integra��o com SAP.
  /// </summary>
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

procedure TFrmConsultaExecutantesProgramados.AtualizarCodigoSAPTblExecutante(
  const ANomeExecutante, AEmpresa, ACodigoSAPNovo: string);
var
  Q: TADOQuery;
begin
  if Trim(ANomeExecutante) = '' then
    Exit;

  if Trim(ACodigoSAPNovo) = '' then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionConsulta;
    Q.ParamCheck := True;
    Q.SQL.Clear;
    Q.SQL.Text :=
      'UPDATE tblExecutante ' +
      'SET CodigoSAP = :pCodigoSAP ' +
      'WHERE UCase(Trim(NomeExecutante)) = :pNome ' +
      '  AND UCase(Trim(txtEmpresa)) = :pEmpresa';

    Q.Parameters.ParamByName('pCodigoSAP').DataType := ftString;
    Q.Parameters.ParamByName('pCodigoSAP').Value := Trim(ACodigoSAPNovo);

    Q.Parameters.ParamByName('pNome').DataType := ftString;
    Q.Parameters.ParamByName('pNome').Value := UpperCase(Trim(ANomeExecutante));

    Q.Parameters.ParamByName('pEmpresa').DataType := ftString;
    Q.Parameters.ParamByName('pEmpresa').Value := UpperCase(Trim(AEmpresa));

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.PersistirCodigoSAPNormalizado(
  const AIdProgramacaoExecutante: Integer;
  const ANomeExecutante, AEmpresa, ACodigoSAPOriginal: string;
  out ACodigoSAPNormalizado: string);
var
  TipoRTNormalizado: string;
begin
  ACodigoSAPNormalizado := NormalizarCodigoSAP(ACodigoSAPOriginal);

  if Trim(ACodigoSAPOriginal) = '' then
    Exit;

  if Trim(ACodigoSAPNormalizado) = '' then
    Exit;

  if SameText(Trim(ACodigoSAPOriginal), Trim(ACodigoSAPNormalizado)) then
    Exit;

  TipoRTNormalizado := DeterminarTipoRTAutomatico(ACodigoSAPNormalizado);

  AtualizarCodigoSAPProgramacaoExecutante(
    AIdProgramacaoExecutante,
    ACodigoSAPNormalizado
  );

  AtualizarCodigoSAPTblExecutante(
    ANomeExecutante,
    AEmpresa,
    ACodigoSAPNormalizado
  );
end;


function TFrmConsultaExecutantesProgramados.ExecutanteAceitaRT(
  DS: TDataSet): Boolean;
var
  AOrigem, ADestino: string;
  PlataformaOrigem, PlataformaDestino: TDadosPlataforma;
begin
  Result := False;

  if (DS = nil) or (not DS.Active) then
    Exit;

  AOrigem := Trim(DS.FieldByName('Origem').AsString);

  if DS.FindField('txtDestino') <> nil then
    ADestino := Trim(DS.FieldByName('txtDestino').AsString)
  else
    ADestino := '';

  PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(AOrigem);
  PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(ADestino);

  Result := DeveCriarRTAutomaticamente(
    AOrigem,
    ADestino,
    PlataformaOrigem,
    PlataformaDestino
  );
end;

function TFrmConsultaExecutantesProgramados.ExecutanteIdAceitaRT(
  const AIdProgramacaoExecutante: Integer): Boolean;
var
  Q: TADOQuery;
begin
  Result := False;

  if AIdProgramacaoExecutante <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'SELECT pd.DataProgramacao, pd.txtDestino, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pe.idProgramacaoExecutante = :ID';

    Q.Parameters.ParamByName('ID').Value := AIdProgramacaoExecutante;
    Q.Open;

    if not Q.Eof then
      Result := ExecutanteAceitaRT(Q);
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarVinculoProgramacaoRTExecutante(
  const AIdProgramacaoRT, AIdProgramacaoExecutante: Integer);
var
  Q: TADOQuery;
begin
  if AIdProgramacaoRT <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT SET ' +
      '  idProgramacaoExecutante = :idProgramacaoExecutante ' +
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    if AIdProgramacaoExecutante > 0 then
      Q.Parameters.ParamByName('idProgramacaoExecutante').Value := AIdProgramacaoExecutante
    else
      Q.Parameters.ParamByName('idProgramacaoExecutante').Value := Null;

    Q.Parameters.ParamByName('idProgramacaoRT').Value := AIdProgramacaoRT;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.LimparProgramacaoExecutanteRT(
  const AIdProgramacaoExecutante: Integer;
  const AMotivo: string);
var
  Q: TADOQuery;
begin
  if AIdProgramacaoExecutante <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  RT_Mensagem = :RT_Mensagem, ' +
      '  [RT] = Null, ' +
      '  RT_Status = Null ' +
      'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    Q.Parameters.ParamByName('RT_Mensagem').Value :=
      Copy(Trim(AMotivo), 1, 100);
    Q.Parameters.ParamByName('idProgramacaoExecutante').Value :=
      AIdProgramacaoExecutante;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarProgramacaoExecutanteParaSincronizarRT(
  DSRT: TDataSet): Integer;
var
  Q: TADOQuery;
  DtIni, DtFimMais1, DataIdaRT: TDateTime;
  OrigemRT, DestinoRT, OrigemLocal, DestinoLocal: string;
  CodigoSAPRT, ChavePassRT, TipoPass, ValorPass: string;
  DocumentoNumerico, DocumentoFmt: string;

  function ExtrairTipoValorChavePassageiro(
    const AChave: string;
    out ATipo, AValor: string): Boolean;
  var
    P: Integer;
    S: string;
  begin
    Result := False;
    ATipo := '';
    AValor := '';

    S := Trim(AChave);
    if S = '' then
      Exit;

    P := Pos('=', S);
    if P <= 0 then
      Exit;

    ATipo := UpperCase(Trim(Copy(S, 1, P - 1)));
    AValor := Trim(Copy(S, P + 1, MaxInt));

    Result := (ATipo <> '') and (AValor <> '');
  end;

begin
  Result := 0;

  if (DSRT = nil) or (not DSRT.Active) then
    Exit;

  if VarIsNull(DSRT.FieldByName('DataIda').Value) then
    Exit;

  DataIdaRT := Trunc(DSRT.FieldByName('DataIda').AsDateTime);

  OrigemRT := Trim(DSRT.FieldByName('Origem').AsString);
  DestinoRT := Trim(DSRT.FieldByName('txtDestino').AsString);

  OrigemLocal := PlataformaPorNomeSAP(OrigemRT);
  DestinoLocal := PlataformaPorNomeSAP(DestinoRT);

  CodigoSAPRT := RemoverZerosEsquerda(Trim(DSRT.FieldByName('CodigoSAP').AsString));
  ChavePassRT := Trim(DSRT.FieldByName('ChavePassageiro').AsString);

  DtIni := DataIdaRT;
  DtFimMais1 := DtIni + 1;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;

    Q.SQL.Clear;
    Q.SQL.Add(
      'SELECT pd.DataProgramacao, pd.txtDestino, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      '  AND Trim(pe.Origem) = :Origem ' +
      '  AND Trim(pd.txtDestino) = :Destino '
    );

    Q.Parameters.ParamByName('DtIni').Value := DtIni;
    Q.Parameters.ParamByName('DtFimMais1').Value := DtFimMais1;
    Q.Parameters.ParamByName('Origem').Value := OrigemLocal;
    Q.Parameters.ParamByName('Destino').Value := DestinoLocal;

    if CodigoSAPRT <> '' then
    begin
      Q.SQL.Add('  AND Trim(pe.CodigoSAP) = :CodigoSAP ');
      Q.Parameters.ParamByName('CodigoSAP').Value := CodigoSAPRT;
    end
    else if ExtrairTipoValorChavePassageiro(ChavePassRT, TipoPass, ValorPass) then
    begin
      if TipoPass = 'SAP' then
      begin
        Q.SQL.Add('  AND Trim(pe.CodigoSAP) = :CodigoSAP ');
        Q.Parameters.ParamByName('CodigoSAP').Value := ValorPass;
      end
      else if (TipoPass = 'C') or (TipoPass = 'DOC') then
      begin
        DocumentoNumerico := FrmPrincipal.SomenteNumero(ValorPass);
        DocumentoFmt := FrmPrincipal.FormatarCPF(DocumentoNumerico);

        Q.SQL.Add(
          '  AND (Trim(pe.Documento) = :Documento1 ' +
          '       OR Trim(pe.Documento) = :Documento2) '
        );
        Q.Parameters.ParamByName('Documento1').Value := DocumentoNumerico;
        Q.Parameters.ParamByName('Documento2').Value := DocumentoFmt;
      end
      else if TipoPass = 'P' then
      begin
        Q.SQL.Add('  AND Trim(pe.OutroDocumento) = :Passaporte ');
        Q.Parameters.ParamByName('Passaporte').Value := ValorPass;
      end
      else
        Exit;
    end
    else
      Exit;

    Q.SQL.Add('ORDER BY pe.idProgramacaoExecutante DESC');
    Q.Open;

    while not Q.Eof do
    begin
      if ExecutanteAceitaRT(Q) then
      begin
        Result := Q.FieldByName('idProgramacaoExecutante').AsInteger;
        Exit;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ComplementarProgramacaoExecutanteViaSAPImport(
  const AIdProgramacaoExecutante: Integer;
  const DadosSAP: TRTSapImportDados);
var
  Q: TADOQuery;
begin
  if AIdProgramacaoExecutante <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  RT_HoraIda = IIf((RT_HoraIda IS NULL OR Trim(RT_HoraIda) = ''''), :RT_HoraIda, RT_HoraIda), ' +
      '  RT_Modal   = IIf((RT_Modal IS NULL OR Trim(RT_Modal) = ''''), :RT_Modal, RT_Modal), ' +
      '  RT_Classe  = IIf((RT_Classe IS NULL OR Trim(RT_Classe) = ''''), :RT_Classe, RT_Classe) ' +
      'WHERE idProgramacaoExecutante = :ID';

    Q.Parameters.ParamByName('RT_HoraIda').Value := Trim(DadosSAP.HoraViagem);
    Q.Parameters.ParamByName('RT_Modal').Value := Trim(DadosSAP.RT_Modal);
    Q.Parameters.ParamByName('RT_Classe').Value := Trim(DadosSAP.RT_Classe);
    Q.Parameters.ParamByName('ID').Value := AIdProgramacaoExecutante;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ComplementarProgramacaoRTViaSAPImport(
  const AIdProgramacaoRT: Integer;
  const DadosSAP: TRTSapImportDados);
var
  Q: TADOQuery;
begin
  if AIdProgramacaoRT <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT SET ' +
      '  RT_HoraIda = IIf((RT_HoraIda IS NULL OR Trim(RT_HoraIda) = ''''), :RT_HoraIda, RT_HoraIda), ' +
      '  RT_Modal   = IIf((RT_Modal IS NULL OR Trim(RT_Modal) = ''''), :RT_Modal, RT_Modal), ' +
      '  RT_Classe  = IIf((RT_Classe IS NULL OR Trim(RT_Classe) = ''''), :RT_Classe, RT_Classe) ' +
      'WHERE idProgramacaoRT = :ID';

    Q.Parameters.ParamByName('RT_HoraIda').Value := Trim(DadosSAP.HoraViagem);
    Q.Parameters.ParamByName('RT_Modal').Value := Trim(DadosSAP.RT_Modal);
    Q.Parameters.ParamByName('RT_Classe').Value := Trim(DadosSAP.RT_Classe);
    Q.Parameters.ParamByName('ID').Value := AIdProgramacaoRT;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.RecuperarRetornoRTsViaImportacaoSAP(
  const ADataIni, ADataFim: TDateTime);
var
  QImp: TADOQuery;
  DadosSAP: TRTSapImportDados;
  IdExec, IdRT, IdExecJaVinculado: Integer;
  QtRecuperadas, QtSemExec, QtSemRT, QtJaExistiam, QtEspelhosOrfos: Integer;
  DtIni, DtFimMais1: TDateTime;
  ChaveIda, ChaveCompleta: string;
  QtdeMatch: Integer;
begin
  QtRecuperadas := 0;
  QtSemExec := 0;
  QtSemRT := 0;
  QtJaExistiam := 0;
  QtEspelhosOrfos := 0;

  DtIni := Trunc(ADataIni);
  DtFimMais1 := Trunc(ADataFim) + 1;

  QImp := TADOQuery.Create(nil);
  try
    QImp.Connection := FrmDataModule.ADOConnectionRT;
    QImp.ParamCheck := True;
    QImp.Close;
    QImp.SQL.Clear;
    QImp.Parameters.Clear;

    QImp.SQL.Text :=
      'SELECT tblRTSapImport.* ' +
      'FROM tblRTSapImport ' +
      'WHERE [DataViagem] >= :pDtIni ' +
      '  AND [DataViagem] < :pDtFimMais1 ' +
      '  AND [QMNUM] IS NOT NULL ' +
      '  AND [QMNUM] <> '''' ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = 0) ' +
      'ORDER BY [DataViagem], [QMNUM], [idRTSapImport]';

    QImp.Parameters.ParamByName('pDtIni').DataType := ftDateTime;
    QImp.Parameters.ParamByName('pDtIni').Direction := pdInput;
    QImp.Parameters.ParamByName('pDtIni').Value := DtIni;

    QImp.Parameters.ParamByName('pDtFimMais1').DataType := ftDateTime;
    QImp.Parameters.ParamByName('pDtFimMais1').Direction := pdInput;
    QImp.Parameters.ParamByName('pDtFimMais1').Value := DtFimMais1;

    QImp.Open;

    FrmPrincipal.ProgressBarIncializa(
      QImp.RecordCount,
      'Recuperando RTs pela tblRTSapImport...'
    );

    while not QImp.Eof do
    begin
      try
        DadosSAP := Default(TRTSapImportDados);

        DadosSAP.idRTSapImport := QImp.FieldByName('idRTSapImport').AsInteger;
        DadosSAP.QMNUM := Trim(QImp.FieldByName('QMNUM').AsString);
        DadosSAP.QMART := Trim(QImp.FieldByName('QMART').AsString);
        DadosSAP.IWERK := Trim(QImp.FieldByName('IWERK').AsString);
        DadosSAP.INGRP := Trim(QImp.FieldByName('INGRP').AsString);
        DadosSAP.Origem := Trim(QImp.FieldByName('Origem').AsString);
        DadosSAP.Destino := Trim(QImp.FieldByName('txtDestino').AsString);

        if not VarIsNull(QImp.FieldByName('DataViagem').Value) then
          DadosSAP.DataViagem := QImp.FieldByName('DataViagem').AsDateTime
        else
          DadosSAP.DataViagem := 0;

        DadosSAP.HoraViagem := Trim(QImp.FieldByName('HoraViagem').AsString);
        DadosSAP.PERNR := Trim(QImp.FieldByName('PERNR').AsString);
        DadosSAP.TipoDoc := Trim(QImp.FieldByName('TipoDoc').AsString);
        DadosSAP.Documento := Trim(QImp.FieldByName('Documento').AsString);
        DadosSAP.Passageiro := Trim(QImp.FieldByName('Passageiro').AsString);
        DadosSAP.QMTXT := Trim(QImp.FieldByName('QMTXT').AsString);
        DadosSAP.RT_Modal := Trim(QImp.FieldByName('RT_Modal').AsString);
        DadosSAP.RT_Classe := Trim(QImp.FieldByName('RT_Classe').AsString);
        DadosSAP.StatusItem := Trim(QImp.FieldByName('StatusItem').AsString);
        DadosSAP.StatusDescricao := Trim(QImp.FieldByName('StatusDescricao').AsString);

        ChaveIda := '';
        ChaveCompleta := '';

        if QImp.FindField('ChaveIda') <> nil then
          ChaveIda := Trim(QImp.FieldByName('ChaveIda').AsString);

        if QImp.FindField('ChaveCompleta') <> nil then
          ChaveCompleta := Trim(QImp.FieldByName('ChaveCompleta').AsString);

        IdRT := 0;
        IdExecJaVinculado := 0;

        // =========================================================
        // 1) SEMPRE localizar executante na base LOCAL
        //    (não confiar no idProgramacaoExecutante gravado na importação)
        // =========================================================
        try
          IdExec := BuscarProgramacaoExecutanteParaRecuperacao(DadosSAP);
        except
          on E: Exception do
          begin
            ShowMessage(
              'Falha ao localizar executante para recuperação.' + sLineBreak +
              'idRTSapImport=' + IntToStr(DadosSAP.idRTSapImport) + sLineBreak +
              'QMNUM=' + DadosSAP.QMNUM + sLineBreak +
              'Passageiro=' + DadosSAP.Passageiro + sLineBreak +
              'Origem=' + DadosSAP.Origem + sLineBreak +
              'Destino=' + DadosSAP.Destino + sLineBreak +
              'Erro=' + E.Message
            );
            raise;
          end;
        end;

        // =========================================================
        // 2) SEMPRE localizar RT na base LOCAL
        //    (não confiar no idProgramacaoRT gravado na importação)
        // =========================================================
        QtdeMatch := 0;

        if Trim(DadosSAP.QMNUM) <> '' then
          QtdeMatch := BuscarProgramacaoRTPorNumeroRT(
            DadosSAP.QMNUM,
            IdRT,
            IdExecJaVinculado
          );

        if (QtdeMatch = 0) and (Trim(ChaveIda) <> '') then
          QtdeMatch := BuscarProgramacaoRTPorChave(
            'ChaveIda',
            ChaveIda,
            IdRT,
            IdExecJaVinculado
          );

        if (QtdeMatch = 0) and (Trim(ChaveCompleta) <> '') then
          QtdeMatch := BuscarProgramacaoRTPorChave(
            'ChaveCompleta',
            ChaveCompleta,
            IdRT,
            IdExecJaVinculado
          );

        if QtdeMatch > 0 then
          Inc(QtJaExistiam);

        // =========================================================
        // 3) Se não existe RT local, cria
        // =========================================================
        if IdRT <= 0 then
        begin
          if IdExec > 0 then
            IdRT := CriarRTLocalViaImportacaoSAP(DadosSAP, IdExec)
          else
          begin
            IdRT := CriarEspelhoLocalRTOrfaViaImportacaoSAP(DadosSAP);
            if IdRT > 0 then
              Inc(QtEspelhosOrfos);
          end;
        end;

        // =========================================================
        // 4) Se ainda assim não conseguiu criar/localizar RT local
        // =========================================================
        if IdRT <= 0 then
        begin
          AtualizarVinculoRTSapImport(
            DadosSAP.idRTSapImport,
            0,
            IdExec,
            False,
            'Falhou ao criar/localizar tblProgramacaoRT'
          );

          if IdExec <= 0 then
            Inc(QtSemExec)
          else
            Inc(QtSemRT);

          QImp.Next;
          FrmPrincipal.ProgressBarIncremento(1);
          Continue;
        end;

        // =========================================================
        // 5) Atualiza RT local com status importado do SAP
        // =========================================================
        AtualizarProgramacaoRTViaImportacaoSAP(
          IdRT,
          DadosSAP.QMNUM,
          DadosSAP.StatusItem,
          DadosSAP.StatusDescricao,
          'RECUPERACAO_IMPORTACAO_SAP'
        );

        ComplementarProgramacaoRTViaSAPImport(IdRT, DadosSAP);

        // =========================================================
        // 6) Se encontrou executante local, atualiza também
        // =========================================================
        if IdExec > 0 then
        begin
          AtualizarProgramacaoExecutanteViaImportacaoSAP(
            IdExec,
            DadosSAP.QMNUM,
            DadosSAP.StatusItem,
            DadosSAP.StatusDescricao,
            'RECUPERACAO_IMPORTACAO_SAP'
          );

          ComplementarProgramacaoExecutanteViaSAPImport(IdExec, DadosSAP);
        end
        else
        begin
          Inc(QtSemExec);
        end;

        // =========================================================
        // 7) Grava o vínculo local reconstruído
        // =========================================================
        AtualizarVinculoRTSapImport(
          DadosSAP.idRTSapImport,
          IdRT,
          IdExec,
          True,
          'Recuperado automaticamente a partir da tblRTSapImport'
        );

        Inc(QtRecuperadas);
      except
        on E: Exception do
        begin
          raise Exception.CreateFmt(
            'Erro no idRTSapImport=%d, QMNUM=%s, Passageiro=%s. Erro: %s',
            [
              QImp.FieldByName('idRTSapImport').AsInteger,
              Trim(QImp.FieldByName('QMNUM').AsString),
              Trim(QImp.FieldByName('Passageiro').AsString),
              E.Message
            ]
          );
        end;
      end;

      QImp.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

    FrmPrincipal.ProgressBarAtualizar;

    ShowMessage(
      'Recuperação concluída.' + sLineBreak +
      'Recuperadas/criadas localmente: ' + IntToStr(QtRecuperadas) + sLineBreak +
      'RT local já existente: ' + IntToStr(QtJaExistiam) + sLineBreak +
      'Espelhos órfãos criados: ' + IntToStr(QtEspelhosOrfos) + sLineBreak +
      'Sem executante local: ' + IntToStr(QtSemExec) + sLineBreak +
      'Sem RT local criada/localizada: ' + IntToStr(QtSemRT)
    );

  finally
    QImp.Free;
  end;
end;

function GetADODataSource(AConn: TADOConnection): string;
var
  S: string;
  P1, P2: Integer;
begin
  Result := '';

  if AConn = nil then
    Exit;

  S := AConn.ConnectionString;
  P1 := Pos('Data Source=', S);
  if P1 > 0 then
  begin
    S := Copy(S, P1 + Length('Data Source='), MaxInt);
    P2 := Pos(';', S);
    if P2 > 0 then
      S := Copy(S, 1, P2 - 1);
    Result := Trim(S);
  end;
end;

function TFrmConsultaExecutantesProgramados.PlataformaPorNomeSAP(
  const ANomeSAP: string): string;
var
  Q: TADOQuery;
  Nome: string;
begin
  Nome := Trim(ANomeSAP);
  Result := Nome;

  if Nome = '' then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    try
      Q.Connection := FrmDataModule.ADOConnectionConsulta;
      Q.Close;
      Q.SQL.Clear;
      Q.Parameters.Clear;
      Q.ParamCheck := True;

      Q.SQL.Text :=
        'SELECT TOP 1 [Plataforma] ' +
        'FROM [tblPlataforma] ' +
        'WHERE Trim([NomeSAP]) = :pNomeSAP1 ' +
        '   OR Trim([Plataforma]) = :pNomeSAP2 ' +
        'ORDER BY [idPlataforma]';

      Q.Parameters.ParamByName('pNomeSAP1').DataType := ftString;
      Q.Parameters.ParamByName('pNomeSAP1').Value := Nome;

      Q.Parameters.ParamByName('pNomeSAP2').DataType := ftString;
      Q.Parameters.ParamByName('pNomeSAP2').Value := Nome;

      Q.Open;

      if not Q.Eof then
        Result := Trim(Q.FieldByName('Plataforma').AsString);
    except
      Result := Nome;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarProgramacaoExecutantePorImportacaoSAP(
  const DadosSAP: TRTSapImportDados;
  out AIdProgramacaoExecutante: Integer): Boolean;
begin
  AIdProgramacaoExecutante :=
    BuscarProgramacaoExecutanteParaRecuperacao(DadosSAP);

  Result := AIdProgramacaoExecutante > 0;
end;

function TFrmConsultaExecutantesProgramados.CriarRTLocalViaImportacaoSAP(
  const DadosSAP: TRTSapImportDados;
  const AIdProgramacaoExecutante: Integer): Integer;
var
  Q: TADOQuery;
  Dados: TDadosRT;
  PlataformaOrigem, PlataformaDestino, PlataformaRetorno: TDadosPlataforma;
  CodigoCusto: TCodicoCusto;
  RecolherParaLocal: string;
  ClasseCalc: string;
begin
  Result := 0;

  if AIdProgramacaoExecutante <= 0 then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.SQL.Text :=
      'SELECT TOP 1 pd.DataProgramacao, pd.txtDestino, pe.* ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pe.idProgramacaoExecutante = :ID';

    Q.Parameters.ParamByName('ID').Value := AIdProgramacaoExecutante;
    Q.Open;

    if Q.Eof then
      Exit;

    if not FrmDataModule.ADOQueryConfigRT.Active then
      FrmDataModule.ADOQueryConfigRT.Active := True;

    Dados := Default(TDadosRT);

    Dados.idProgramacaoExecutante := AIdProgramacaoExecutante;

    PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(
      Trim(Q.FieldByName('Origem').AsString)
    );

    PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(
      Trim(Q.FieldByName('txtDestino').AsString)
    );

    RecolherParaLocal := Trim(Q.FieldByName('RecolherPara').AsString);
    if RecolherParaLocal <> '' then
      PlataformaRetorno := TProgramacaoRTUtils.DadosPlataforma_RT(RecolherParaLocal)
    else
      PlataformaRetorno := Default(TDadosPlataforma);

    Dados.Origem := PlataformaOrigem.NomeSAP;
    Dados.Destino := PlataformaDestino.NomeSAP;

    if RecolherParaLocal <> '' then
      Dados.Retorno := PlataformaRetorno.NomeSAP
    else
      Dados.Retorno := Dados.Origem;

    Dados.MatriculaPax := RemoverZerosEsquerda(Trim(Q.FieldByName('CodigoSAP').AsString));
    Dados.CPF := Trim(Q.FieldByName('Documento').AsString);
    Dados.Passaporte := Trim(Q.FieldByName('OutroDocumento').AsString);

    Dados.TipoRT := Trim(DadosSAP.QMART);
    if Dados.TipoRT = '' then
      Dados.TipoRT := DeterminarTipoRTAutomatico(Dados.MatriculaPax);

    Dados.Modal := Trim(DadosSAP.RT_Modal);
    if Dados.Modal = '' then
      Dados.Modal := DeterminarModalAutomatico(
        Trim(Q.FieldByName('Origem').AsString),
        Trim(Q.FieldByName('txtDestino').AsString)
      );

    Dados.booleanRecolhimento := Q.FieldByName('booleanRecolhimento').AsBoolean;

    ClasseCalc := Trim(DadosSAP.RT_Classe);
    if ClasseCalc = '' then
      ClasseCalc := DetermineClasse(
        Dados.TipoRT,
        Trim(Q.FieldByName('Origem').AsString),
        Trim(Q.FieldByName('txtDestino').AsString),
        FrmPrincipal.OrigemPlataformas,
        Trim(Q.FieldByName('RT_Classe').AsString),
        Dados.booleanRecolhimento
      );
    Dados.Classe := ClasseCalc;

    Dados.Requisitante  := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_Requisitante').AsString);
    Dados.PessoaContato := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_PessoaContato').AsString);
    Dados.TelContato    := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_TelContato').AsString);
    Dados.CentroPlan    := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_CentroPlan').AsString);
    Dados.GrpPlan       := Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_GrpPlan').AsString);

    Dados.DataIda := DataSAP(Q.FieldByName('DataProgramacao').AsDateTime);

    Dados.HoraIda := Trim(DadosSAP.HoraViagem);
    if Dados.HoraIda = '' then
      Dados.HoraIda := Trim(Q.FieldByName('RT_HoraIda').AsString);
    if Dados.HoraIda = '' then
      Dados.HoraIda := '07:00';

    Dados.HoraVolta := Trim(Q.FieldByName('RT_HoraVolta').AsString);
    if Dados.HoraVolta = '' then
      Dados.HoraVolta := '16:30';

    if not VarIsNull(Q.FieldByName('DataVolta').Value) then
      Dados.DataVolta := DataSAP(Q.FieldByName('DataVolta').AsDateTime)
    else
      Dados.DataVolta := '';

    if Trim(DadosSAP.QMTXT) <> '' then
      Dados.DescricaoRT := Copy(Trim(DadosSAP.QMTXT), 1, 40)
    else
      Dados.DescricaoRT := Copy(
        Dados.Origem + ': ' + Trim(Q.FieldByName('NomeExecutante').AsString), 1, 40
      );

    if PlataformaDestino.booleanProntidao then
    begin
      Dados.CentroCusto  := PlataformaDestino.CentroCusto;
      Dados.DiagramaRede := PlataformaDestino.DiagramaRede;
      Dados.OperRede     := PlataformaDestino.OperRede;
      Dados.ElementoPEP  := PlataformaDestino.ElementoPEP;
    end
    else
    begin
      CodigoCusto := CustoExecutante(
        Dados.MatriculaPax,
        Dados.CPF,
        Dados.Passaporte
      );

      if (Trim(CodigoCusto.CentroCusto) = '') and
         (Trim(CodigoCusto.ElementoPEP) = '') and
         ((Trim(CodigoCusto.DiagramaRede) = '') or (Trim(CodigoCusto.OperRede) = '')) then
      begin
        Dados.CentroCusto  := PlataformaDestino.CentroCusto;
        Dados.DiagramaRede := PlataformaDestino.DiagramaRede;
        Dados.OperRede     := PlataformaDestino.OperRede;
        Dados.ElementoPEP  := PlataformaDestino.ElementoPEP;
      end
      else
      begin
        Dados.CentroCusto  := CodigoCusto.CentroCusto;
        Dados.DiagramaRede := CodigoCusto.DiagramaRede;
        Dados.OperRede     := CodigoCusto.OperRede;
        Dados.ElementoPEP  := CodigoCusto.ElementoPEP;
      end;
    end;

    Dados.RTCobertura := False;

    Result := CriarRegistroTabelaRT(Dados);
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.RegistroDentroDoPeriodo(
  const DS: TDataSet;
  const ADataIni, ADataFim: TDateTime): Boolean;
var
  Dt: TDateTime;
begin
  Result := False;

  if (DS = nil) or (DS.FindField('DataProgramacao') = nil) then
    Exit;

  if DS.FieldByName('DataProgramacao').IsNull then
    Exit;

  Dt := Trunc(DS.FieldByName('DataProgramacao').AsDateTime);
  Result := (Dt >= Trunc(ADataIni)) and (Dt <= Trunc(ADataFim));
end;

procedure TFrmConsultaExecutantesProgramados.AvaliarNecessidadeCriacaoRT(
  const ADataIni, ADataFim: TDateTime;
  const AForcarTudo: Boolean = False);
var
  DS: TDataSet;
  Status, Msg: string;
begin
  DS := FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet;
  if DS = nil then
    Exit;

  DS.DisableControls;
  try
    DS.First;
    while not DS.Eof do
    begin
      if RegistroDentroDoPeriodo(DS, ADataIni, ADataFim) then
      begin
        if PrecisaPrepararRegistro(DS, AForcarTudo) then
        begin
          Status := '';
          Msg := '';

          AvaliarNecessidadeCriacaoRT(DS, Status, Msg);

          if Status = '' then
            Status := RT_STATUS_PENDENTE;

          if Trim(Msg) = '' then
            Msg := MSG_RT_PENDENTE;

          DS.Edit;

          if DS.FindField('RT_Status') <> nil then
            DS.FieldByName('RT_Status').AsString := Status;

          if DS.FindField('RT_Mensagem') <> nil then
            DS.FieldByName('RT_Mensagem').AsString := Msg;

          DS.Post;
        end;
      end;

      DS.Next;
    end;
  finally
    DS.EnableControls;
  end;
end;

function TFrmConsultaExecutantesProgramados.AvaliarNecessidadeCriacaoRT(
  const DS: TDataSet;
  out Status, Mensagem: string): Boolean;

  function CampoExiste(const ANome: string): Boolean;
  begin
    Result := (DS <> nil) and (DS.FindField(ANome) <> nil);
  end;

  function TextoCampo(const ANome: string): string;
  begin
    if CampoExiste(ANome) then
      Result := Trim(DS.FieldByName(ANome).AsString)
    else
      Result := '';
  end;

  function BoolCampo(const ANome: string): Boolean;
  begin
    if CampoExiste(ANome) then
      Result := DS.FieldByName(ANome).AsBoolean
    else
      Result := False;
  end;

  function DataCampoValida(const ANome: string): Boolean;
  begin
    Result := CampoExiste(ANome) and
      (not DS.FieldByName(ANome).IsNull) and
      (DS.FieldByName(ANome).AsDateTime > 0);
  end;

  function TipoRTValido(const ATipoRT: string): Boolean;
  begin
    Result := SameText(Trim(ATipoRT), 'R3') or SameText(Trim(ATipoRT), 'R7');
  end;

  function DocumentoR7Preenchido: Boolean;
  begin
    Result :=
      (TextoCampo('Documento') <> '') or
      (TextoCampo('OutroDocumento') <> '');
  end;

  function TemRateioPreenchido(out AMensagemRateio: string): Boolean;
  var
    CentroCusto, DiagramaRede, OperRede, ElementoPEP: string;
    CodigoSAPUsado: string;
    NomeExecutante, Empresa, CodigoSAPLocal, OrigemLocal, DestinoLocal: string;
    IdProgramacaoExecutante: Integer;
    DestinoEhProntidao, OrigemEhProntidao: Boolean;
    CCD, DRD, ORD, EPD: string;
    CCO, DRO, ORO, EPO: string;
  begin
    Result := False;
    AMensagemRateio := '';

    NomeExecutante := TextoCampo('NomeExecutante');
    Empresa := TextoCampo('Empresa');
    if Empresa = '' then
      Empresa := TextoCampo('txtEmpresa');

    CodigoSAPLocal := TextoCampo('CodigoSAP');
    OrigemLocal := TextoCampo('Origem');
    DestinoLocal := TextoCampo('txtDestino');

    if CampoExiste('idProgramacaoExecutante') then
      IdProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger
    else
      IdProgramacaoExecutante := 0;

    { 1. Destino prontidão }
    DestinoEhProntidao := False;
    if ObterInfoTblPlataforma(DestinoLocal, DestinoEhProntidao, CCD, DRD, ORD, EPD)
       and DestinoEhProntidao then
    begin
      if RateioValido(CCD, DRD, ORD, EPD) then
        Exit(True);

      AMensagemRateio :=
        'Plataforma de destino de prontidão sem rateio na tblPlataforma: ' +
        DestinoLocal + '. Preencher CentroCusto, ou DiagramaRede + OperRede, ou ElementoPEP.';
      Exit(False);
    end;

    { 2. Origem prontidão }
    OrigemEhProntidao := False;
    if ObterInfoTblPlataforma(OrigemLocal, OrigemEhProntidao, CCO, DRO, ORO, EPO)
       and OrigemEhProntidao then
    begin
      if RateioValido(CCO, DRO, ORO, EPO) then
        Exit(True);

      AMensagemRateio :=
        'Plataforma de origem de prontidão sem rateio na tblPlataforma: ' +
        OrigemLocal + '. Preencher CentroCusto, ou DiagramaRede + OperRede, ou ElementoPEP.';
      Exit(False);
    end;

    { 3. Executante: CodigoSAP -> fallback Nome + Empresa }
    if ObterRateioExecutanteComFallback(
         IdProgramacaoExecutante,
         CodigoSAPLocal,
         NomeExecutante,
         Empresa,
         CodigoSAPUsado,
         CentroCusto,
         DiagramaRede,
         OperRede,
         ElementoPEP) then
      Exit(True);

    if Trim(CodigoSAPLocal) <> '' then
      AMensagemRateio :=
        'Preencher CentroCusto, ou DiagramaRede + OperRede, ou ElementoPEP na tblExecutante para o CodigoSAP ' +
        Trim(CodigoSAPLocal) + '.'
    else
      AMensagemRateio :=
        'Preencher CentroCusto, ou DiagramaRede + OperRede, ou ElementoPEP na tblExecutante para o executante ' +
        NomeExecutante + ' / ' + Empresa + '.';
  end;

var
  TipoRT, CodigoSAP, HoraIda, HoraVolta, Modal, Classe: string;
  Origem, Destino, RecolherPara, MsgRateio: string;
  BooleanRecolhimento: Boolean;
  PlataformaOrigem, PlataformaDestino: TDadosPlataforma;
  RTAtual, StatusAtual, MsgAtual: string;
  MsgAtualUp: string;
begin
  Result := False;
  Status := '';
  Mensagem := '';

  if DS = nil then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Dataset não informado para avaliação da RT.';
    Exit;
  end;

  RTAtual := Trim(CampoAsString(DS, 'RT'));
  StatusAtual := StatusRTNormalizado(CampoAsString(DS, 'RT_Status'));
  MsgAtual := Trim(CampoAsString(DS, 'RT_Mensagem'));
  MsgAtualUp := UpperCase(MsgAtual);

  { 1. Se já veio marcado como JA_EXISTE pela importação SAP, preservar }
  if StatusAtual = RT_STATUS_JA_EXISTE then
  begin
    Status := RT_STATUS_JA_EXISTE;

    if MsgAtual <> '' then
      Mensagem := MsgAtual
    else if RTAtual <> '' then
      Mensagem := 'RT já existente no SAP: ' + RTAtual
    else
      Mensagem := MSG_RT_JA_EXISTE;

    Exit(False);
  end;

  { 2. Se veio da importação SAP com status diferente de cancelado, preservar como JA_EXISTE }
  if (Pos('IMPORTADO DO SAP POR', MsgAtualUp) > 0) and
     (Pos('STATUS 09', MsgAtualUp) = 0) then
  begin
    Status := RT_STATUS_JA_EXISTE;

    if MsgAtual <> '' then
      Mensagem := MsgAtual
    else
      Mensagem := MSG_RT_JA_EXISTE;

    Exit(False);
  end;

  { 3. Se já existe RT válida, preservar }
  if (RTAtual <> '') and (RTAtual <> 'SEM RT') and
     (StatusAtual <> RT_STATUS_CANCELADA) then
  begin
    if StatusAtual = RT_STATUS_JA_EXISTE then
      Status := RT_STATUS_JA_EXISTE
    else
      Status := RT_STATUS_EMITIDA;

    if MsgAtual <> '' then
      Mensagem := MsgAtual
    else if Status = RT_STATUS_JA_EXISTE then
      Mensagem := MSG_RT_JA_EXISTE
    else
      Mensagem := MSG_RT_EMITIDA;

    Exit(False);
  end;

  TipoRT := TextoCampo('RT_Tipo');
  CodigoSAP := TextoCampo('CodigoSAP');
  HoraIda := TextoCampo('RT_HoraIda');
  HoraVolta := TextoCampo('RT_HoraVolta');
  Modal := TextoCampo('RT_Modal');
  Classe := TextoCampo('RT_Classe');
  Origem := TextoCampo('Origem');
  Destino := TextoCampo('txtDestino');
  RecolherPara := TextoCampo('RecolherPara');
  BooleanRecolhimento := BoolCampo('booleanRecolhimento');

  PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(Origem);
  PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(Destino);

  if not DeveCriarRTAutomaticamente(Origem, Destino, PlataformaOrigem, PlataformaDestino) then
  begin
    Status := RT_STATUS_NAO_CRIAR;
    Mensagem := MSG_RT_NAO_CRIAR;
    Exit(False);
  end;

  if not TipoRTValido(TipoRT) then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Tipo de RT inválido ou não preenchido.';
    Exit;
  end;

  if Origem = '' then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Origem não preenchida.';
    Exit;
  end;

  if Destino = '' then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Destino não preenchido.';
    Exit;
  end;

  if not DataCampoValida('DataProgramacao') then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Data da programação não preenchida.';
    Exit;
  end;

  if HoraIda = '' then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Hora de ida não preenchida.';
    Exit;
  end;

  if Modal = '' then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Modal não preenchido.';
    Exit;
  end;

  if Classe = '' then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'Classe não preenchida.';
    Exit;
  end;

  if SameText(TipoRT, 'R3') and (CodigoSAP = '') then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'RT do tipo R3 exige Código SAP preenchido.';
    Exit;
  end;

  if SameText(TipoRT, 'R7') and (not DocumentoR7Preenchido) then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := 'RT do tipo R7 exige Documento ou OutroDocumento.';
    Exit;
  end;

  if BooleanRecolhimento then
  begin
    if RecolherPara = '' then
    begin
      Status := RT_STATUS_PENDENTE;
      Mensagem := 'Recolhimento marcado sem campo RecolherPara preenchido.';
      Exit;
    end;

    if HoraVolta = '' then
    begin
      Status := RT_STATUS_PENDENTE;
      Mensagem := 'Recolhimento marcado sem HoraVolta preenchida.';
      Exit;
    end;

    if not DataCampoValida('DataVolta') then
    begin
      Status := RT_STATUS_PENDENTE;
      Mensagem := 'Recolhimento marcado sem DataVolta preenchida.';
      Exit;
    end;
  end;

  if not TemRateioPreenchido(MsgRateio) then
  begin
    Status := RT_STATUS_PENDENTE;
    Mensagem := MsgRateio;
    Exit;
  end;

  Status := RT_STATUS_PRONTO_EMITIR;
  Mensagem := MSG_RT_PRONTO_EMITIR;
  Result := True;
end;

function TFrmConsultaExecutantesProgramados.ExecutanteSustentaRT(
  DS: TDataSet): Boolean;
var
  AOrigem, ADestino: string;
  PlataformaOrigem, PlataformaDestino: TDadosPlataforma;
  StatusRT: string;
begin
  Result := False;

  if (DS = nil) or (not DS.Active) then
    Exit;

  AOrigem := Trim(DS.FieldByName('Origem').AsString);
  ADestino := Trim(DS.FieldByName('txtDestino').AsString);

  PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(AOrigem);
  PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(ADestino);

  // regra soberana: se não deve criar RT, não sustenta
  if not DeveCriarRTAutomaticamente(AOrigem, ADestino, PlataformaOrigem, PlataformaDestino) then
    Exit(False);

  StatusRT := CampoAsString(DS, 'RT_Status');

  Result := StatusRTSustentaRT(StatusRT);
end;

function TFrmConsultaExecutantesProgramados.StatusRTSustentaRT(
  const AStatus: string): Boolean;
var
  S: string;
begin
  S := StatusRTNormalizado(AStatus);

  Result :=
    (S = RT_STATUS_PRONTO_EMITIR) or
    (S = RT_STATUS_EMITIDA) or
    (S = RT_STATUS_JA_EXISTE) or
    (S = RT_STATUS_PENDENTE);
end;

procedure TFrmConsultaExecutantesProgramados.actImportRTSAPExecute(
  Sender: TObject);
begin
  if MessageDlg(
       'Deseja excluir as RTs importadas do período selecionado e importar novamente do SAP?',
       mtConfirmation, [mbYes, mbNo], 0
     ) = mrYes then
  begin
    ReimportarRTsSAPPeriodo;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actImprmirExecute(Sender: TObject);
  var
    SelRegistro,NumLinhas: Integer;
    strDataInicio,strDataFim,txt2,txt1: String;
begin
  RLTemporario.RowCount:= 1;
  RLTemporario.ColCount:= 4;
  RLTemporario.Cells[0,0]:= 'Nome';
  RLTemporario.Cells[1,0]:= 'Função';
  RLTemporario.Cells[2,0]:= 'Empresa';
  RLTemporario.Cells[3,0]:= 'Assinatura';
  SelRegistro:= FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
  while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
  begin
    numLinhas:= RLTemporario.RowCount;
    RLTemporario.RowCount:= numLinhas+1;
    RLTemporario.Cells[0,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('NomeExecutante').AsString;
    RLTemporario.Cells[1,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('Funcao').AsString;
    RLTemporario.Cells[2,numLinhas]:= FrmDataModule.
    DataSourceConsultaExecutantesProgramados.DataSet.FieldByName('Empresa').AsString;
    RLTemporario.Cells[3,numLinhas]:= '';
    FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
  end;
  FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecNo:= SelRegistro;
  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
  AutoFitGrade(RLTemporario);
  RLTemporario.ColWidths[0]:= RLTemporario.ColWidths[0]+20;
  RLTemporario.ColWidths[1]:= RLTemporario.ColWidths[1]+20;
  RLTemporario.ColWidths[2]:= RLTemporario.ColWidths[2]+20;
  RLTemporario.ColWidths[3]:= RLTemporario.ColWidths[3]+20;
  //===================================================================
  if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;

  strDataInicio:= FrmPrincipal.corrigirData(dataInicio.DateTime);
  strDataFim:= FrmPrincipal.corrigirData(dataFim.DateTime);
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


procedure TFrmConsultaExecutantesProgramados.actProcurarSomenteTransbordosExecute(Sender: TObject);
var
  SQLString, SQLBase, DataProcuraInicio, DataProcuraFim: string;
begin
  DataProcuraInicio := FormatDateTime('mm/dd/yyyy', dataInicio.Date);
  DataProcuraFim    := FormatDateTime('mm/dd/yyyy', dataFim.Date);

  //=======================================================
  SQLString := FrmPrincipal.SQLStringFiltroTabela(ColunasLayoutExecutanteProgramado, False);
  if Trim(SQLString) <> '' then
    SQLString := ' AND ' + SQLString;

  //=======================================================
  SQLBase :=
    'SELECT pd.*, pe.* ' +
    'FROM tblProgramacaoDiaria pd ' +
    'INNER JOIN tblProgramacaoExecutante pe ' +
    '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
    'WHERE (pd.DataProgramacao BETWEEN #' + DataProcuraInicio + '# AND #' + DataProcuraFim + '#) ' +
    '  AND (pe.RT_Transbordo = True) ' +
    '  AND (pe.RT_GrupoTransbordo IS NOT NULL) ' +
    '  AND (pe.RT_GrupoTransbordo <> '''') ' +
    SQLString +
    ' ORDER BY ' +
    '   pd.DataProgramacao, ' +
    '   pe.RT_GrupoTransbordo, ' +
    '   pe.RT_SeqTransbordo, ' +
    '   pe.NomeExecutante, ' +
    '   pe.Origem, ' +
    '   pd.txtDestino;';

  FrmDataModule.ProcuraQueryCompleta(
    SQLBase,
    FrmDataModule.ADOQueryConsultaExecutantesProgramados,
    StatusBarExecutantes
  );
end;

function TFrmConsultaExecutantesProgramados.PodeGerarRTNoRegistro(
  DS: TDataSet; out MotivoBloqueio: string): Boolean;
var
  Status, RTNumero: string;
begin
  MotivoBloqueio := '';

  Status := StatusRTNormalizado(CampoAsString(DS, 'RT_Status'));
  RTNumero := Trim(CampoAsString(DS, 'RT'));

  if (RTNumero <> '') and (RTNumero <> 'SEM RT') then
  begin
    MotivoBloqueio := 'Registro já possui RT.';
    Exit(False);
  end;

  if Status = RT_STATUS_PRONTO_EMITIR then
    Exit(True);

  if Status = RT_STATUS_NAO_CRIAR then
    MotivoBloqueio := 'Marcado como NÃO CRIAR.'
  else if Status = RT_STATUS_PENDENTE then
    MotivoBloqueio := 'Registro pendente de revisão.'
  else if Status = RT_STATUS_JA_EXISTE then
    MotivoBloqueio := 'Registro já possui RT existente.'
  else if Status = RT_STATUS_EMITIDA then
    MotivoBloqueio := 'Registro já emitido.'
  else if Status = RT_STATUS_CANCELADA then
    MotivoBloqueio := 'Registro cancelado.'
  else if Status = RT_STATUS_ERRO_NAO_ATIVO then
    MotivoBloqueio := 'Código SAP não ativo.'
  else if Status = RT_STATUS_ERRO_CONFLITO_RT then
    MotivoBloqueio := 'Conflito com RT já existente.'
  else if Status = RT_STATUS_ERRO_EMISSAO then
    MotivoBloqueio := 'Erro de emissão.'
  else
    MotivoBloqueio := 'Registro ainda não está apto para emissão.';

  Result := False;
end;

function TFrmConsultaExecutantesProgramados.NormalizarMensagemLog(
  const AMensagem: string): string;
begin
  Result := Trim(AMensagem);
  Result := StringReplace(Result, ' n o ', ' não ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' nao ', ' não ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' est ', ' está ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' est  ', ' está ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' obrigat rio ', ' obrigatório ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' nÃ£o ', ' não ', [rfReplaceAll, rfIgnoreCase]);
  Result := StringReplace(Result, ' estÃ¡ ', ' está ', [rfReplaceAll, rfIgnoreCase]);
end;

procedure TFrmConsultaExecutantesProgramados.actLimparExecute(Sender: TObject);

  procedure LimparCampoString(DS: TDataSet; const ACampo: string);
  begin
    if DS.FindField(ACampo) <> nil then
      DS.FieldByName(ACampo).AsString := '';
  end;

  procedure LimparCampoClear(DS: TDataSet; const ACampo: string);
  begin
    if DS.FindField(ACampo) <> nil then
      DS.FieldByName(ACampo).Clear;
  end;

  procedure LimparCampoBoolean(DS: TDataSet; const ACampo: string; const AValor: Boolean);
  begin
    if DS.FindField(ACampo) <> nil then
      DS.FieldByName(ACampo).AsBoolean := AValor;
  end;

var
  DS: TDataSet;
begin
  DS := FrmDataModule.ADOQueryConsultaExecutantesProgramados;

  FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled := False;
  try
    DS.First;
    while not DS.Eof do
    begin
      DS.Edit;
      try
        LimparCampoString(DS, 'RT');
        LimparCampoBoolean(DS, 'RT_Transbordo', False);
        LimparCampoBoolean(DS, 'RT_TransbordoAereo', False);
        LimparCampoBoolean(DS, 'booleanRecolhimento', False);

        LimparCampoString(DS, 'RT_HoraIda');
        LimparCampoString(DS, 'RT_HoraVolta');
        LimparCampoString(DS, 'RT_GrupoTransbordo');
        LimparCampoString(DS, 'RT_PrimeiraOrigemTransbordo');
        LimparCampoString(DS, 'RT_PrimeiroDestinoTransbordo');
        LimparCampoString(DS, 'RT_SeqTransbordo');

        LimparCampoString(DS, 'RT_Status');
        LimparCampoString(DS, 'RT_Mensagem');

        LimparCampoString(DS, 'RT_Classe');
        LimparCampoString(DS, 'RT_Modal');
        LimparCampoString(DS, 'RT_Tipo');
        LimparCampoString(DS, 'RecolherPara');

        LimparCampoClear(DS, 'DataVolta');

        DS.Post;
      except
        DS.Cancel;
        raise;
      end;

      DS.Next;
    end;

    DS.First;
  finally
    FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled := True;
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

procedure TFrmConsultaExecutantesProgramados.actProcurarProgramacaoRTExecute(
  Sender: TObject);
var
  SQLString, SQLBase, DataProcuraInicio, DataProcuraFim: string;
begin
  SQLString := FrmPrincipal.SQLStringFiltroTabela(ColunasLayoutRT, False);
  if Trim(SQLString) <> '' then
    SQLString := ' AND ' + SQLString;

  DataProcuraInicio := FormatDateTime('mm/dd/yyyy', Trunc(dataInicio.Date));
  DataProcuraFim    := FormatDateTime('mm/dd/yyyy', Trunc(dataFim.Date + 1));

  SQLBase :=
    'SELECT tblProgramacaoRT.* ' +
    'FROM tblProgramacaoRT ' +
    'WHERE [DataIda] >= #' + DataProcuraInicio + '# ' +
    '  AND [DataIda] < #' + DataProcuraFim + '# ' +SQLString+
    ' ORDER BY [DataIda], [txtDestino], [Origem]';

  FrmDataModule.ProcuraQueryCompleta(
    SQLBase,
    FrmDataModule.ADOQueryGestaoRT,
    StatusBarGestaoRT
  );
end;

procedure TFrmConsultaExecutantesProgramados.actProcurarBuscaEmbarqueExecute(
  Sender: TObject);
var
  SQLString, SQLBase,
  DataProcuraInicio, DataProcuraFim: String;
begin
  DataProcuraInicio := FormatDateTime('mm/dd/yyyy', Trunc(dataInicio.Date));
  DataProcuraFim    := FormatDateTime('mm/dd/yyyy', Trunc(dataFim.Date + 1));

  //=======================================================
  SQLString := frmPrincipal.SQLStringFiltroTabela(RLLayoutBuscaEmbarque, False);

  if SQLString <> '' then
    SQLString := ' AND ' + SQLString;

  SQLBase :=
    'SELECT ' +
    '  tblProgramacaoDiaria.DataProgramacao, ' +
    '  tblProgramacaoDiaria.txtDestino, ' +
    '  tblProgramacaoExecutante.CodigoSAP, ' +
    '  tblProgramacaoExecutante.NomeExecutante, ' +
    '  tblProgramacaoExecutante.Funcao, ' +
    '  tblProgramacaoExecutante.Empresa, ' +
    '  tblRoteamento.DataRoteamento, ' +
    '  tblRoteamento.NomeRota, ' +
    '  tblRoteamento.HoraRoteamento, ' +
    '  tblRoteamento.NomeEmbarcacao, ' +
    '  tblRoteamento.CapacidadePAX, ' +
    '  tblRoteamento.RotaSequencia, ' +
    '  tblProgramacaoExecutante.Origem ' +
    'FROM ' +
    '  ((tblProgramacaoDiaria ' +
    '  INNER JOIN tblProgramacaoExecutante ' +
    '    ON tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria) ' +
    '  INNER JOIN (tblRoteamento ' +
    '    INNER JOIN tblAux_Rota_Distribuicao ' +
    '      ON tblRoteamento.idRoteamento = tblAux_Rota_Distribuicao.CodigoRota) ' +
    '    ON tblProgramacaoExecutante.idProgramacaoExecutante = tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ' +
    'WHERE ' +
    '  tblProgramacaoDiaria.DataProgramacao >= #' + DataProcuraInicio + '# ' +
    '  AND tblProgramacaoDiaria.DataProgramacao < #' + DataProcuraFim + '# ' +
    SQLString +
    ' ORDER BY ' +
    '  tblProgramacaoDiaria.DataProgramacao, ' +
    '  tblProgramacaoDiaria.txtDestino, ' +
    '  tblProgramacaoExecutante.Origem, ' +
    '  tblRoteamento.HoraRoteamento;';

  FrmDataModule.ProcuraQueryCompleta(
    SQLBase,
    FrmDataModule.ADOQueryFrequenciaResumo,
    StatusBarFrequenciaResumo
  );
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
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutExecutanteProgramado,false);
  //Query de procura
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE (DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'+
    SQLString+SQL_OrigemDestino+' ORDER BY DataProgramacao,txtDestino,Origem,txtTipoEtapaServico;';
  FrmDataModule.ProcuraQueryCompleta(SQLBase,FrmDataModule.
  ADOQueryConsultaExecutantesProgramados,StatusBarExecutantes);
end;

procedure TFrmConsultaExecutantesProgramados.actProcurarSAPImportExecute(
  Sender: TObject);
var
  SQLString, SQLBase, DataProcuraInicio, DataProcuraFim: string;
begin
  DataProcuraInicio := FormatDateTime('mm/dd/yyyy', Trunc(dataInicio.Date));
  DataProcuraFim    := FormatDateTime('mm/dd/yyyy', Trunc(dataFim.Date + 1));

  SQLString := FrmPrincipal.SQLStringFiltroTabela(ColunasLayoutSAPImport, False);
  if Trim(SQLString) <> '' then
    SQLString := ' AND ' + SQLString;

  SQLBase :=
    'SELECT tblRTSapImport.* ' +
    'FROM tblRTSapImport ' +
    'WHERE [DataViagem] >= #' + DataProcuraInicio + '# ' +
    '  AND [DataViagem] <  #' + DataProcuraFim + '# ' +
    SQLString +
    ' ORDER BY [txtDestino], [Origem];';

  FrmDataModule.ProcuraQueryCompleta(
    SQLBase,
    FrmDataModule.ADOQuerySAPImport,
    StatusBarSAPImport
  );
end;

procedure TFrmConsultaExecutantesProgramados.LimparRTSapImportPeriodo(
  const ADataIni, ADataFim: TDateTime);
var
  Q: TADOQuery;
  DataIni, DataFimMaisUm: TDateTime;
begin
  DataIni := Trunc(ADataIni);
  DataFimMaisUm := IncDay(Trunc(ADataFim), 1);

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'DELETE FROM tblRTSapImport ' +
      'WHERE [DataViagem] >= :DataIni ' +
      '  AND [DataViagem] < :DataFimMaisUm';

    Q.Parameters.ParamByName('DataIni').DataType := ftDateTime;
    Q.Parameters.ParamByName('DataIni').Value := DataIni;

    Q.Parameters.ParamByName('DataFimMaisUm').DataType := ftDateTime;
    Q.Parameters.ParamByName('DataFimMaisUm').Value := DataFimMaisUm;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ReimportarRTsSAPPeriodo;
begin
  LimparRTSapImportPeriodo(dataInicio.Date, dataFim.Date);
  ImportarRTsSAPPeriodo;
  actProcurarSAPImport.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.actRT_TMIBExecute(Sender: TObject);
  var
    Origem: String;
begin
  if not Assigned(FrmTabela) then
    Application.CreateForm(TFrmTabela, FrmTabela);
  FrmTabela.Show;   // ou ShowModal conforme o caso

  FrmTabela.Caption:= 'Requisições de transporte';
  Origem:=  'TMIB';
  FrmTabela.PanelTitulo.Caption:= 'Requisições de transporte com origem: '+Origem;
  TProgramacaoRTUtils.ListaRTEmbarque(FrmTabela.RLTabela, dataInicio.DateTime,
   Origem);
end;

procedure TFrmConsultaExecutantesProgramados.actHistorizarProgramacao_DadosSAPExecute(Sender: TObject);
var
  Bmk: TBookmark;
  DS: TDataSet;
  HasBookmark: Boolean;
  idProgramacaoExecutante: Integer;
  MensagemRT, RT_Numero, RT_Status,
  RT_HoraIda, RT_HoraVolta, RT_Modal, RT_Classe, RecolherPara: string;
  booleanRecolhimento: Boolean;
  DataVolta: TDateTime;
begin
  if Application.MessageBox(PChar(
    'Deseja realmente historizar a programação de RT e de executantes com as RTs importadas do SAP?'),
    '.::ATENÇÃO::.', 36) <> 6 then
    Exit;

  RecuperarRetornoRTsViaImportacaoSAP(dataInicio.Date, dataFim.Date);

  FrmDataModule.ADOQueryGestaoRT.Active := False;
  FrmDataModule.ADOQueryGestaoRT.Active := True;

  DS := FrmDataModule.DataSourceGestaoRT.DataSet;
  if (DS = nil) or (not DS.Active) or DS.IsEmpty then
    Exit;

  HasBookmark := False;
  if not DS.IsEmpty then
  begin
    Bmk := DS.GetBookmark;
    HasBookmark := True;
  end;

  FrmPrincipal.ProgressBarIncializa(
    DS.RecordCount,
    'Atualizando dados da tblProgramacaoExecutante.'
  );

  try
    DS.DisableControls;
    try
      DS.First;
      while not DS.Eof do
      begin
        idProgramacaoExecutante := DS.FieldByName('idProgramacaoExecutante').AsInteger;

        if idProgramacaoExecutante > 0 then
        begin
          MensagemRT := CampoAsString(DS, 'RT_Mensagem');
          RT_Numero := CampoAsString(DS, 'RT');
          RT_Status := CampoAsString(DS, 'RT_Status');
          RT_HoraIda := CampoAsString(DS, 'RT_HoraIda');
          RT_HoraVolta := CampoAsString(DS, 'RT_HoraVolta');
          RT_Modal := CampoAsString(DS, 'RT_Modal');
          RT_Classe := CampoAsString(DS, 'RT_Classe');
          RecolherPara := CampoAsString(DS, 'RecolherPara');
          booleanRecolhimento := (DS.FindField('booleanRecolhimento') <> nil) and
                                 DS.FieldByName('booleanRecolhimento').AsBoolean;

          if (DS.FindField('DataVolta') <> nil) and (not VarIsNull(DS.FieldByName('DataVolta').Value)) then
            DataVolta := DS.FieldByName('DataVolta').AsDateTime
          else
            DataVolta := 0;

          AtualizarProgramacaoExecutante(
            idProgramacaoExecutante,
            MensagemRT,
            RT_Numero,
            RT_Status,
            RT_HoraIda,
            RT_HoraVolta,
            RT_Modal,
            RT_Classe,
            RecolherPara,
            booleanRecolhimento,
            DataVolta
          );
        end;

        DS.Next;
        FrmPrincipal.ProgressBarIncremento(1);
      end;
    finally
      if HasBookmark then
      begin
        try
          if DS.BookmarkValid(Bmk) then
            DS.GotoBookmark(Bmk);
        finally
          DS.FreeBookmark(Bmk);
        end;
      end;

      DS.EnableControls;
      FrmPrincipal.ProgressBarAtualizar;
    end;
  except
    on E: Exception do
      MessageBox(
        0,
        PChar('Erro ao atualizar dados da RT na programação: ' + E.Message),
        'Colibri',
        MB_ICONERROR
      );
  end;
end;

procedure TFrmConsultaExecutantesProgramados.actSelAllRecolhimentoExecute(
  Sender: TObject);
begin
FrmPrincipal.selecaoDBGrid(DBGridExecutantesProgramados,true,'booleanRecolhimento');
end;

procedure TFrmConsultaExecutantesProgramados.actSelAllSelecaoExecute(
  Sender: TObject);
begin
FrmPrincipal.selecaoDBGrid(DBGridExecutantesProgramados,true,'booleanSelecao');
end;

procedure TFrmConsultaExecutantesProgramados.actSelClearRecolhimentoExecute(
  Sender: TObject);
begin
FrmPrincipal.selecaoDBGrid(DBGridExecutantesProgramados,false,'booleanRecolhimento');
end;

procedure TFrmConsultaExecutantesProgramados.actSelClearSelecaoExecute(
  Sender: TObject);
begin
FrmPrincipal.selecaoDBGrid(DBGridExecutantesProgramados,false,'booleanSelecao');
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
begin
  if not Assigned(FrmTabela) then
    Application.CreateForm(TFrmTabela, FrmTabela);
  FrmTabela.Show;   // ou ShowModal conforme o caso

  TProgramacaoRTUtils.PainelTransbordos(DateToStr(dataInicio.Date));
end;

procedure TFrmConsultaExecutantesProgramados.dataFimCloseUp(Sender: TObject);
begin
  actProcurarProgramacaoExecutante.Execute;
  actProcurarProgramacaoRT.Execute;
  actProcurarSAPImport.Execute;
  AtualizarDiasDisponiveisSimulacao;
end;

procedure TFrmConsultaExecutantesProgramados.dataFimKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.dataInicioCloseUp(Sender: TObject);
begin
  actProcurarProgramacaoExecutante.Execute;
  actProcurarProgramacaoRT.Execute;
  actProcurarSAPImport.Execute;
  AtualizarDiasDisponiveisSimulacao;
end;

procedure TFrmConsultaExecutantesProgramados.dataInicioKeyPress(Sender: TObject;
  var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

procedure TFrmConsultaExecutantesProgramados.DBEdit1Change(Sender: TObject);
begin
CarregarDetalheFrequencia;
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

procedure TFrmConsultaExecutantesProgramados.DBGridExecutantesProgramadosDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  if (Column.Field.FieldName = 'InseridoProgramacaoTransporte') then
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
  end
  else if (Column.Field.FieldName = 'RT_Status') then
  begin
    if SameText(Column.Field.AsString, RT_STATUS_EMITIDA) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clLime
    else if SameText(Column.Field.AsString, RT_STATUS_PENDENTE) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clYellow
    else if SameText(Column.Field.AsString, RT_STATUS_PRONTO_EMITIR) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clAqua
    else if SameText(Column.Field.AsString, RT_STATUS_NAO_CRIAR) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clSilver
    else if SameText(Column.Field.AsString, RT_STATUS_JA_EXISTE) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clSkyBlue
    else if SameText(Column.Field.AsString, RT_STATUS_CANCELADA) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clRed
    else if SameText(Column.Field.AsString, RT_STATUS_ORFA) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clYellow
    else if SameText(Column.Field.AsString, RT_STATUS_ERRO_NAO_ATIVO) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clRed
    else if SameText(Column.Field.AsString, RT_STATUS_ERRO_CONFLITO_RT) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clRed
    else if SameText(Column.Field.AsString, RT_STATUS_ERRO_EMISSAO) then
      DBGridExecutantesProgramados.Canvas.Brush.Color := clRed;

    DBGridExecutantesProgramados.Font.Color := clBlack;
    DBGridExecutantesProgramados.Canvas.FillRect(Rect);
    DBGridExecutantesProgramados.DefaultDrawColumnCell(Rect, DataCol, Column, State);
  end;
end;

procedure TFrmConsultaExecutantesProgramados.DBGridProgramcaoRTDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  with DBGridProgramcaoRT do
  begin
    if (Column.Field.FieldName = 'RT_Status') then
    begin
      if SameText(Column.Field.AsString, RT_STATUS_EMITIDA) then
        Canvas.Brush.Color := clLime
      else if SameText(Column.Field.AsString, RT_STATUS_PENDENTE) then
        Canvas.Brush.Color := clYellow
      else if SameText(Column.Field.AsString, RT_STATUS_PRONTO_EMITIR) then
        Canvas.Brush.Color := clAqua
      else if SameText(Column.Field.AsString, RT_STATUS_NAO_CRIAR) then
        Canvas.Brush.Color := clSilver
      else if SameText(Column.Field.AsString, RT_STATUS_JA_EXISTE) then
        Canvas.Brush.Color := clSkyBlue
      else if SameText(Column.Field.AsString, RT_STATUS_CANCELADA) then
        Canvas.Brush.Color := clRed
      else if SameText(Column.Field.AsString, RT_STATUS_ORFA) then
        Canvas.Brush.Color := clYellow
      else if SameText(Column.Field.AsString, RT_STATUS_ERRO_NAO_ATIVO) then
        Canvas.Brush.Color := clRed
      else if SameText(Column.Field.AsString, RT_STATUS_ERRO_CONFLITO_RT) then
        Canvas.Brush.Color := clRed
      else if SameText(Column.Field.AsString, RT_STATUS_ERRO_EMISSAO) then
        Canvas.Brush.Color := clRed;

      Font.Color := clBlack;
      Canvas.FillRect(Rect);
      DefaultDrawColumnCell(Rect, DataCol, Column, State);
    end;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.edtExecutanteKeyPress(
  Sender: TObject; var Key: Char);
begin
  if Key = #13 then
    actProcurarProgramacaoExecutante.Execute;
end;

function TFrmConsultaExecutantesProgramados.DeterminarModalAutomatico(
  const AOrigem, ADestino: string): string;
var
  PlataformaOrigem, PlataformaDestino: TDadosPlataforma;
  ModO, ModD: string;
begin
  PlataformaOrigem := TProgramacaoRTUtils.DadosPlataforma_RT(AOrigem);
  PlataformaDestino := TProgramacaoRTUtils.DadosPlataforma_RT(ADestino);

  ModO := UpperCase(Trim(PlataformaOrigem.RT_Modal));
  ModD := UpperCase(Trim(PlataformaDestino.RT_Modal));

  if ((ModO = 'T') and (ModD = 'M')) or
     ((ModO = 'M') and (ModD = 'T')) then
    Result := 'M'
  else if (ModO = 'A') or (ModD = 'A') then
    Result := 'A'
  else if (ModO = 'T') and (ModD = 'T') then
    Result := 'T'
  else
    Result := 'M';
end;

function TFrmConsultaExecutantesProgramados.CodigoSAPEhNumerico(
  const CodigoSAP: string): Boolean;
var
  I: Integer;
  S: string;
begin
  S := Trim(CodigoSAP);
  Result := S <> '';

  if not Result then
    Exit;

  for I := 1 to Length(S) do
    if not CharInSet(S[I], ['0'..'9']) then
      Exit(False);
end;

function TFrmConsultaExecutantesProgramados.BuscarNovoCodigoSAPExecutante(
  const ANomeExecutante, AEmpresa: string;
  out ACodigoSAP: string): Boolean;
var
  Q: TADOQuery;
  CodigoLido: string;
begin
  Result := False;
  ACodigoSAP := '';

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionConsulta;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'SELECT TOP 20 CodigoSAP ' +
      'FROM tblExecutante ' +
      'WHERE UCase(Trim(NomeExecutante)) = :pNome ' +
      '  AND UCase(Trim(txtEmpresa)) = :pEmpresa ' +
      '  AND CodigoSAP IS NOT NULL ' +
      '  AND Trim(CodigoSAP) <> '''' ' +
      'ORDER BY idExecutante DESC';

    Q.Parameters.ParamByName('pNome').DataType := ftString;
    Q.Parameters.ParamByName('pNome').Value := UpperCase(Trim(ANomeExecutante));

    Q.Parameters.ParamByName('pEmpresa').DataType := ftString;
    Q.Parameters.ParamByName('pEmpresa').Value := UpperCase(Trim(AEmpresa));

    Q.Open;

    while not Q.Eof do
    begin
      CodigoLido := Trim(Q.FieldByName('CodigoSAP').AsString);

      if CodigoSAPEhNumerico(CodigoLido) then
      begin
        ACodigoSAP := CodigoLido;
        Result := True;
        Break;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarCodigoSAPProgramacaoExecutante(
  const AIdProgramacaoExecutante: Integer;
  const ACodigoSAPNovo: string);
var
  Q: TADOQuery;
begin
  if (AIdProgramacaoExecutante <= 0) or (Trim(ACodigoSAPNovo) = '') then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante ' +
      'SET CodigoSAP = :pCodigoSAP ' +
      'WHERE idProgramacaoExecutante = :pId';

    Q.Parameters.ParamByName('pCodigoSAP').Value := Trim(ACodigoSAPNovo);
    Q.Parameters.ParamByName('pId').Value := AIdProgramacaoExecutante;
    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.TentarReclassificarExecutanteParaR3(
  const AIdProgramacaoExecutante: Integer;
  const ANomeExecutante, AEmpresa, ACodigoSAPAtual: string;
  out ACodigoSAPFinal, ATipoRTFinal: string): Boolean;
var
  NovoCodigoSAP: string;
begin
  Result := False;

  ACodigoSAPFinal := NormalizarCodigoSAP(ACodigoSAPAtual);
  ATipoRTFinal := 'R7';

  if CodigoSAPEhNumerico(ACodigoSAPFinal) then
  begin
    ATipoRTFinal := 'R3';

    if not SameText(Trim(ACodigoSAPAtual), Trim(ACodigoSAPFinal)) then
    begin
      AtualizarCodigoSAPProgramacaoExecutante(
        AIdProgramacaoExecutante,
        ACodigoSAPFinal
      );

      AtualizarCodigoSAPTblExecutante(
        ANomeExecutante,
        AEmpresa,
        ACodigoSAPFinal
      );
    end;

    Exit(False);
  end;

  if BuscarNovoCodigoSAPExecutante(ANomeExecutante, AEmpresa, NovoCodigoSAP) then
  begin
    ACodigoSAPFinal := NormalizarCodigoSAP(NovoCodigoSAP);
    ATipoRTFinal := 'R3';

    AtualizarCodigoSAPProgramacaoExecutante(
      AIdProgramacaoExecutante,
      ACodigoSAPFinal
    );

    AtualizarCodigoSAPTblExecutante(
      ANomeExecutante,
      AEmpresa,
      ACodigoSAPFinal
    );

    Result := True;
  end;
end;

function TFrmConsultaExecutantesProgramados.DeterminarTipoRTAutomatico(
  const CodigoSAP: string): string;
begin
  if CodigoSAPEhNumerico(NormalizarCodigoSAP(CodigoSAP)) then
    Result := 'R3'
  else
    Result := 'R7';
end;

function TFrmConsultaExecutantesProgramados.DeveCriarRTAutomaticamente(
  const AOrigem, ADestino: string;
  const PlataformaOrigem, PlataformaDestino: TDadosPlataforma
): Boolean;
begin
  Result :=
    (Trim(AOrigem) <> '') and
    (Trim(ADestino) <> '') and
    (AOrigem <> ADestino) and
    (not PlataformaOrigem.booleanNaoCriarRT) and
    (not PlataformaDestino.booleanNaoCriarRT);
end;

procedure TFrmConsultaExecutantesProgramados.AplicarRecolhimentoNoRegistro(
  DS: TDataSet;
  const Horario: THorario;
  const AOrigem: string;
  const ABooleanRecolhimento: Boolean);
begin
  if ABooleanRecolhimento then
  begin
    if Trim(DS.FieldByName('RecolherPara').AsString) = '' then
      DS.FieldByName('RecolherPara').AsString := AOrigem;

    if Trim(DS.FieldByName('RT_HoraVolta').AsString) = '' then
      DS.FieldByName('RT_HoraVolta').AsString := Horario.HoraVolta;

    if VarIsNull(DS.FieldByName('DataVolta').Value) then
      DS.FieldByName('DataVolta').AsDateTime := Horario.DataVolta;
  end
  else
  begin
    DS.FieldByName('RecolherPara').AsString := '';
    DS.FieldByName('RT_HoraVolta').AsString := '';
    DS.FieldByName('DataVolta').Clear;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.PreencherCamposAutomaticosRTNoRegistro(
  DS: TDataSet;
  const TipoRT, Modal, Classe: string;
  const Horario: THorario);
begin
  if Trim(DS.FieldByName('RT_Tipo').AsString) <> TipoRT then
    DS.FieldByName('RT_Tipo').AsString := TipoRT;

  if Trim(DS.FieldByName('RT_Modal').AsString) = '' then
    DS.FieldByName('RT_Modal').AsString := Modal;

  if Trim(DS.FieldByName('RT_Classe').AsString) = '' then
    DS.FieldByName('RT_Classe').AsString := Classe;

  if Trim(DS.FieldByName('RT_HoraIda').AsString) = '' then
    DS.FieldByName('RT_HoraIda').AsString := Trim(Horario.HoraIda);
end;

procedure TFrmConsultaExecutantesProgramados.ProcessarRetornoCancelamentoSAP(const LinhaLog: string);
var
  idProgRT: Integer;
  Acao, Status, Msg, MsgNormalizada: string;
  p1, p2, p3: Integer;
  CancelamentoEfetivado: Boolean;
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

  CancelamentoEfetivado :=
    SameText(Trim(Status), 'FINALIZADO') or
    SameText(Trim(Status), 'OK');

  if not CancelamentoEfetivado then
  begin
    MsgNormalizada := UpperCase(NormalizarMensagemLog(Msg));
    CancelamentoEfetivado :=
      (Pos('ENCERRAD', MsgNormalizada) > 0) or
      (Pos('CANCELAD', MsgNormalizada) > 0);
  end;

  if CancelamentoEfetivado then
    AtualizarProgramacaoRT_Cancelamento(idProgRT, True, Msg, RT_STATUS_CANCELADA)
  else
    AtualizarProgramacaoRT_Cancelamento(idProgRT, False, Msg, '');
end;

procedure TFrmConsultaExecutantesProgramados.ImportarRTsSAPPeriodoPorTipo(
  const ATipoRT: string;
  const APeriodoInicio, APeriodoFim: TDateTime);
var
  EnderecoVbs, EnderecoLog: string;
  LogRetorno: TStringList;
  I: Integer;

  function VbsQuote(const S: string): string;
  begin
    Result := '"' + StringReplace(S, '"', '""', [rfReplaceAll]) + '"';
  end;

  procedure MemoAdd(const S: string);
  begin
    MemoSAP.Lines.Add(S);
  end;

begin
  EnderecoVbs := ExtractFilePath(ParamStr(0)) + 'RT_Import_' + UpperCase(ATipoRT) + '.vbs';
  EnderecoLog := ExtractFilePath(ParamStr(0)) + 'sap_log_import_' + LowerCase(ATipoRT) + '.txt';

  MemoSAP.Lines.Clear;

  MemoAdd('Option Explicit');
  MemoAdd('');
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

  MemoAdd('Function San(ByVal v)');
  MemoAdd('  v = CStr(v)');
  MemoAdd('  v = Replace(v, "|", "/")');
  MemoAdd('  v = Replace(v, vbCr, " ")');
  MemoAdd('  v = Replace(v, vbLf, " ")');
  MemoAdd('  San = Trim(v)');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function GridCell(ByVal grid, ByVal row, ByVal colId)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  GridCell = ""');
  MemoAdd('  GridCell = grid.GetCellValue(row, colId)');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    GridCell = ""');
  MemoAdd('  End If');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function GridCellAny(ByVal grid, ByVal row, ByVal arr)');
  MemoAdd('  Dim i, v');
  MemoAdd('  GridCellAny = ""');
  MemoAdd('  For i = 0 To UBound(arr)');
  MemoAdd('    v = Trim(CStr(GridCell(grid, row, CStr(arr(i)))))');
  MemoAdd('    If Len(v) > 0 Then');
  MemoAdd('      GridCellAny = v');
  MemoAdd('      Exit Function');
  MemoAdd('    End If');
  MemoAdd('  Next');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function PropagarRepetido(ByVal atual, ByRef anterior)');
  MemoAdd('  atual = Trim(CStr(atual))');
  MemoAdd('  If Len(atual) = 0 Then');
  MemoAdd('    PropagarRepetido = Trim(CStr(anterior))');
  MemoAdd('  Else');
  MemoAdd('    anterior = atual');
  MemoAdd('    PropagarRepetido = atual');
  MemoAdd('  End If');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function InicializarSAP()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Dim i, j');
  MemoAdd('  Dim connTmp, sessTmp');
  MemoAdd('  InicializarSAP = False');
  MemoAdd('');
  MemoAdd('  Set SapGuiAuto = Nothing');
  MemoAdd('  Set application = Nothing');
  MemoAdd('  Set connection = Nothing');
  MemoAdd('  Set session = Nothing');
  MemoAdd('');
  MemoAdd('  Set SapGuiAuto = GetObject("SAPGUI")');
  MemoAdd('  If Err.Number <> 0 Or (SapGuiAuto Is Nothing) Then');
  MemoAdd('    LogLine "ERRO|SAPGUI_NAO_ENCONTRADO|" & Err.Number & "|" & Err.Description');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  Set application = SapGuiAuto.GetScriptingEngine');
  MemoAdd('  If Err.Number <> 0 Or (application Is Nothing) Then');
  MemoAdd('    LogLine "ERRO|SCRIPT_ENGINE_NAO_ENCONTRADO|" & Err.Number & "|" & Err.Description');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  LogLine "DEBUG|APPLICATION_CHILDREN=" & application.Children.Count');
  MemoAdd('');
  MemoAdd('  If application.Children.Count <= 0 Then');
  MemoAdd('    LogLine "ERRO|SEM_CONEXAO_SAP"');
  MemoAdd('    Err.Clear');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  For i = 0 To application.Children.Count - 1');
  MemoAdd('    Set connTmp = Nothing');
  MemoAdd('    Set connTmp = application.Children(CLng(i))');
  MemoAdd('    If Err.Number <> 0 Then');
  MemoAdd('      LogLine "DEBUG|ERRO_LENDO_CONEXAO|" & i & "|" & Err.Number & "|" & Err.Description');
  MemoAdd('      Err.Clear');
  MemoAdd('    ElseIf Not (connTmp Is Nothing) Then');
  MemoAdd('      LogLine "DEBUG|CONEXAO|" & i & "|SESSOES=" & connTmp.Children.Count');
  MemoAdd('      If connTmp.Children.Count > 0 Then');
  MemoAdd('        For j = 0 To connTmp.Children.Count - 1');
  MemoAdd('          Set sessTmp = Nothing');
  MemoAdd('          Set sessTmp = connTmp.Children(CLng(j))');
  MemoAdd('          If Err.Number <> 0 Then');
  MemoAdd('            LogLine "DEBUG|ERRO_LENDO_SESSAO|" & i & "|" & j & "|" & Err.Number & "|" & Err.Description');
  MemoAdd('            Err.Clear');
  MemoAdd('          ElseIf Not (sessTmp Is Nothing) Then');
  MemoAdd('            Set connection = connTmp');
  MemoAdd('            Set session = sessTmp');
  MemoAdd('            InicializarSAP = True');
  MemoAdd('            LogLine "DEBUG|SESSAO_SELECIONADA|CONEXAO=" & i & "|SESSAO=" & j');
  MemoAdd('            On Error GoTo 0');
  MemoAdd('            Exit Function');
  MemoAdd('          End If');
  MemoAdd('        Next');
  MemoAdd('      End If');
  MemoAdd('    End If');
  MemoAdd('  Next');
  MemoAdd('');
  MemoAdd('  LogLine "ERRO|SEM_SESSAO_SAP"');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function ObterGridResultado()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  Set ObterGridResultado = Nothing');
  MemoAdd('  Set ObterGridResultado = session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell")');
  MemoAdd('  If Err.Number = 0 Then');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    Exit Function');
  MemoAdd('  End If');
  MemoAdd('  Err.Clear');
  MemoAdd('  Set ObterGridResultado = Nothing');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Function GetVisibleRows(ByVal grid)');
  MemoAdd('  On Error Resume Next');
  MemoAdd('  GetVisibleRows = grid.VisibleRowCount');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    Err.Clear');
  MemoAdd('    GetVisibleRows = 20');
  MemoAdd('  End If');
  MemoAdd('  If GetVisibleRows <= 0 Then GetVisibleRows = 20');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Function');
  MemoAdd('');

  MemoAdd('Sub ExportarGrid()');
  MemoAdd('  On Error Resume Next');
  MemoAdd('');
  MemoAdd('  Dim grid, n, vis, ini, fim, i');
  MemoAdd('  Dim qmnum, qmart, iwerk, ingrp, origem, destino, dtviagem, hviagem');
  MemoAdd('  Dim pernr, tipodoc, documento, passageiro, qmtxt, modal, classe, stitem');
  MemoAdd('');
  MemoAdd('  Dim lastQmnum, lastQmart, lastIwerk, lastIngrp, lastOrigem, lastDestino');
  MemoAdd('  Dim lastDtviagem, lastHviagem, lastPernr, lastTipoDoc, lastDocumento');
  MemoAdd('  Dim lastPassageiro, lastQmtxt, lastModal, lastClasse, lastStitem');
  MemoAdd('');
  MemoAdd('  lastQmnum = ""');
  MemoAdd('  lastQmart = ""');
  MemoAdd('  lastIwerk = ""');
  MemoAdd('  lastIngrp = ""');
  MemoAdd('  lastOrigem = ""');
  MemoAdd('  lastDestino = ""');
  MemoAdd('  lastDtviagem = ""');
  MemoAdd('  lastHviagem = ""');
  MemoAdd('  lastPernr = ""');
  MemoAdd('  lastTipoDoc = ""');
  MemoAdd('  lastDocumento = ""');
  MemoAdd('  lastPassageiro = ""');
  MemoAdd('  lastQmtxt = ""');
  MemoAdd('  lastModal = ""');
  MemoAdd('  lastClasse = ""');
  MemoAdd('  lastStitem = ""');
  MemoAdd('');
  MemoAdd('  Set grid = Nothing');
  MemoAdd('  Set grid = ObterGridResultado()');
  MemoAdd('');
  MemoAdd('  If grid Is Nothing Then');
  MemoAdd('    LogLine "ERRO|GRID_NAO_ENCONTRADO|' + UpperCase(ATipoRT) + '"');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  n = 0');
  MemoAdd('  n = grid.RowCount');
  MemoAdd('  If Err.Number <> 0 Then');
  MemoAdd('    LogLine "ERRO|ROWCOUNT|' + UpperCase(ATipoRT) + '|" & Err.Number & "|" & Err.Description');
  MemoAdd('    Err.Clear');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  vis = GetVisibleRows(grid)');
  MemoAdd('');
  MemoAdd('  LogLine "IMPORT_INFO|TIPO=' + UpperCase(ATipoRT) + '|ROWCOUNT=" & n & "|VISIBLE=" & vis');
  MemoAdd('');
  MemoAdd('  If n <= 0 Then');
  MemoAdd('    LogLine "IMPORT_SEM_RESULTADO|TIPO=' + UpperCase(ATipoRT) + '"');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('    Exit Sub');
  MemoAdd('  End If');
  MemoAdd('');
  MemoAdd('  ini = 0');
  MemoAdd('  Do While ini < n');
  MemoAdd('    On Error Resume Next');
  MemoAdd('    grid.firstVisibleRow = ini');
  MemoAdd('    Err.Clear');
  MemoAdd('    WScript.Sleep 100');
  MemoAdd('    On Error GoTo 0');
  MemoAdd('');
  MemoAdd('    fim = ini + vis - 1');
  MemoAdd('    If fim >= n Then fim = n - 1');
  MemoAdd('');
  MemoAdd('    For i = ini To fim');
  MemoAdd('      qmnum      = PropagarRepetido(GridCellAny(grid, i, Array("QMNUM")), lastQmnum)');
  MemoAdd('      qmart      = PropagarRepetido(GridCellAny(grid, i, Array("QMART")), lastQmart)');
  MemoAdd('      iwerk      = PropagarRepetido(GridCellAny(grid, i, Array("IWERK")), lastIwerk)');
  MemoAdd('      ingrp      = PropagarRepetido(GridCellAny(grid, i, Array("INGRP")), lastIngrp)');
  MemoAdd('      origem     = PropagarRepetido(GridCellAny(grid, i, Array("ORIGEM","LOCAL_ORIGEM")), lastOrigem)');
  MemoAdd('      destino    = PropagarRepetido(GridCellAny(grid, i, Array("DESTINO","LOCAL_DESTINO")), lastDestino)');
  MemoAdd('      dtviagem   = PropagarRepetido(GridCellAny(grid, i, Array("DTVIAGEM","DATA_VIAGEM")), lastDtviagem)');
  MemoAdd('      hviagem    = PropagarRepetido(GridCellAny(grid, i, Array("HVIAGEM","HORA_VIAGEM")), lastHviagem)');
  MemoAdd('      pernr      = PropagarRepetido(GridCellAny(grid, i, Array("PERNR")), lastPernr)');
  MemoAdd('      tipodoc    = PropagarRepetido(GridCellAny(grid, i, Array("TIPO_DOC")), lastTipoDoc)');
  MemoAdd('      documento  = PropagarRepetido(GridCellAny(grid, i, Array("DOCUMENT","DOCUMENTO")), lastDocumento)');
  MemoAdd('      passageiro = PropagarRepetido(GridCellAny(grid, i, Array("PASSAGEIRO")), lastPassageiro)');
  MemoAdd('      qmtxt      = PropagarRepetido(GridCellAny(grid, i, Array("QMTXT")), lastQmtxt)');
  MemoAdd('      modal      = PropagarRepetido(GridCellAny(grid, i, Array("CODTRASG","MODAL")), lastModal)');
  MemoAdd('      classe     = PropagarRepetido(GridCellAny(grid, i, Array("CLASSEPAS","CLASSE")), lastClasse)');
  MemoAdd('      stitem     = PropagarRepetido(GridCellAny(grid, i, Array("STATUS")), lastStitem)');
  MemoAdd('');
  MemoAdd('      LogLine "IMPORT|" & _');
  MemoAdd('        San(qmnum)      & "|" & _');
  MemoAdd('        San(qmart)      & "|" & _');
  MemoAdd('        San(iwerk)      & "|" & _');
  MemoAdd('        San(ingrp)      & "|" & _');
  MemoAdd('        San(origem)     & "|" & _');
  MemoAdd('        San(destino)    & "|" & _');
  MemoAdd('        San(dtviagem)   & "|" & _');
  MemoAdd('        San(hviagem)    & "|" & _');
  MemoAdd('        San(pernr)      & "|" & _');
  MemoAdd('        San(tipodoc)    & "|" & _');
  MemoAdd('        San(documento)  & "|" & _');
  MemoAdd('        San(passageiro) & "|" & _');
  MemoAdd('        San(qmtxt)      & "|" & _');
  MemoAdd('        San(modal)      & "|" & _');
  MemoAdd('        San(classe)     & "|" & _');
  MemoAdd('        San(stitem)');
  MemoAdd('    Next');
  MemoAdd('');
  MemoAdd('    ini = fim + 1');
  MemoAdd('  Loop');
  MemoAdd('');
  MemoAdd('  On Error GoTo 0');
  MemoAdd('End Sub');
  MemoAdd('');

  MemoAdd('LogLine "INICIO_IMPORT|TIPO=' + UpperCase(ATipoRT) + '"');
  MemoAdd('');
  MemoAdd('If Not InicializarSAP() Then');
  MemoAdd('  LogLine "FIM_IMPORT|TIPO=' + UpperCase(ATipoRT) + '|ERRO_INICIALIZACAO"');
  MemoAdd('  WScript.Quit');
  MemoAdd('End If');
  MemoAdd('');
  MemoAdd('On Error Resume Next');
  MemoAdd('session.findById("wnd[0]").maximize');
  MemoAdd('If Err.Number <> 0 Then');
  MemoAdd('  LogLine "DEBUG|MAXIMIZE_FALHOU|" & Err.Number & "|" & Err.Description');
  MemoAdd('  Err.Clear');
  MemoAdd('End If');
  MemoAdd('On Error GoTo 0');
  MemoAdd('');
  MemoAdd('session.findById("wnd[0]/tbar[0]/okcd").text = "/nYSCSXFRT"');
  MemoAdd('session.findById("wnd[0]").sendVKey 0');
  MemoAdd('WScript.Sleep 500');
  MemoAdd('');
  MemoAdd('On Error Resume Next');
  MemoAdd('session.findById("wnd[0]/usr/chkDY_MAB").selected = True');
  MemoAdd('If Err.Number <> 0 Then');
  MemoAdd('  LogLine "DEBUG|CHK_DY_MAB_NAO_ENCONTRADO|" & Err.Number & "|" & Err.Description');
  MemoAdd('  Err.Clear');
  MemoAdd('End If');
  MemoAdd('On Error GoTo 0');
  MemoAdd('');
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0100/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0100/ctxtIWERK-LOW").text = ' +
    VbsQuote(Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_CentroPlan').AsString)));
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0100/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0100/ctxtINGRP-LOW").text = ' +
    VbsQuote(Trim(FrmDataModule.ADOQueryConfigRT.FieldByName('RT_GrpPlan').AsString)));
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0100/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0100/ctxtQMART").text = ' +
    VbsQuote(UpperCase(ATipoRT)));
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0100/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0100/chkP_NOCOBE").selected = false');
  MemoAdd('');
  MemoAdd('On Error Resume Next');
  MemoAdd('session.findById("wnd[0]/usr/ctxtPC_VARI").text = "/colibri"');
  MemoAdd('If Err.Number <> 0 Then');
  MemoAdd('  LogLine "DEBUG|VARIANTE_NAO_APLICADA|" & Err.Number & "|" & Err.Description');
  MemoAdd('  Err.Clear');
  MemoAdd('End If');
  MemoAdd('On Error GoTo 0');
  MemoAdd('');
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0400").select');
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0400/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0400/ctxtDTVIAGEM-LOW").text = ' +
    VbsQuote(FormatDateTime('dd.mm.yyyy', APeriodoInicio)));
  MemoAdd('On Error Resume Next');
  MemoAdd('session.findById("wnd[0]/usr/tabsTABSTRIP_TAB_PAR/tabpBTN_SCREEN_0400/ssub%_SUBSCREEN_TAB_PAR:YSRPP_FILTRO_RT:0400/ctxtDTVIAGEM-HIGH").text = ' +
    VbsQuote(FormatDateTime('dd.mm.yyyy', APeriodoFim)));
  MemoAdd('If Err.Number <> 0 Then');
  MemoAdd('  LogLine "DEBUG|CAMPO_DATA_HIGH_NAO_ENCONTRADO|" & Err.Number & "|" & Err.Description');
  MemoAdd('  Err.Clear');
  MemoAdd('End If');
  MemoAdd('On Error GoTo 0');
  MemoAdd('');
  MemoAdd('session.findById("wnd[0]/tbar[1]/btn[8]").press');
  MemoAdd('WScript.Sleep 1000');
  MemoAdd('Call ExportarGrid()');
  MemoAdd('LogLine "FIM_IMPORT|TIPO=' + UpperCase(ATipoRT) + '"');

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
        for I := 0 to LogRetorno.Count - 1 do
          ProcessarRetornoImportacaoRTSAP(
            LogRetorno[I],
            'PERIODO',
            APeriodoInicio,
            APeriodoFim
          );
      finally
        LogRetorno.Free;
      end;
    end
    else
      ShowMessage('Script terminou, mas não encontrei o log: ' + EnderecoLog);
  end
  else
    ShowMessage('Erro ao iniciar o script VBS de importação SAP.');
end;

function TFrmConsultaExecutantesProgramados.BuscarProgramacaoRTPorNumeroRT(
  const ART: string;
  out AIdProgramacaoRT, AIdProgramacaoExecutante: Integer): Integer;
var
  Q: TADOQuery;
begin
  Result := 0;
  AIdProgramacaoRT := 0;
  AIdProgramacaoExecutante := 0;

  if Trim(ART) = '' then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'SELECT TOP 2 ' +
      '  idProgramacaoRT, ' +
      '  idProgramacaoExecutante ' +
      'FROM tblProgramacaoRT ' +
      'WHERE ([RT] IS NOT NULL) ' +
      '  AND (Trim([RT]) = :RT) ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      'ORDER BY idProgramacaoRT DESC';

    Q.Parameters.ParamByName('RT').Value := Trim(ART);
    Q.Open;

    while not Q.Eof do
    begin
      Inc(Result);

      if Result = 1 then
      begin
        AIdProgramacaoRT := Q.FieldByName('idProgramacaoRT').AsInteger;

        if not VarIsNull(Q.FieldByName('idProgramacaoExecutante').Value) then
          AIdProgramacaoExecutante := Q.FieldByName('idProgramacaoExecutante').AsInteger
        else
          AIdProgramacaoExecutante := 0;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarProgramacaoRTPorChave(
  const ACampoChave, AValorChave: string;
  out AIdProgramacaoRT, AIdProgramacaoExecutante: Integer): Integer;
var
  Q: TADOQuery;
begin
  Result := 0;
  AIdProgramacaoRT := 0;
  AIdProgramacaoExecutante := 0;

  if (Trim(ACampoChave) = '') or (Trim(AValorChave) = '') then
    Exit;

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.SQL.Text :=
      'SELECT TOP 2 ' +
      '  idProgramacaoRT, ' +
      '  idProgramacaoExecutante ' +
      'FROM tblProgramacaoRT ' +
      'WHERE ([' + ACampoChave + '] IS NOT NULL) ' +
      '  AND Trim([' + ACampoChave + ']) = :VALOR ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      'ORDER BY idProgramacaoRT DESC';

    Q.Parameters.ParamByName('VALOR').Value := Trim(AValorChave);
    Q.Open;

    while not Q.Eof do
    begin
      Inc(Result);

      if Result = 1 then
      begin
        AIdProgramacaoRT := Q.FieldByName('idProgramacaoRT').AsInteger;

        if not VarIsNull(Q.FieldByName('idProgramacaoExecutante').Value) then
          AIdProgramacaoExecutante := Q.FieldByName('idProgramacaoExecutante').AsInteger
        else
          AIdProgramacaoExecutante := 0;
      end;

      Q.Next;
    end;
  finally
    Q.Free;
  end;
end;


procedure TFrmConsultaExecutantesProgramados.AtualizarVinculoRTSapImport(
  const AIdRTSapImport, AIdProgramacaoRT, AIdProgramacaoExecutante: Integer;
  const AImportadoColibri: Boolean;
  const AObservacao: string);
var
  Q: TADOQuery;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;
    Q.SQL.Clear;
    Q.SQL.Text :=
      'UPDATE tblRTSapImport SET ' +
      '  idProgramacaoRT = :idProgramacaoRT, ' +
      '  idProgramacaoExecutante = :idProgramacaoExecutante, ' +
      '  ImportadoColibri = :ImportadoColibri, ' +
      '  Observacao = :Observacao ' +
      'WHERE idRTSapImport = :idRTSapImport';

    with Q.Parameters.ParamByName('idProgramacaoRT') do
    begin
      DataType := ftInteger;
      Direction := pdInput;
      if AIdProgramacaoRT > 0 then
        Value := AIdProgramacaoRT
      else
        Value := Null;
    end;

    with Q.Parameters.ParamByName('idProgramacaoExecutante') do
    begin
      DataType := ftInteger;
      Direction := pdInput;
      if AIdProgramacaoExecutante > 0 then
        Value := AIdProgramacaoExecutante
      else
        Value := Null;
    end;

    with Q.Parameters.ParamByName('ImportadoColibri') do
    begin
      DataType := ftBoolean;
      Direction := pdInput;
      Value := AImportadoColibri;
    end;

    with Q.Parameters.ParamByName('Observacao') do
    begin
      DataType := ftString;
      Size := 255;
      Direction := pdInput;
      Value := Copy(Trim(AObservacao), 1, 255);
    end;

    with Q.Parameters.ParamByName('idRTSapImport') do
    begin
      DataType := ftInteger;
      Direction := pdInput;
      Value := AIdRTSapImport;
    end;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoRTViaImportacaoSAP(
  const AIdProgramacaoRT: Integer;
  const ARTNumero, AStatusItem, AStatusDescricao, AOrigemVinculo: string);
var
  Q: TADOQuery;
  Msg, StatusUnico: string;
  CanceladaSAP: Boolean;
begin
  if AIdProgramacaoRT <= 0 then
    Exit;

  StatusUnico := StatusImportacaoSAP(ARTNumero, AStatusItem, AStatusDescricao);
  CanceladaSAP := SameText(StatusUnico, RT_STATUS_CANCELADA);

  Msg := Copy(
    Trim('Importado do SAP por ' + AOrigemVinculo +
         ' - Status ' + Trim(AStatusItem) + ' ' + Trim(AStatusDescricao)),
    1, 255
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoRT SET ' +
      '  [RT] = :RT, ' +
      '  RT_Status = :RT_Status, ' +
      '  RT_Mensagem = :RT_Mensagem, ' +
      '  RT_Cancelada = :RT_Cancelada ' +
      'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.Parameters.ParamByName('RT').Value := Trim(ARTNumero);
    Q.Parameters.ParamByName('RT_Status').Value :=
      AjustarTextoParaCampo(
        FrmDataModule.ADOConnectionRT,
        'tblProgramacaoRT',
        'RT_Status',
        StatusUnico,
        40
      );
    Q.Parameters.ParamByName('RT_Mensagem').Value :=
      AjustarTextoParaCampo(
        FrmDataModule.ADOConnectionRT,
        'tblProgramacaoRT',
        'RT_Mensagem',
        Msg,
        255
      );
    Q.Parameters.ParamByName('RT_Cancelada').Value := CanceladaSAP;
    Q.Parameters.ParamByName('idProgramacaoRT').Value := AIdProgramacaoRT;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoExecutanteViaImportacaoSAP(
  const AIdProgramacaoExecutante: Integer;
  const ARTNumero, AStatusItem, AStatusDescricao, AOrigemVinculo: string);
var
  Q: TADOQuery;
  Msg, StatusUnico: string;
begin
  if AIdProgramacaoExecutante <= 0 then
    Exit;

  StatusUnico := StatusImportacaoSAP(ARTNumero, AStatusItem, AStatusDescricao);

  Msg := Copy(
    Trim('Importado do SAP por ' + AOrigemVinculo +
         ' - Status ' + Trim(AStatusItem) + ' ' + Trim(AStatusDescricao)),
    1, 255
  );

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionColibri;
    Q.ParamCheck := True;
    Q.SQL.Text :=
      'UPDATE tblProgramacaoExecutante SET ' +
      '  [RT] = :RT, ' +
      '  RT_Status = :RT_Status, ' +
      '  RT_Mensagem = :RT_Mensagem ' +
      'WHERE idProgramacaoExecutante = :idProgramacaoExecutante';

    Q.Parameters.ParamByName('RT').Value := Trim(ARTNumero);
    Q.Parameters.ParamByName('RT_Status').Value :=
      AjustarTextoParaCampo(
        FrmDataModule.ADOConnectionColibri,
        'tblProgramacaoExecutante',
        'RT_Status',
        StatusUnico,
        40
      );
    Q.Parameters.ParamByName('RT_Mensagem').Value :=
      AjustarTextoParaCampo(
        FrmDataModule.ADOConnectionColibri,
        'tblProgramacaoExecutante',
        'RT_Mensagem',
        Msg,
        255
      );
    Q.Parameters.ParamByName('idProgramacaoExecutante').Value := AIdProgramacaoExecutante;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.VincularRTSapImportComProgramacao(
  const APeriodoInicio, APeriodoFim: TDateTime);
var
  DadosSAP: TRTSapImportDados;
  IdExecDireto, IdRTCriado: Integer;
  QImp: TADOQuery;
  IdRTSapImport: Integer;
  IdProgramacaoRT, IdProgramacaoExecutante: Integer;
  QtdeMatch: Integer;
  QMNUM, ChaveImport, StatusItem, StatusDescricao: string;
  OrigemVinculo, Observacao: string;
  DtIni, DtFim: TDateTime;
  QtVinculados, QtSemMatch, QtMultiplos: Integer;
begin
  DtIni := Trunc(APeriodoInicio);
  DtFim := Trunc(APeriodoFim) + EncodeTime(23, 59, 59, 0);

  QtVinculados := 0;
  QtSemMatch := 0;
  QtMultiplos := 0;

  QImp := TADOQuery.Create(nil);
  try
    QImp.Connection := FrmDataModule.ADOConnectionRT;
    QImp.SQL.Text :=
      'SELECT tblRTSapImport.* ' +
      'FROM tblRTSapImport ' +
      'WHERE DataViagem BETWEEN :DtIni AND :DtFim ' +
      'ORDER BY DataViagem, QMNUM, idRTSapImport';

    QImp.Parameters.ParamByName('DtIni').Value := DtIni;
    QImp.Parameters.ParamByName('DtFim').Value := DtFim;
    QImp.Open;

    FrmPrincipal.ProgressBarIncializa(QImp.RecordCount, 'Vinculando importação SAP...');

    while not QImp.Eof do
    begin
      IdRTSapImport := QImp.FieldByName('idRTSapImport').AsInteger;
      QMNUM := Trim(QImp.FieldByName('QMNUM').AsString);
      ChaveImport := Trim(QImp.FieldByName('ChaveIda').AsString);
      StatusItem := Trim(QImp.FieldByName('StatusItem').AsString);
      StatusDescricao := Trim(QImp.FieldByName('StatusDescricao').AsString);


      DadosSAP := Default(TRTSapImportDados);
      DadosSAP.idRTSapImport := QImp.FieldByName('idRTSapImport').AsInteger;
      DadosSAP.QMNUM := Trim(QImp.FieldByName('QMNUM').AsString);
      DadosSAP.QMART := Trim(QImp.FieldByName('QMART').AsString);
      DadosSAP.IWERK := Trim(QImp.FieldByName('IWERK').AsString);
      DadosSAP.INGRP := Trim(QImp.FieldByName('INGRP').AsString);
      DadosSAP.Origem := Trim(QImp.FieldByName('Origem').AsString);
      DadosSAP.Destino := Trim(QImp.FieldByName('txtDestino').AsString);
      DadosSAP.DataViagem := QImp.FieldByName('DataViagem').AsDateTime;
      DadosSAP.HoraViagem := Trim(QImp.FieldByName('HoraViagem').AsString);
      DadosSAP.PERNR := Trim(QImp.FieldByName('PERNR').AsString);
      DadosSAP.TipoDoc := Trim(QImp.FieldByName('TipoDoc').AsString);
      DadosSAP.Documento := Trim(QImp.FieldByName('Documento').AsString);
      DadosSAP.Passageiro := Trim(QImp.FieldByName('Passageiro').AsString);
      DadosSAP.QMTXT := Trim(QImp.FieldByName('QMTXT').AsString);
      DadosSAP.RT_Modal := Trim(QImp.FieldByName('RT_Modal').AsString);
      DadosSAP.RT_Classe := Trim(QImp.FieldByName('RT_Classe').AsString);
      DadosSAP.StatusItem := Trim(QImp.FieldByName('StatusItem').AsString);
      DadosSAP.StatusDescricao := Trim(QImp.FieldByName('StatusDescricao').AsString);

      IdProgramacaoRT := 0;
      IdProgramacaoExecutante := 0;
      OrigemVinculo := '';
      Observacao := '';

      // 1) tenta pelo número da RT
      QtdeMatch := BuscarProgramacaoRTPorNumeroRT(
        QMNUM, IdProgramacaoRT, IdProgramacaoExecutante);

      if QtdeMatch = 1 then
      begin
        OrigemVinculo := 'NUMERO_RT';
      end
      else if QtdeMatch > 1 then
      begin
        Observacao := 'Mais de uma correspondência por número da RT';
      end
      else
      begin
        // 2) tenta pela chave de ida
        QtdeMatch := BuscarProgramacaoRTPorChave(
          'ChaveIda', ChaveImport, IdProgramacaoRT, IdProgramacaoExecutante);

        if QtdeMatch = 1 then
        begin
          OrigemVinculo := 'CHAVE_IDA';
        end
        else if QtdeMatch > 1 then
        begin
          Observacao := 'Mais de uma correspondência por ChaveIda';
          IdProgramacaoRT := 0;
          IdProgramacaoExecutante := 0;
        end
        else
        begin
          // 3) tenta pela chave de volta
          QtdeMatch := BuscarProgramacaoRTPorChave(
            'ChaveVolta', ChaveImport, IdProgramacaoRT, IdProgramacaoExecutante);

          if QtdeMatch = 1 then
          begin
            OrigemVinculo := 'CHAVE_VOLTA';
          end
          else if QtdeMatch > 1 then
          begin
            Observacao := 'Mais de uma correspondência por ChaveVolta';
            IdProgramacaoRT := 0;
            IdProgramacaoExecutante := 0;
          end
          else
          begin
            Observacao := 'Sem correspondência na tblProgramacaoRT';
          end;
        end;
      end;

      if IdProgramacaoRT > 0 then
      begin

        AtualizarVinculoRTSapImport(
          IdRTSapImport,
          IdProgramacaoRT,
          IdProgramacaoExecutante,
          True,
          'Vinculado por ' + OrigemVinculo
        );

        AtualizarProgramacaoRTViaImportacaoSAP(
          IdProgramacaoRT,
          QMNUM,
          StatusItem,
          StatusDescricao,
          OrigemVinculo
        );

        AtualizarProgramacaoExecutanteViaImportacaoSAP(
          IdProgramacaoExecutante,
          QMNUM,
          StatusItem,
          StatusDescricao,
          OrigemVinculo
        );

        Inc(QtVinculados);
      end
      else
      begin
        IdExecDireto := 0;

        if (Trim(StatusItem) <> '09') and
           BuscarProgramacaoExecutantePorImportacaoSAP(DadosSAP, IdExecDireto) then
        begin
          IdRTCriado := CriarRTLocalViaImportacaoSAP(DadosSAP, IdExecDireto);

          if IdRTCriado > 0 then
          begin
            AtualizarVinculoRTSapImport(
              IdRTSapImport,
              IdRTCriado,
              IdExecDireto,
              True,
              'RT local criada a partir da importação SAP'
            );

            AtualizarProgramacaoRTViaImportacaoSAP(
              IdRTCriado,
              QMNUM,
              StatusItem,
              StatusDescricao,
              'IMPORTACAO_SAP'
            );

            AtualizarProgramacaoExecutanteViaImportacaoSAP(
              IdExecDireto,
              QMNUM,
              StatusItem,
              StatusDescricao,
              'IMPORTACAO_SAP'
            );

            Inc(QtVinculados);
          end
          else
          begin
            AtualizarVinculoRTSapImport(
              IdRTSapImport,
              0,
              IdExecDireto,
              False,
              'Executante localizado, mas falhou ao criar tblProgramacaoRT'
            );
            Inc(QtSemMatch);
          end;
        end
        else
        begin
          AtualizarVinculoRTSapImport(
            IdRTSapImport,
            0,
            0,
            False,
            Observacao
          );

          if Pos('Mais de uma', Observacao) > 0 then
            Inc(QtMultiplos)
          else
            Inc(QtSemMatch);
        end;
      end;

      QImp.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

    FrmPrincipal.ProgressBarAtualizar;

    ShowMessage(
      'Vinculação concluída.' + sLineBreak +
      'Vinculados: ' + IntToStr(QtVinculados) + sLineBreak +
      'Sem correspondência: ' + IntToStr(QtSemMatch) + sLineBreak +
      'Múltiplos candidatos: ' + IntToStr(QtMultiplos)
    );

  finally
    QImp.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarProgramacaoRT_Cancelamento(
  const idProgramacaoRT: Integer;
  const Cancelada: Boolean;
  const Mensagem, RT_Status: string);
var
  Q: TADOQuery;
  SQL, StatusSafe, MensagemSafe: string;
begin
  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;

    MensagemSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionRT,
      'tblProgramacaoRT',
      'RT_Mensagem',
      Trim(Mensagem),
      255
    );

    StatusSafe := AjustarTextoParaCampo(
      FrmDataModule.ADOConnectionRT,
      'tblProgramacaoRT',
      'RT_Status',
      StatusRTNormalizado(RT_Status),
      40
    );

    SQL :=
      'UPDATE tblProgramacaoRT SET ' +
      '  RT_Cancelada = :RT_Cancelada, ' +
      '  RT_Mensagem = :RT_Mensagem ';

    if StatusSafe <> '' then
      SQL := SQL + ', RT_Status = :RT_Status ';

    SQL := SQL + 'WHERE idProgramacaoRT = :idProgramacaoRT';

    Q.SQL.Text := SQL;

    Q.Parameters.ParamByName('RT_Cancelada').Value := Cancelada;
    Q.Parameters.ParamByName('RT_Mensagem').Value := MensagemSafe;

    if StatusSafe <> '' then
      Q.Parameters.ParamByName('RT_Status').Value := StatusSafe;

    Q.Parameters.ParamByName('idProgramacaoRT').Value := idProgramacaoRT;

    Q.ExecSQL;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.BuscarRTLocalPorChave(
  const Dados: TDadosRT;
  out AIdRT: Integer;
  out ART, AMensagem, AStatus: string): Boolean;
var
  Q: TADOQuery;

  function BuscarPorCampo(const ACampo, AValor: string): Boolean;
  begin
    Result := False;

    if Trim(AValor) = '' then
      Exit;

    Q.Close;
    Q.SQL.Clear;
    Q.SQL.Text :=
      'SELECT TOP 1 ' +
      '  idProgramacaoRT, RT, RT_Mensagem, RT_Status ' +
      'FROM tblProgramacaoRT ' +
      'WHERE ( ([' + ACampo + '] = :Valor) ' +
      '        OR ([' + ACampo + '] IS NOT NULL AND Trim([' + ACampo + ']) = :Valor) ) ' +
      '  AND ([RT_Cancelada] IS NULL OR [RT_Cancelada] = False) ' +
      'ORDER BY IIf([RT] IS NULL OR Trim([RT]) = '''', 1, 0), idProgramacaoRT DESC';

    Q.Parameters.ParamByName('Valor').Value := Trim(AValor);
    Q.Open;

    Result := not Q.IsEmpty;
  end;

begin
  Result := False;
  AIdRT := 0;
  ART := '';
  AMensagem := '';
  AStatus := '';

  Q := TADOQuery.Create(nil);
  try
    Q.Connection := FrmDataModule.ADOConnectionRT;
    Q.ParamCheck := True;

    if BuscarPorCampo('ChaveCompleta', Dados.ChaveCompleta) or
       BuscarPorCampo('ChaveIda', Dados.ChaveIda) or
       ((Trim(Dados.ChaveVolta) <> '') and BuscarPorCampo('ChaveVolta', Dados.ChaveVolta)) then
    begin
      AIdRT := Q.FieldByName('idProgramacaoRT').AsInteger;
      ART := Trim(Q.FieldByName('RT').AsString);
      AMensagem := Trim(Q.FieldByName('RT_Mensagem').AsString);
      AStatus := Trim(Q.FieldByName('RT_Status').AsString);
      Result := True;
    end;
  finally
    Q.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.Existe_RT(
  const Dados: TDadosRT): TRT;
var
  IdRTLocal: Integer;
  RTLocal, MsgLocal, StatusLocal: string;
  DadosSAP: TRTSapImportDados;
begin
  Result := Default(TRT);

  { 1. procura primeiro na base local }
  if BuscarRTLocalPorChave(Dados, IdRTLocal, RTLocal, MsgLocal, StatusLocal) then
  begin
    Result.booRTExiste := True;
    Result.idProgramacaoRT := IdRTLocal;
    Result.RT_Numero := Trim(RTLocal);
    Exit;
  end;

  { 2. se não achou local, procura na importação SAP }
  if BuscarRTSapImportPorChave(Dados, DadosSAP) then
  begin
    Result.booRTExiste := True;
    Result.idProgramacaoRT := DadosSAP.idProgramacaoRT; // pode ser 0
    Result.RT_Numero := Trim(DadosSAP.QMNUM);
    Exit;
  end;
end;


procedure TFrmConsultaExecutantesProgramados.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmConsultaExecutantesProgramados:=nil;
end;

procedure TFrmConsultaExecutantesProgramados.DefinirLinhaSimulacao(
  const ARow: Integer; const ANome, AValor, AObservacao: string);
begin
  GridSimulacaoParametros.Cells[SIM_COL_PARAMETRO, ARow] := ANome;
  GridSimulacaoParametros.Cells[SIM_COL_VALOR, ARow] := AValor;
  GridSimulacaoParametros.Cells[SIM_COL_OBSERVACAO, ARow] := AObservacao;
end;

function TFrmConsultaExecutantesProgramados.ValorSimulacaoTexto(
  const ARow: Integer): string;
begin
  Result := Trim(GridSimulacaoParametros.Cells[SIM_COL_VALOR, ARow]);
end;

function TFrmConsultaExecutantesProgramados.TryParseNumeroSimulacao(
  const ATexto: string; out AValor: Double): Boolean;
var
  TextoNormalizado: string;
  FS: TFormatSettings;
  UltimoPonto: Integer;
  UltimaVirgula: Integer;
begin
  Result := False;
  TextoNormalizado := Trim(ATexto);
  if TextoNormalizado = '' then
    Exit;

  FS := TFormatSettings.Create;
  UltimoPonto := LastDelimiter('.', TextoNormalizado);
  UltimaVirgula := LastDelimiter(',', TextoNormalizado);

  if (UltimoPonto > 0) and (UltimaVirgula > 0) then
  begin
    if UltimoPonto > UltimaVirgula then
    begin
      TextoNormalizado := StringReplace(TextoNormalizado, ',', '', [rfReplaceAll]);
      TextoNormalizado := StringReplace(TextoNormalizado, '.', FS.DecimalSeparator, [rfReplaceAll]);
    end
    else
    begin
      TextoNormalizado := StringReplace(TextoNormalizado, '.', '', [rfReplaceAll]);
      TextoNormalizado := StringReplace(TextoNormalizado, ',', FS.DecimalSeparator, [rfReplaceAll]);
    end;
  end
  else
  begin
    if FS.DecimalSeparator = ',' then
      TextoNormalizado := StringReplace(TextoNormalizado, '.', FS.DecimalSeparator, [rfReplaceAll])
    else
      TextoNormalizado := StringReplace(TextoNormalizado, ',', FS.DecimalSeparator, [rfReplaceAll]);
  end;

  Result := TryStrToFloat(TextoNormalizado, AValor, FS);
end;

function TFrmConsultaExecutantesProgramados.LerFloatSimulacao(
  const ARow: Integer; const ANome: string): Double;
begin
  if not TryParseNumeroSimulacao(ValorSimulacaoTexto(ARow), Result) then
    raise Exception.CreateFmt('Valor invalido para "%s": %s',
      [ANome, ValorSimulacaoTexto(ARow)]);
end;

function TFrmConsultaExecutantesProgramados.LerInteiroSimulacao(
  const ARow: Integer; const ANome: string): Integer;
begin
  Result := Round(LerFloatSimulacao(ARow, ANome));
end;

function TFrmConsultaExecutantesProgramados.FormatarNumeroSimulacao(
  const AValor: Double; const ACasas: Integer): string;
var
  FS: TFormatSettings;
  Mascara: string;
begin
  FS := TFormatSettings.Create;
  Mascara := '0';
  if ACasas > 0 then
    Mascara := Mascara + FS.DecimalSeparator + StringOfChar('0', ACasas);
  Result := FormatFloat(Mascara, AValor, FS);
end;

function TFrmConsultaExecutantesProgramados.ClassificarModalSimulacao(
  const ANomeEmbarcacao, ATipoEmbarcacao: string; const ACapacidade: Integer;
  const AUsaBridgeMesmoGrupo: Boolean): string;
var
  NomeUpper, TipoUpper: string;
begin
  NomeUpper := UpperCase(Trim(ANomeEmbarcacao));
  TipoUpper := UpperCase(Trim(ATipoEmbarcacao));

  if (Pos('AQUA', NomeUpper) > 0) or
     (Pos('HELIX', NomeUpper) > 0) or
     (ACapacidade >= 100) then
    Exit('AQUA');

  if (Pos('BAE', NomeUpper) > 0) or
     (Pos('AUTO', TipoUpper) > 0) or
     (Pos('ELEV', TipoUpper) > 0) or
     (Pos('JACK', TipoUpper) > 0) then
    Exit('BAE');

  if AUsaBridgeMesmoGrupo or
     (Pos('SOV', NomeUpper) > 0) or
     (Pos('BRIDGE', NomeUpper) > 0) or
     (Pos('SOV', TipoUpper) > 0) then
    Exit('SOV');

  Result := 'SURFER';
end;

function TFrmConsultaExecutantesProgramados.EstimarDuracaoRotaModalHoras(
  const AModal, ARotaSequencia: string; const ATempoTransbordo: Integer): Double;
var
  Pontos: TStringList;
  I, QtPontos, MinimoMin, DuracaoMin: Integer;
begin
  Pontos := TStringList.Create;
  try
    Pontos.Delimiter := ';';
    Pontos.StrictDelimiter := True;
    Pontos.DelimitedText := Trim(ARotaSequencia);

    QtPontos := 0;
    for I := 0 to Pontos.Count - 1 do
      if Trim(Pontos[I]) <> '' then
        Inc(QtPontos);

    DuracaoMin := ((QtPontos + 1) * 15) +
      (Max(0, QtPontos - 1) * Max(0, ATempoTransbordo));

    if SameText(AModal, 'SURFER') then
      MinimoMin := 60
    else if SameText(AModal, 'AQUA') then
      MinimoMin := 40
    else if SameText(AModal, 'SOV') then
      MinimoMin := 45
    else if SameText(AModal, 'BAE') then
      MinimoMin := 60
    else
      MinimoMin := 45;

    Result := Max(MinimoMin, DuracaoMin) / 60;
  finally
    Pontos.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AplicarMetricasModalNaGrade(
  const AModal: string; const AQtdEmbarcacoes: Integer; const ADisponibilidade,
  AHorasUteis, AProdutividade: Double);
var
  RowQtd, RowDisp, RowHoras, RowProd: Integer;
begin
  if SameText(AModal, 'SURFER') then
  begin
    RowQtd := SIM_ROW_QTD_SURFER;
    RowDisp := SIM_ROW_DISP_SURFER;
    RowHoras := SIM_ROW_HORAS_SURFER;
    RowProd := SIM_ROW_PROD_SURFER;
  end
  else if SameText(AModal, 'SOV') then
  begin
    RowQtd := SIM_ROW_QTD_SOV;
    RowDisp := SIM_ROW_DISP_SOV;
    RowHoras := SIM_ROW_HORAS_SOV;
    RowProd := SIM_ROW_PROD_SOV;
  end
  else if SameText(AModal, 'BAE') then
  begin
    RowQtd := SIM_ROW_QTD_BAE;
    RowDisp := SIM_ROW_DISP_BAE;
    RowHoras := SIM_ROW_HORAS_BAE;
    RowProd := SIM_ROW_PROD_BAE;
  end
  else if SameText(AModal, 'AQUA') then
  begin
    RowQtd := SIM_ROW_QTD_AQUA;
    RowDisp := SIM_ROW_DISP_AQUA;
    RowHoras := SIM_ROW_HORAS_AQUA;
    RowProd := SIM_ROW_PROD_AQUA;
  end
  else
    Exit;

  GridSimulacaoParametros.Cells[SIM_COL_VALOR, RowQtd] := IntToStr(AQtdEmbarcacoes);
  GridSimulacaoParametros.Cells[SIM_COL_VALOR, RowDisp] :=
    FormatarNumeroSimulacao(ADisponibilidade, 1);
  GridSimulacaoParametros.Cells[SIM_COL_VALOR, RowHoras] :=
    FormatarNumeroSimulacao(AHorasUteis, 1);
  GridSimulacaoParametros.Cells[SIM_COL_VALOR, RowProd] :=
    FormatarNumeroSimulacao(AProdutividade, 2);
end;

procedure TFrmConsultaExecutantesProgramados.CarregarParametrosSimulacaoComBaseReal;
var
  QryRotas, QryEmbarcacoes: TADOQuery;
  InfoEmbarcacoes: TDictionary<string, TInfoEmbarcacaoSimulacao>;
  Info: TInfoEmbarcacaoSimulacao;
  MetSurfer, MetSOV, MetBAE, MetAQUA: TMetricasModalReal;
  MetricasModal: TMetricasModalReal;
  Modal, NomeEmbarcacao, ChaveEmbarcacao, DataKey, RotaKey: string;
  BacklogReal, DiasPeriodo, Capacidade, TempoTransbordo: Integer;
  LogResumo: TStringList;
  function ObterMetricas(const AModal: string): TMetricasModalReal;
  begin
    if SameText(AModal, 'SOV') then
      Exit(MetSOV);
    if SameText(AModal, 'BAE') then
      Exit(MetBAE);
    if SameText(AModal, 'AQUA') then
      Exit(MetAQUA);
    Result := MetSurfer;
  end;
  function LerBooleanoNullAsFalse(const AQry: TADOQuery;
    const ACampo: string): Boolean;
  begin
    Result := (AQry.FindField(ACampo) <> nil) and
      (not VarIsNull(AQry.FieldByName(ACampo).Value)) and
      AQry.FieldByName(ACampo).AsBoolean;
  end;
  function LerInteiroNullAsZero(const AQry: TADOQuery;
    const ACampo: string): Integer;
  begin
    if (AQry.FindField(ACampo) <> nil) and
       (not VarIsNull(AQry.FieldByName(ACampo).Value)) then
      Result := AQry.FieldByName(ACampo).AsInteger
    else
      Result := 0;
  end;
  function CalcularDisponibilidade(const AMetricas: TMetricasModalReal): Double;
  begin
    if (AMetricas.Embarcacoes.Count > 0) and (DiasPeriodo > 0) then
      Result := (AMetricas.EmbarcacaoDiasAtivos.Count /
        (AMetricas.Embarcacoes.Count * DiasPeriodo)) * 100
    else
      Result := 0;
  end;
  function CalcularHorasUteis(const AMetricas: TMetricasModalReal): Double;
  begin
    if AMetricas.EmbarcacaoDiasAtivos.Count > 0 then
      Result := AMetricas.TotalHorasRota / AMetricas.EmbarcacaoDiasAtivos.Count
    else
      Result := 0;
  end;
  function CalcularProdutividade(const AMetricas: TMetricasModalReal): Double;
  begin
    if AMetricas.TotalHorasRota > 0 then
      Result := AMetricas.TotalMovimentacoes / AMetricas.TotalHorasRota
    else
      Result := 0;
  end;
  procedure AdicionarResumoModal(const ATitulo: string;
    const AMetricas: TMetricasModalReal);
  begin
    LogResumo.Add(Format('%s: embarcacoes=%d | mov=%d | ocupacao media=%.1f%% | disponibilidade=%.1f%% | horas uteis/dia ativo=%.1f | produtividade=%.2f',
      [ATitulo,
       AMetricas.Embarcacoes.Count,
       AMetricas.TotalMovimentacoes,
       AMetricas.OcupacaoMediaPercentual,
       CalcularDisponibilidade(AMetricas),
       CalcularHorasUteis(AMetricas),
       CalcularProdutividade(AMetricas)]));
  end;
begin
  AtualizarDiasDisponiveisSimulacao;
  actProcurarProgramacaoExecutante.Execute;
  DiasPeriodo := LerInteiroSimulacao(SIM_ROW_DIAS_DISPONIVEIS, 'Dias disponiveis');

  if FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active then
    BacklogReal := Max(0, FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecordCount)
  else
    BacklogReal := 0;

  GridSimulacaoParametros.Cells[SIM_COL_VALOR, SIM_ROW_BACKLOG] :=
    IntToStr(BacklogReal);

  QryRotas := TADOQuery.Create(nil);
  QryEmbarcacoes := TADOQuery.Create(nil);
  InfoEmbarcacoes := TDictionary<string, TInfoEmbarcacaoSimulacao>.Create;
  MetSurfer := TMetricasModalReal.Create;
  MetSOV := TMetricasModalReal.Create;
  MetBAE := TMetricasModalReal.Create;
  MetAQUA := TMetricasModalReal.Create;
  LogResumo := TStringList.Create;
  try
    QryEmbarcacoes.Connection := FrmDataModule.ADOConnectionConsulta;
    QryEmbarcacoes.SQL.Text :=
      'SELECT NomeEmbarcacao, TipoEmbarcacao, UsaBridgeMesmoGrupo ' +
      'FROM tblEmbarcacao';
    QryEmbarcacoes.Open;

    while not QryEmbarcacoes.Eof do
    begin
      ChaveEmbarcacao := UpperCase(Trim(QryEmbarcacoes.FieldByName('NomeEmbarcacao').AsString));
      if ChaveEmbarcacao <> '' then
      begin
        Info := Default(TInfoEmbarcacaoSimulacao);
        Info.TipoEmbarcacao := Trim(QryEmbarcacoes.FieldByName('TipoEmbarcacao').AsString);
        Info.UsaBridgeMesmoGrupo := LerBooleanoNullAsFalse(QryEmbarcacoes,
          'UsaBridgeMesmoGrupo');
        if InfoEmbarcacoes.ContainsKey(ChaveEmbarcacao) then
          InfoEmbarcacoes[ChaveEmbarcacao] := Info
        else
          InfoEmbarcacoes.Add(ChaveEmbarcacao, Info);
      end;
      QryEmbarcacoes.Next;
    end;

    QryRotas.Connection := FrmDataModule.ADOConnectionColibri;
    QryRotas.SQL.Text :=
      'SELECT ' +
      '  pd.DataProgramacao, ' +
      '  pe.idProgramacaoExecutante, ' +
      '  r.idRoteamento, ' +
      '  r.NomeEmbarcacao, ' +
      '  r.CapacidadePAX, ' +
      '  r.RotaSequencia, ' +
      '  r.TempoTransbordo ' +
      'FROM ' +
      '  ((tblProgramacaoDiaria pd ' +
      '  INNER JOIN tblProgramacaoExecutante pe ' +
      '    ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria) ' +
      '  INNER JOIN (tblRoteamento r ' +
      '    INNER JOIN tblAux_Rota_Distribuicao ard ' +
      '      ON r.idRoteamento = ard.CodigoRota) ' +
      '    ON pe.idProgramacaoExecutante = ard.CodigoProgramacaoExecutante) ' +
      'WHERE ' +
      '  pd.DataProgramacao >= :DtIni ' +
      '  AND pd.DataProgramacao < :DtFimMais1 ' +
      'ORDER BY pd.DataProgramacao, r.NomeEmbarcacao, r.idRoteamento, pe.idProgramacaoExecutante';
    QryRotas.Parameters.ParamByName('DtIni').Value := Trunc(dataInicio.Date);
    QryRotas.Parameters.ParamByName('DtFimMais1').Value := Trunc(dataFim.Date) + 1;
    QryRotas.Open;

    if QryRotas.IsEmpty then
    begin
      MemoSimulacaoRelatorio.Lines.Text :=
        'Nao foram encontradas alocacoes reais em rota no periodo filtrado.' + sLineBreak +
        'O backlog foi atualizado com base na consulta atual, mas os demais parametros foram preservados.';
      StatusBarSimulacao.SimpleText :=
        'Sem rotas alocadas no periodo para calibrar a base real.';
      Exit;
    end;

    while not QryRotas.Eof do
    begin
      NomeEmbarcacao := Trim(QryRotas.FieldByName('NomeEmbarcacao').AsString);
      ChaveEmbarcacao := UpperCase(NomeEmbarcacao);
      Capacidade := LerInteiroNullAsZero(QryRotas, 'CapacidadePAX');
      TempoTransbordo := LerInteiroNullAsZero(QryRotas, 'TempoTransbordo');

      if not InfoEmbarcacoes.TryGetValue(ChaveEmbarcacao, Info) then
        Info := Default(TInfoEmbarcacaoSimulacao);

      Modal := ClassificarModalSimulacao(NomeEmbarcacao, Info.TipoEmbarcacao,
        Capacidade, Info.UsaBridgeMesmoGrupo);
      MetricasModal := ObterMetricas(Modal);
      if MetricasModal <> nil then
      begin
        Inc(MetricasModal.TotalMovimentacoes);

        if NomeEmbarcacao <> '' then
          MetricasModal.Embarcacoes.Add(ChaveEmbarcacao);

        DataKey := FormatDateTime('yyyymmdd',
          QryRotas.FieldByName('DataProgramacao').AsDateTime);
        MetricasModal.DiasAtivos.Add(DataKey);
        MetricasModal.EmbarcacaoDiasAtivos.Add(DataKey + '|' + ChaveEmbarcacao);

        RotaKey := DataKey + '|' + IntToStr(QryRotas.FieldByName('idRoteamento').AsInteger);
        if MetricasModal.RotasProcessadas.IndexOf(RotaKey) < 0 then
        begin
          MetricasModal.RotasProcessadas.Add(RotaKey);
          if Capacidade > 0 then
            MetricasModal.TotalCapacidadeOferta := MetricasModal.TotalCapacidadeOferta + Capacidade;
          MetricasModal.TotalHorasRota := MetricasModal.TotalHorasRota + EstimarDuracaoRotaModalHoras(Modal,
            QryRotas.FieldByName('RotaSequencia').AsString, TempoTransbordo);
        end;
      end;

      QryRotas.Next;
    end;

    AplicarMetricasModalNaGrade('SURFER', MetSurfer.Embarcacoes.Count,
      CalcularDisponibilidade(MetSurfer), CalcularHorasUteis(MetSurfer),
      CalcularProdutividade(MetSurfer));
    AplicarMetricasModalNaGrade('SOV', MetSOV.Embarcacoes.Count,
      CalcularDisponibilidade(MetSOV), CalcularHorasUteis(MetSOV),
      CalcularProdutividade(MetSOV));
    AplicarMetricasModalNaGrade('BAE', MetBAE.Embarcacoes.Count,
      CalcularDisponibilidade(MetBAE), CalcularHorasUteis(MetBAE),
      CalcularProdutividade(MetBAE));
    AplicarMetricasModalNaGrade('AQUA', MetAQUA.Embarcacoes.Count,
      CalcularDisponibilidade(MetAQUA), CalcularHorasUteis(MetAQUA),
      CalcularProdutividade(MetAQUA));

    LogResumo.Add('Base real carregada do banco para o periodo filtrado.');
    LogResumo.Add(Format('Backlog filtrado atual: %d executantes/movimentacoes.',
      [BacklogReal]));
    LogResumo.Add('As quantidades, disponibilidades, horas uteis e produtividades foram sugeridas a partir das alocacoes reais em rota.');
    LogResumo.Add('Classificacao de modal inferida por cadastro da embarcacao, bridge, nome e capacidade quando necessario.');
    LogResumo.Add('');
    AdicionarResumoModal('Surfer', MetSurfer);
    AdicionarResumoModal('SOV', MetSOV);
    AdicionarResumoModal('BAE', MetBAE);
    AdicionarResumoModal('AQUA', MetAQUA);
    LogResumo.Add('');
    LogResumo.Add('Custos e mobilizacao permaneceram manuais para revisao contratual.');

    MemoSimulacaoRelatorio.Lines.Text := LogResumo.Text;
    StatusBarSimulacao.SimpleText :=
      'Base real carregada do banco. Revise custos e mobilizacao antes de rodar a simulacao.';
  finally
    LogResumo.Free;
    MetAQUA.Free;
    MetBAE.Free;
    MetSOV.Free;
    MetSurfer.Free;
    InfoEmbarcacoes.Free;
    QryEmbarcacoes.Free;
    QryRotas.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.LerParametrosSimulacaoAtual: TSimulacaoParametros;
begin
  Result.QtdSurfer := LerInteiroSimulacao(SIM_ROW_QTD_SURFER, 'Qtd. Surfer');
  Result.QtdSOV := LerInteiroSimulacao(SIM_ROW_QTD_SOV, 'Qtd. SOV');
  Result.QtdBAE := LerInteiroSimulacao(SIM_ROW_QTD_BAE, 'Qtd. BAE');
  Result.QtdAQUA := LerInteiroSimulacao(SIM_ROW_QTD_AQUA, 'Qtd. AQUA');
  Result.Backlog := LerInteiroSimulacao(SIM_ROW_BACKLOG, 'Backlog');
  Result.DiasDisponiveis := LerInteiroSimulacao(SIM_ROW_DIAS_DISPONIVEIS,
    'Dias disponiveis');
  Result.CustoSurfer := LerFloatSimulacao(SIM_ROW_CUSTO_SURFER,
    'Custo diario Surfer');
  Result.CustoSOV := LerFloatSimulacao(SIM_ROW_CUSTO_SOV, 'Custo diario SOV');
  Result.CustoBAE := LerFloatSimulacao(SIM_ROW_CUSTO_BAE, 'Custo diario BAE');
  Result.CustoAQUA := LerFloatSimulacao(SIM_ROW_CUSTO_AQUA,
    'Custo diario AQUA');
  Result.ProdutividadeSurfer := LerFloatSimulacao(SIM_ROW_PROD_SURFER,
    'Produtividade Surfer');
  Result.ProdutividadeSOV := LerFloatSimulacao(SIM_ROW_PROD_SOV,
    'Produtividade SOV');
  Result.ProdutividadeBAE := LerFloatSimulacao(SIM_ROW_PROD_BAE,
    'Produtividade BAE');
  Result.ProdutividadeAQUA := LerFloatSimulacao(SIM_ROW_PROD_AQUA,
    'Produtividade AQUA');
  Result.HorasUteisSurfer := LerFloatSimulacao(SIM_ROW_HORAS_SURFER,
    'Horas uteis Surfer');
  Result.HorasUteisSOV := LerFloatSimulacao(SIM_ROW_HORAS_SOV,
    'Horas uteis SOV');
  Result.HorasUteisBAE := LerFloatSimulacao(SIM_ROW_HORAS_BAE,
    'Horas uteis BAE');
  Result.HorasUteisAQUA := LerFloatSimulacao(SIM_ROW_HORAS_AQUA,
    'Horas uteis AQUA');
  Result.DisponibilidadeSurfer := LerFloatSimulacao(SIM_ROW_DISP_SURFER,
    'Disponibilidade Surfer');
  Result.DisponibilidadeSOV := LerFloatSimulacao(SIM_ROW_DISP_SOV,
    'Disponibilidade SOV');
  Result.DisponibilidadeBAE := LerFloatSimulacao(SIM_ROW_DISP_BAE,
    'Disponibilidade BAE');
  Result.DisponibilidadeAQUA := LerFloatSimulacao(SIM_ROW_DISP_AQUA,
    'Disponibilidade AQUA');
  Result.MobilizacaoSurfer := LerFloatSimulacao(SIM_ROW_MOB_SURFER,
    'Mobilizacao Surfer');
  Result.MobilizacaoSOV := LerFloatSimulacao(SIM_ROW_MOB_SOV,
    'Mobilizacao SOV');
  Result.MobilizacaoBAE := LerFloatSimulacao(SIM_ROW_MOB_BAE,
    'Mobilizacao BAE');
  Result.MobilizacaoAQUA := LerFloatSimulacao(SIM_ROW_MOB_AQUA,
    'Mobilizacao AQUA');
end;

function TFrmConsultaExecutantesProgramados.DescreverFrotaSimulacao(
  const AParam: TSimulacaoParametros): string;
begin
  Result := Format('Surfer=%d | SOV=%d | BAE=%d | AQUA=%d',
    [AParam.QtdSurfer, AParam.QtdSOV, AParam.QtdBAE, AParam.QtdAQUA]);
end;

function TFrmConsultaExecutantesProgramados.SimularCenarioLogistico(
  const ANome: string; const AParam: TSimulacaoParametros): TSimulacaoCenario;
begin
  Result.Nome := ANome;
  Result.Parametros := AParam;
  Result.Resultado := TSimulacaoLogistica.Simular(AParam);
end;

procedure TFrmConsultaExecutantesProgramados.MontarCenariosComparativos(
  const ABase: TSimulacaoParametros; var ACenarios: array of TSimulacaoCenario);
var
  Param: TSimulacaoParametros;
begin
  if Length(ACenarios) < 5 then
    raise Exception.Create('Quantidade insuficiente de cenarios para o comparativo.');

  ACenarios[0] := SimularCenarioLogistico('Atual', ABase);

  Param := ABase;
  Inc(Param.QtdSOV);
  ACenarios[1] := SimularCenarioLogistico('Atual + 1 SOV', Param);

  Param := ABase;
  Inc(Param.QtdBAE);
  ACenarios[2] := SimularCenarioLogistico('Atual + 1 BAE', Param);

  Param := ABase;
  Inc(Param.QtdAQUA);
  ACenarios[3] := SimularCenarioLogistico('Atual + 1 AQUA', Param);

  Param := ABase;
  Inc(Param.QtdSOV);
  Inc(Param.QtdBAE);
  ACenarios[4] := SimularCenarioLogistico('Atual + 1 SOV + 1 BAE', Param);
end;

function TFrmConsultaExecutantesProgramados.GerarRelatorioComparativoTexto(
  const ACenarios: array of TSimulacaoCenario): string;
var
  I: Integer;
  Base: TSimulacaoCenario;
  Comparativo: TSimulacaoComparativo;
  Linhas: TStringList;
  MelhorPrazoIdx: Integer;
  MelhorResilienciaIdx: Integer;
  MelhorROIIdx: Integer;
  MelhorROI: Double;
  MelhorComparativo: TSimulacaoComparativo;
  function TextoPrazoComparativo(
    const AComparativo: TSimulacaoComparativo): string;
  begin
    if AComparativo.TornaViavel then
      Result := 'torna viavel concluir o backlog'
    else if AComparativo.PrazoGanhoDias > 0 then
      Result := Format('%d dias mais rapido', [AComparativo.PrazoGanhoDias])
    else if AComparativo.PrazoGanhoDias < 0 then
      Result := Format('%d dias mais lento', [Abs(AComparativo.PrazoGanhoDias)])
    else
      Result := 'sem ganho de prazo';
  end;
begin
  if Length(ACenarios) = 0 then
    Exit('');

  Base := ACenarios[Low(ACenarios)];
  MelhorPrazoIdx := Low(ACenarios);
  MelhorResilienciaIdx := Low(ACenarios);
  MelhorROIIdx := -1;
  MelhorROI := -MaxDouble;

  for I := Low(ACenarios) to High(ACenarios) do
  begin
    if (ACenarios[I].Resultado.PrazoEstimado >= 0) and
       ((ACenarios[MelhorPrazoIdx].Resultado.PrazoEstimado < 0) or
        (ACenarios[I].Resultado.PrazoEstimado <
         ACenarios[MelhorPrazoIdx].Resultado.PrazoEstimado)) then
      MelhorPrazoIdx := I;

    if ACenarios[I].Resultado.IndiceResiliencia >
       ACenarios[MelhorResilienciaIdx].Resultado.IndiceResiliencia then
      MelhorResilienciaIdx := I;

    if I > Low(ACenarios) then
    begin
      Comparativo := TSimulacaoLogistica.CompararCenarios(Base, ACenarios[I]);
      if (Comparativo.CustoAdicional > 0) and
         (Comparativo.ROIOperacional > MelhorROI) then
      begin
        MelhorROI := Comparativo.ROIOperacional;
        MelhorROIIdx := I;
        MelhorComparativo := Comparativo;
      end;
    end;
  end;

  Linhas := TStringList.Create;
  try
    Linhas.Add('--- Comparativo de Cenarios Logistica ---');
    Linhas.Add('Premissas: disponibilidade e mobilizacao entram como proxies de clima, contrato e indisponibilidade.');
    Linhas.Add('ROI e payback sao proxies calculados pelo custo unitario do cenario Atual.');
    Linhas.Add('');
    Linhas.Add('Resumo executivo');
    Linhas.Add('Melhor prazo projetado: ' + ACenarios[MelhorPrazoIdx].Nome +
      ' | ' + TSimulacaoLogistica.PrazoComoTexto(
        ACenarios[MelhorPrazoIdx].Resultado));
    Linhas.Add(Format('Maior resiliencia: %s | %.1f%%',
      [ACenarios[MelhorResilienciaIdx].Nome,
       ACenarios[MelhorResilienciaIdx].Resultado.IndiceResiliencia]));
    if MelhorROIIdx >= 0 then
      Linhas.Add('Melhor ROI operacional proxy: ' + ACenarios[MelhorROIIdx].Nome +
        ' | ' + TSimulacaoLogistica.ROIComoTexto(MelhorComparativo) +
        ' | Payback: ' + TSimulacaoLogistica.PaybackComoTexto(MelhorComparativo))
    else
      Linhas.Add('Melhor ROI operacional proxy: nao houve cenario adicional com custo extra comparavel.');
    Linhas.Add('');

    for I := Low(ACenarios) to High(ACenarios) do
    begin
      Linhas.Add(Format('%d. Cenario: %s', [I + 1, ACenarios[I].Nome]));
      Linhas.Add('   Frota: ' + DescreverFrotaSimulacao(ACenarios[I].Parametros));
      Linhas.Add('   ' + TSimulacaoLogistica.PrazoComoTexto(ACenarios[I].Resultado));
      Linhas.Add(Format('   Movimentacoes por dia: %.1f',
        [ACenarios[I].Resultado.MovimentacoesPorDia]));
      Linhas.Add(Format('   Horas planejadas/efetivas por dia: %.1f / %.1f',
        [ACenarios[I].Resultado.HorasPlanejadasPorDia,
         ACenarios[I].Resultado.HorasEfetivasPorDia]));
      Linhas.Add(Format('   Resiliencia estimada: %.1f%%',
        [ACenarios[I].Resultado.IndiceResiliencia]));
      Linhas.Add(Format('   Custo operacional: R$ %.2f | Mobilizacao: R$ %.2f',
        [ACenarios[I].Resultado.CustoOperacionalTotal,
         ACenarios[I].Resultado.CustoMobilizacaoTotal]));
      Linhas.Add(Format('   Custo total: R$ %.2f',
        [ACenarios[I].Resultado.CustoTotal]));
      Linhas.Add(Format('   Horas uteis geradas: %.1f',
        [ACenarios[I].Resultado.HorasUteisGeradas]));
      Linhas.Add(Format('   Custo por hora util: R$ %.2f | Custo por movimentacao: R$ %.2f',
        [ACenarios[I].Resultado.CustoPorHoraUtil,
         ACenarios[I].Resultado.CustoPorMovimentacao]));
      Linhas.Add(Format('   Backlog atendido: %d (%.1f%%)',
        [ACenarios[I].Resultado.BacklogAtendido,
         ACenarios[I].Resultado.PercentualBacklogAtendido]));
      Linhas.Add(Format('   Backlog restante: %d',
        [ACenarios[I].Resultado.BacklogRestante]));

      if I > Low(ACenarios) then
      begin
        Comparativo := TSimulacaoLogistica.CompararCenarios(Base, ACenarios[I]);
        Linhas.Add('   Comparado ao Atual: ' + TextoPrazoComparativo(Comparativo));
        Linhas.Add(Format('   Delta de movimentacoes/dia: %.1f (%.1f%%)',
          [Comparativo.MovimentacoesAdicionaisPorDia,
           Comparativo.GanhoCapacidadePercentual]));
        Linhas.Add(Format('   Horas uteis adicionais no periodo: %.1f',
          [Comparativo.HorasUteisAdicionais]));
        Linhas.Add(Format('   Backlog adicional atendido: %d',
          [Comparativo.BacklogAdicionalAtendido]));
        Linhas.Add(Format('   Custo adicional: R$ %.2f',
          [Comparativo.CustoAdicional]));
        if Comparativo.CustoPorDiaGanho > 0 then
          Linhas.Add(Format('   Custo por dia ganho: R$ %.2f',
            [Comparativo.CustoPorDiaGanho]));
        Linhas.Add(Format('   Valor operacional equivalente: R$ %.2f',
          [Comparativo.ValorOperacionalEquivalente]));
        Linhas.Add('   ROI operacional proxy: ' +
          TSimulacaoLogistica.ROIComoTexto(Comparativo));
        Linhas.Add('   Payback proxy: ' +
          TSimulacaoLogistica.PaybackComoTexto(Comparativo));
      end;

      Linhas.Add('');
    end;

    Result := Linhas.Text;
  finally
    Linhas.Free;
  end;
end;

function TFrmConsultaExecutantesProgramados.GerarRelatorioComparativoCSV(
  const ACenarios: array of TSimulacaoCenario): string;
var
  I: Integer;
  Comparativo: TSimulacaoComparativo;
  Linhas: TStringList;
  PrazoTexto: string;
  ROITexto: string;
  PaybackTexto: string;
begin
  Linhas := TStringList.Create;
  try
    Linhas.Add('Cenario;Surfer;SOV;BAE;AQUA;MovimentacoesPorDia;PrazoEstimadoDias;HorasPlanejadasDia;HorasEfetivasDia;IndiceResiliencia;CustoOperacionalTotal;CustoMobilizacaoTotal;CustoTotal;HorasUteisGeradas;CustoPorHoraUtil;CustoPorMovimentacao;BacklogAtendido;BacklogRestante;PercentualBacklogAtendido;DeltaMovimentacoesDia;PrazoGanhoDias;TornaViavel;CustoAdicional;BacklogAdicionalAtendido;ValorOperacionalEquivalente;ROIProxy;PaybackDias');

    for I := Low(ACenarios) to High(ACenarios) do
    begin
      if I = Low(ACenarios) then
        Comparativo := Default(TSimulacaoComparativo)
      else
        Comparativo := TSimulacaoLogistica.CompararCenarios(
          ACenarios[Low(ACenarios)], ACenarios[I]);

      if ACenarios[I].Resultado.PrazoEstimado < 0 then
        PrazoTexto := 'INVIAVEL'
      else
        PrazoTexto := IntToStr(ACenarios[I].Resultado.PrazoEstimado);

      ROITexto := TSimulacaoLogistica.ROIComoTexto(Comparativo);
      if Comparativo.PaybackDias < 0 then
        PaybackTexto := 'NA'
      else
        PaybackTexto := Format('%.1f', [Comparativo.PaybackDias]);

      Linhas.Add(Format('%s;%d;%d;%d;%d;%.1f;%s;%.1f;%.1f;%.1f;%.2f;%.2f;%.2f;%.1f;%.2f;%.2f;%d;%d;%.1f;%.1f;%d;%s;%.2f;%d;%.2f;%s;%s',
        [ACenarios[I].Nome,
         ACenarios[I].Parametros.QtdSurfer,
         ACenarios[I].Parametros.QtdSOV,
         ACenarios[I].Parametros.QtdBAE,
         ACenarios[I].Parametros.QtdAQUA,
         ACenarios[I].Resultado.MovimentacoesPorDia,
         PrazoTexto,
         ACenarios[I].Resultado.HorasPlanejadasPorDia,
         ACenarios[I].Resultado.HorasEfetivasPorDia,
         ACenarios[I].Resultado.IndiceResiliencia,
         ACenarios[I].Resultado.CustoOperacionalTotal,
         ACenarios[I].Resultado.CustoMobilizacaoTotal,
         ACenarios[I].Resultado.CustoTotal,
         ACenarios[I].Resultado.HorasUteisGeradas,
         ACenarios[I].Resultado.CustoPorHoraUtil,
         ACenarios[I].Resultado.CustoPorMovimentacao,
         ACenarios[I].Resultado.BacklogAtendido,
         ACenarios[I].Resultado.BacklogRestante,
         ACenarios[I].Resultado.PercentualBacklogAtendido,
         Comparativo.MovimentacoesAdicionaisPorDia,
         Comparativo.PrazoGanhoDias,
         IfThen(Comparativo.TornaViavel, 'SIM', 'NAO'),
         Comparativo.CustoAdicional,
         Comparativo.BacklogAdicionalAtendido,
         Comparativo.ValorOperacionalEquivalente,
         ROITexto,
         PaybackTexto]));
    end;

    Result := Linhas.Text;
  finally
    Linhas.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.AtualizarDiasDisponiveisSimulacao;
var
  DiasDisponiveis: Integer;
begin
  if not Assigned(GridSimulacaoParametros) then
    Exit;

  if GridSimulacaoParametros.RowCount <= SIM_ROW_DIAS_DISPONIVEIS then
    Exit;

  DiasDisponiveis := DaysBetween(Trunc(dataFim.Date), Trunc(dataInicio.Date)) + 1;
  if DiasDisponiveis <= 0 then
    DiasDisponiveis := 1;

  GridSimulacaoParametros.Cells[SIM_COL_VALOR, SIM_ROW_DIAS_DISPONIVEIS] :=
    IntToStr(DiasDisponiveis);
end;

procedure TFrmConsultaExecutantesProgramados.PreencherCenarioBaseSimulacao;
begin
  DefinirLinhaSimulacao(SIM_ROW_QTD_SURFER, 'Qtd. Surfer', '5',
    'Frota base atual');
  DefinirLinhaSimulacao(SIM_ROW_QTD_SOV, 'Qtd. SOV', '0',
    'Norwind Gale Bridge');
  DefinirLinhaSimulacao(SIM_ROW_QTD_BAE, 'Qtd. BAE', '0',
    'Auto-elevatoria');
  DefinirLinhaSimulacao(SIM_ROW_QTD_AQUA, 'Qtd. AQUA', '0',
    'Gangway telescopica');
  DefinirLinhaSimulacao(SIM_ROW_BACKLOG, 'Backlog (movimentacoes)', '280',
    'Ajuste conforme a carteira analisada');
  DefinirLinhaSimulacao(SIM_ROW_DIAS_DISPONIVEIS, 'Dias disponiveis', '1',
    'Sincronizado com o periodo do filtro');
  DefinirLinhaSimulacao(SIM_ROW_CUSTO_SURFER, 'Custo diario Surfer', '10000',
    'R$ por dia');
  DefinirLinhaSimulacao(SIM_ROW_CUSTO_SOV, 'Custo diario SOV', '516000',
    'R$ por dia');
  DefinirLinhaSimulacao(SIM_ROW_CUSTO_BAE, 'Custo diario BAE', '0',
    'Preencher quando aplicavel');
  DefinirLinhaSimulacao(SIM_ROW_CUSTO_AQUA, 'Custo diario AQUA', '0',
    'Preencher quando aplicavel');
  DefinirLinhaSimulacao(SIM_ROW_PROD_SURFER, 'Produtividade Surfer', '1',
    'Mov/hora util');
  DefinirLinhaSimulacao(SIM_ROW_PROD_SOV, 'Produtividade SOV', '1',
    'Mov/hora util');
  DefinirLinhaSimulacao(SIM_ROW_PROD_BAE, 'Produtividade BAE', '1',
    'Mov/hora util');
  DefinirLinhaSimulacao(SIM_ROW_PROD_AQUA, 'Produtividade AQUA', '1',
    'Mov/hora util');
  DefinirLinhaSimulacao(SIM_ROW_HORAS_SURFER, 'Horas uteis Surfer', '6',
    'Base operacional atual');
  DefinirLinhaSimulacao(SIM_ROW_HORAS_SOV, 'Horas uteis SOV', '24',
    'Campanha continua');
  DefinirLinhaSimulacao(SIM_ROW_HORAS_BAE, 'Horas uteis BAE', '24',
    'Campanha continua');
  DefinirLinhaSimulacao(SIM_ROW_HORAS_AQUA, 'Horas uteis AQUA', '24',
    'Campanha continua');
  DefinirLinhaSimulacao(SIM_ROW_DISP_SURFER, 'Disponibilidade Surfer (%)', '60',
    'Proxy clima/janela operacional');
  DefinirLinhaSimulacao(SIM_ROW_DISP_SOV, 'Disponibilidade SOV (%)', '85',
    'Melhor tolerancia operacional');
  DefinirLinhaSimulacao(SIM_ROW_DISP_BAE, 'Disponibilidade BAE (%)', '90',
    'Campanha mais resiliente');
  DefinirLinhaSimulacao(SIM_ROW_DISP_AQUA, 'Disponibilidade AQUA (%)', '85',
    'Gangway telescopica');
  DefinirLinhaSimulacao(SIM_ROW_MOB_SURFER, 'Mobilizacao Surfer', '0',
    'Custo fixo por unidade no periodo');
  DefinirLinhaSimulacao(SIM_ROW_MOB_SOV, 'Mobilizacao SOV', '0',
    'Preencher custo contratual real');
  DefinirLinhaSimulacao(SIM_ROW_MOB_BAE, 'Mobilizacao BAE', '0',
    'Preencher custo contratual real');
  DefinirLinhaSimulacao(SIM_ROW_MOB_AQUA, 'Mobilizacao AQUA', '0',
    'Preencher custo contratual real');

  AtualizarDiasDisponiveisSimulacao;

  MemoSimulacaoRelatorio.Lines.Text :=
    'Cenario base carregado.' + sLineBreak +
    'Ajuste backlog, custos, disponibilidade e mobilizacao para comparar cenarios.' + sLineBreak +
    'Use "Carregar base real" para calibrar a grade com alocacoes observadas no banco.' + sLineBreak +
    'ROI e payback saem como proxies pela produtividade adicional versus o custo unitario do cenario Atual.';
  StatusBarSimulacao.SimpleText :=
    'Cenario base pronto. Carregue a base real ou revise custos, disponibilidade e mobilizacao antes de comparar.';
end;

procedure TFrmConsultaExecutantesProgramados.InicializarSimulacaoLogistica;
begin
  GridSimulacaoParametros.ColCount := 3;
  GridSimulacaoParametros.RowCount := SIM_ROW_MOB_AQUA + 1;
  GridSimulacaoParametros.FixedRows := 1;
  GridSimulacaoParametros.Options := GridSimulacaoParametros.Options +
    [goEditing, goAlwaysShowEditor];
  GridSimulacaoParametros.Cells[SIM_COL_PARAMETRO, 0] := 'Parametro';
  GridSimulacaoParametros.Cells[SIM_COL_VALOR, 0] := 'Valor';
  GridSimulacaoParametros.Cells[SIM_COL_OBSERVACAO, 0] := 'Observacao';
  GridSimulacaoParametros.ColWidths[SIM_COL_PARAMETRO] := 190;
  GridSimulacaoParametros.ColWidths[SIM_COL_VALOR] := 105;
  GridSimulacaoParametros.ColWidths[SIM_COL_OBSERVACAO] := 210;

  PreencherCenarioBaseSimulacao;
end;

procedure TFrmConsultaExecutantesProgramados.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  FrmDataModule.setFilterDBGrid(DBGridExecutantesProgramados);
  FrmDataModule.setFilterDBGrid(DBGridProgramcaoRT);
  FrmDataModule.setFilterDBGrid(DBGridRTSapImport);
  FrmDataModule.setFilterDBGrid(DBGridResumoFrequencia);
  enderecoColibriRegistro:= registroEndereco('Banco de dados');
  dataInicio.Date:= IncDay(now,1);
  dataFim.Date:= IncDay(now,1);
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
  FrmDataModule.ADOQueryConfigRT.Active:= false;
  FrmDataModule.ADOQueryConfigRT.Active:= true;
  //=====================================
  FrmDataModule.ADOQueryProgramacaoCalendario1.Active:= true;

  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT)) then
  begin
    actStatusSELECIONADO.Enabled:= true;
    actStatusTODOS.Enabled:= true;
    DBGridExecutantesProgramados.ReadOnly:= false;

    actPrepararDadosRT.Enabled:= true;
    actGerarMultiplasRTs.Enabled:= true;
    actCancelarRTForcado.Enabled:= true;
    ConciliarRTsOrfas.Enabled:= true;
    actImportRTSAP.Enabled:= true;
    actLimpar.Enabled:= true;
    actHistorizarProgramacao_DadosSAP.Enabled:= true;
  end
  else if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)) then
  begin
    actStatusSELECIONADO.Enabled:= false;
    actStatusTODOS.Enabled:= false;
    DBGridExecutantesProgramados.ReadOnly:= true;
    actPrepararDadosRT.Enabled:= true;
    actGerarMultiplasRTs.Enabled:= true;
    actCancelarRTForcado.Enabled:= true;
    ConciliarRTsOrfas.Enabled:= true;
    actImportRTSAP.Enabled:= true;
    actLimpar.Enabled:= true;
    actHistorizarProgramacao_DadosSAP.Enabled:= true;
  end
  else
  begin
    actStatusSELECIONADO.Enabled:= false;
    actStatusTODOS.Enabled:= false;
    DBGridExecutantesProgramados.ReadOnly:= true;

    actPrepararDadosRT.Enabled:= false;
    actGerarMultiplasRTs.Enabled:= false;
    actCancelarRTForcado.Enabled:= false;
    ConciliarRTsOrfas.Enabled:= false;
    actImportRTSAP.Enabled:= false;
    actLimpar.Enabled:= false;
    actHistorizarProgramacao_DadosSAP.Enabled:= false;
  end;
  //Incicialização
  actProcurarProgramacaoExecutante.Execute;
  actProcurarProgramacaoRT.Execute;
  actProcurarSAPImport.Execute;
  GarantirTabelaRTRegraRecolhimento;
  InicializarSimulacaoLogistica;
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
    1: actProcurarProgramacaoRT.Execute;
    2: actProcurarSAPImport.Execute;
    3: actProcurarBuscaEmbarque.Execute;
    4: AtualizarDiasDisponiveisSimulacao;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.btnSimulacaoPadraoClick(
  Sender: TObject);
begin
  PreencherCenarioBaseSimulacao;
end;

procedure TFrmConsultaExecutantesProgramados.btnSimulacaoBaseRealClick(
  Sender: TObject);
begin
  try
    CarregarParametrosSimulacaoComBaseReal;
  except
    on E: Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TFrmConsultaExecutantesProgramados.btnSimulacaoRodarClick(
  Sender: TObject);
var
  Param: TSimulacaoParametros;
  Resultado: TSimulacaoResultado;
begin
  try
    Param := LerParametrosSimulacaoAtual;
    Resultado := TSimulacaoLogistica.Simular(Param);

    MemoSimulacaoRelatorio.Lines.Text :=
      'Periodo analisado: ' + FormatDateTime('dd/mm/yyyy', dataInicio.Date) +
      ' a ' + FormatDateTime('dd/mm/yyyy', dataFim.Date) + sLineBreak +
      'Frota configurada: ' + DescreverFrotaSimulacao(Param) +
      sLineBreak + sLineBreak +
      Resultado.Relatorio;

    StatusBarSimulacao.SimpleText := Format(
      '%s | Backlog atendido: %.1f%% | Resiliencia: %.1f%% | Custo/mov: R$ %.2f',
      [TSimulacaoLogistica.PrazoComoTexto(Resultado),
       Resultado.PercentualBacklogAtendido, Resultado.IndiceResiliencia,
       Resultado.CustoPorMovimentacao]);
  except
    on E: Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TFrmConsultaExecutantesProgramados.btnSimulacaoCompararClick(
  Sender: TObject);
var
  Param: TSimulacaoParametros;
  Cenarios: array[0..4] of TSimulacaoCenario;
begin
  try
    Param := LerParametrosSimulacaoAtual;
    MontarCenariosComparativos(Param, Cenarios);

    MemoSimulacaoRelatorio.Lines.Text :=
      'Periodo analisado: ' + FormatDateTime('dd/mm/yyyy', dataInicio.Date) +
      ' a ' + FormatDateTime('dd/mm/yyyy', dataFim.Date) + sLineBreak +
      'Cenario base informado: ' + DescreverFrotaSimulacao(Param) + sLineBreak +
      sLineBreak +
      GerarRelatorioComparativoTexto(Cenarios);

    StatusBarSimulacao.SimpleText :=
      'Comparativo automatico gerado com ROI e payback proxy versus o cenario Atual.';
  except
    on E: Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TFrmConsultaExecutantesProgramados.btnSimulacaoExportarClick(
  Sender: TObject);
var
  Param: TSimulacaoParametros;
  Cenarios: array[0..4] of TSimulacaoCenario;
  TextoExportacao: string;
  Linhas: TStringList;
  Extensao: string;
begin
  try
    SaveDialog1.Title := 'Exportar simulacao logistica';
    SaveDialog1.Filter :=
      'Relatorio de texto (*.txt)|*.txt|Comparativo CSV (*.csv)|*.csv';
    SaveDialog1.DefaultExt := 'txt';
    SaveDialog1.FileName :=
      'simulacao_logistica_' + FormatDateTime('yyyymmdd_hhnnss', Now) + '.txt';

    if not SaveDialog1.Execute then
      Exit;

    Extensao := LowerCase(ExtractFileExt(SaveDialog1.FileName));
    if Extensao = '.csv' then
    begin
      Param := LerParametrosSimulacaoAtual;
      MontarCenariosComparativos(Param, Cenarios);
      TextoExportacao := GerarRelatorioComparativoCSV(Cenarios);
    end
    else
    begin
      TextoExportacao := MemoSimulacaoRelatorio.Lines.Text;
      if Trim(TextoExportacao) = '' then
      begin
        Param := LerParametrosSimulacaoAtual;
        TextoExportacao :=
          'Periodo analisado: ' + FormatDateTime('dd/mm/yyyy', dataInicio.Date) +
          ' a ' + FormatDateTime('dd/mm/yyyy', dataFim.Date) + sLineBreak +
          'Frota configurada: ' + DescreverFrotaSimulacao(Param) +
          sLineBreak + sLineBreak +
          TSimulacaoLogistica.GerarRelatorio(TSimulacaoLogistica.Simular(Param));
      end;
    end;

    Linhas := TStringList.Create;
    try
      Linhas.Text := TextoExportacao;
      Linhas.SaveToFile(SaveDialog1.FileName, TEncoding.UTF8);
    finally
      Linhas.Free;
    end;

    StatusBarSimulacao.SimpleText :=
      'Relatorio exportado para: ' + SaveDialog1.FileName;
  except
    on E: Exception do
      ShowMessage(E.Message);
  end;
end;

procedure TFrmConsultaExecutantesProgramados.ProcessarRetornoImportacaoRTSAP(
  const LinhaLog, AOrigemConsulta: string;
  const APeriodoInicio, APeriodoFim: TDateTime);
var
  Conteudo: string;
  SL: TStringList;
  D: TRTSapImportDados;
  Dt: TDateTime;
begin
  if Trim(LinhaLog) = '' then
    Exit;

  Conteudo := StripPrefixoLog(LinhaLog);
  if Conteudo = '' then
    Exit;

  SL := TStringList.Create;
  try
    SL.StrictDelimiter := True;
    SL.Delimiter := '|';
    SL.DelimitedText := Conteudo;

    if SL.Count = 0 then
      Exit;

    if not SameText(Trim(SL[0]), 'IMPORT') then
      Exit;

    if SL.Count < 17 then
      Exit;

    D := Default(TRTSapImportDados);

    D.DataImportacao := Now;
    D.OrigemConsulta := AOrigemConsulta;
    D.PeriodoInicio := APeriodoInicio;
    D.PeriodoFim := APeriodoFim;

    D.QMNUM := Trim(SL[1]);
    D.QMART := Trim(SL[2]);
    D.IWERK := Trim(SL[3]);
    D.INGRP := Trim(SL[4]);

    D.Origem := Trim(SL[5]);
    D.Destino := Trim(SL[6]);

    if TryParseDataSAPBr(Trim(SL[7]), Dt) then
      D.DataViagem := Dt
    else
      D.DataViagem := 0;

    D.HoraViagem := Trim(SL[8]);

    D.PERNR := Trim(SL[9]);
    D.TipoDoc := Trim(SL[10]);
    D.Documento := Trim(SL[11]);
    D.Passageiro := Trim(SL[12]);
    D.QMTXT := Trim(SL[13]);

    D.RT_Modal := Trim(SL[14]);
    D.RT_Classe := Trim(SL[15]);

    D.StatusItem := Trim(SL[16]);
    D.StatusDescricao := DescricaoStatusRTSAP(D.StatusItem);

    D.idProgramacaoRT := 0;
    D.idProgramacaoExecutante := 0;
    D.ImportadoColibri := False;
    D.Observacao := '';

    InserirOuAtualizarRTSapImport(D);

  finally
    SL.Free;
  end;
end;

procedure TFrmConsultaExecutantesProgramados.RLImpressaoFixedCellClick(
  Sender: TObject; ACol, ARow: Integer);
begin
  FrmPrincipal.clasifica(RLImpressao,ACol,true);
  AutoFitGrade(RLImpressao);
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

function TFrmConsultaExecutantesProgramados.CampoAsString(
  DS: TDataSet; const NomeCampo: string): string;
begin
  Result := '';
  if (DS <> nil) and (DS.FindField(NomeCampo) <> nil) then
    Result := Trim(DS.FieldByName(NomeCampo).AsString);
end;


end.
