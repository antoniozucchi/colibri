unit untGerenciarEmbarcacoes;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics,Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Data.DB, System.Actions,
  Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.OleCtrls, SHDocVw, Vcl.DBCtrls, Vcl.Grids, Vcl.DBGrids,
  Vcl.Mask, Vcl.ToolWin, Registry,DateUtils,Vcl.Buttons,Math, Vcl.Menus,
  Data.Win.ADODB, Vcl.CheckLst,SDL_rchart,SDL_sdlbase, Vcl.ExtDlgs,ComObj, SDL_NumIO,
  Vcl.Samples.Spin,UITYPES,ClipBrd, untDBGridFilter, untConsultaExecutantesProgramados,
  uZucchi, uProgramacaoRTUtils, uDistribuicaoLogistica, uCriacaoRotas;

type
  TFrmGerenciarEmbarcacoes = class(TForm)
    PanelTitulo: TPanel;
    PageControlPrincipal: TPageControl;
    ActionManager1: TActionManager;
    actPanTo: TAction;
    actImprimirMapa: TAction;
    actSalvarImagem: TAction;
    actMarcar: TAction;
    TabSheet3: TTabSheet;
    PanelLateral: TPanel;
    PanelCadastro: TPanel;
    DBGridRoteamento: TFilterDBGrid;
    Splitter1: TSplitter;
    PanelPltaformas: TPanel;
    Splitter2: TSplitter;
    PanelVinculados: TPanel;
    PanelTituloPltaformas: TPanel;
    actLocalizarRoteamentos: TAction;
    actInserirRoteamento: TAction;
    actExcluirRoteamento: TAction;
    actEditarRoteamento: TAction;
    PanelExecutantesVinculados: TPanel;
    actLimparGrafico: TAction;
    actGerarLinhas: TAction;
    actFiltroDestinos: TAction;
    actVincular: TAction;
    actProcuraExecutantes: TAction;
    actExcluirVinculoExecuntanteSelecionado: TAction;
    actAprovarSelecionado: TAction;
    actCancelarSelecionado: TAction;
    ColorBoxLinha: TColorBox;
    ToolBar3: TToolBar;
    btnInserir: TBitBtn;
    btnSalvar: TBitBtn;
    btnExcluirRota: TBitBtn;
    DBEditCodigoRoteamento: TDBEdit;
    PanelTituloCdastro: TPanel;
    PanelAnalise: TPanel;
    actSelProgExec1: TAction;
    actSelProgExec2: TAction;
    actMapaZoomFit: TAction;
    actMapaZoomDinamico: TAction;
    actMapaPan: TAction;
    actMapaZoomWind: TAction;
    SavePictureDialog1: TSavePictureDialog;
    TabSheet2: TTabSheet;
    actRLImprimir_Rota: TAction;
    StatusBarImpressao: TStatusBar;
    actExcelImpressao: TAction;
    actExcelProcuraExecutantes: TAction;
    actGerarLinhasTodas: TAction;
    actImprimirPlanilha: TAction;
    PanelAjuda: TPanel;
    PanelTituloAjuda: TPanel;
    SpeedButton4: TSpeedButton;
    PanelAjudaImpressao: TPanel;
    CheckListBoxImpressao: TCheckListBox;
    ToolBar6: TToolBar;
    btnImprimirRelatorio: TToolButton;
    Memo1: TMemo;
    Panel11: TPanel;
    actKIT_PS: TAction;
    ToolBar7: TToolBar;
    actLanche: TAction;
    actRemoveRow: TAction;
    PopupMenuImpressao: TPopupMenu;
    Excluirlinhaselecionada1: TMenuItem;
    actRemoveCol: TAction;
    RLImpressao: TStringGrid;
    actFinalGridImpressao: TAction;
    Excluircolunaselecionada1: TMenuItem;
    Panel26: TPanel;
    DBNavigator1: TDBNavigator;
    DateTimePickerProgramacao: TDateTimePicker;
    Panel13: TPanel;
    MaskEditHora: TMaskEdit;
    ClassificarCrescentepelacolunaselecionada1: TMenuItem;
    Fixar1Linha1: TMenuItem;
    Fixar1Coluna1: TMenuItem;
    actFixarLinhaColuna: TAction;
    Ajustartamanhodascolunas1: TMenuItem;
    actMapaMundi: TAction;
    PopupMenuMapa: TPopupMenu;
    actMapaMundi1: TMenuItem;
    actLatitudeLongitude: TAction;
    actBrasil: TAction;
    LatitudeseLongitudes1: TMenuItem;
    Brasil1: TMenuItem;
    actSalvarDesenho: TAction;
    SaveDialog1: TSaveDialog;
    actAbrirDesenho: TAction;
    OpenDialog1: TOpenDialog;
    Splitter4: TSplitter;
    Panel4: TPanel;
    ToolBar4: TToolBar;
    Panel5: TPanel;
    ToolBar1: TToolBar;
    actSalvarRota: TAction;
    ToolBar10: TToolBar;
    DBGridRotaExecutantes: TFilterDBGrid;
    StatusBarExecutantes: TStatusBar;
    ToolBar9: TToolBar;
    ToolButton21: TToolButton;
    ToolButton24: TToolButton;
    DBGridProcuraGeral: TFilterDBGrid;
    StatusBarRotaExecutantes: TStatusBar;
    SplitterAnalise: TSplitter;
    actNumerar: TAction;
    PanelTituloExecutantes: TPanel;
    actDisponivel: TAction;
    btnExcluirVinculos: TToolButton;
    PopupMenuExcluirVinculados: TPopupMenu;
    Excluirvinculo1: TMenuItem;
    actExcluirVinculoExecuntanteTodos: TAction;
    Excluirvinculodoexecutantetodos1: TMenuItem;
    actClassificarRegistros: TAction;
    actPreencherRota: TAction;
    DBEdit1: TDBEdit;
    actSelecionados: TAction;
    btnVincluados: TToolButton;
    actMostrarVinculados: TAction;
    StrGridRota: TStringGrid;
    PopupMenuRota: TPopupMenu;
    Excluirlinhaselecionada2: TMenuItem;
    actExcluirLinha: TAction;
    actProcuraProgramados: TAction;
    DBGridProcuraProgramados: TFilterDBGrid;
    btnProgramados: TToolButton;
    actProcuraGeral: TAction;
    actCalcular: TAction;
    BitBtn9: TBitBtn;
    BitBtn5: TBitBtn;
    BitBtn8: TBitBtn;
    BitBtn2: TBitBtn;
    BitBtn7: TBitBtn;
    actAjustar: TAction;
    BitBtn16: TBitBtn;
    RadioGroupImpressao: TRadioGroup;
    actImprimirRotas: TAction;
    BitBtn3: TBitBtn;
    btnDisponivel: TToolButton;
    actNumerarAutomatico: TAction;
    ColunasLayoutRoteamento: TStringGrid;
    ColunasLayoutProcuraGeral: TStringGrid;
    ColunasLayoutProgramados: TStringGrid;
    ColunasLayoutRotaExecutantes: TStringGrid;
    CheckBoxOrigemDestino: TCheckBox;
    actProcuraRotaExecutantes: TAction;
    SpinEditFonte: TSpinEdit;
    actTransbordoInterno: TAction;
    StrGridResumo: TStringGrid;
    actExcelDestinos: TAction;
    ToolButton1: TToolButton;
    btnPendentes: TToolButton;
    btnTotal: TToolButton;
    ComboBoxPlataforma: TComboBox;
    actCopiarCelula: TAction;
    actColarCelula: TAction;
    Copiar1: TMenuItem;
    Colar1: TMenuItem;
    actExcluirCelula: TAction;
    Limpar1: TMenuItem;
    ToolButton2: TToolButton;
    ComboBoxOrigemInicial: TComboBox;
    BitBtn20: TBitBtn;
    BitBtn21: TBitBtn;
    StatusBarRoteamento: TStatusBar;
    actProcuraRoteamento: TAction;
    actMapaRoteiro: TAction;
    ToolButton7: TToolButton;
    ToolButton13: TToolButton;
    actSelRotaExecTODOS: TAction;
    actSelRotaExecLIMPAR: TAction;
    actSelecionadosRotaExecutantes: TAction;
    BitBtn14: TBitBtn;
    PanelTituloRotaExecutantes: TPanel;
    ToolButton6: TToolButton;
    actImprimirExtrato: TAction;
    actListaRTEmbarque: TAction;
    BitBtn12: TBitBtn;
    actExtratoOrigem: TAction;
    BitBtn1: TBitBtn;
    actDuplicados: TAction;
    BitBtn23: TBitBtn;
    PanelTitulo1: TPanel;
    PanelTitulo2: TPanel;
    edtTitulo1: TEdit;
    edtTitulo2: TEdit;
    btnFiltroClearRoteamento: TToolButton;
    btnFiltroClearProcuraGeral: TToolButton;
    btnFiltroClearRotaExecutante: TToolButton;
    btnFiltroClearProgramados: TToolButton;
    btnExcelRotaExecutante: TToolButton;
    btnExcelGeral: TToolButton;
    btnExcelRoteamento: TToolButton;
    TabSheet1: TTabSheet;
    ToolBar2: TToolBar;
    ToolButton23: TToolButton;
    ToolButton5: TToolButton;
    ToolButton8: TToolButton;
    ToolButton10: TToolButton;
    ToolButton3: TToolButton;
    ToolButton11: TToolButton;
    ToolButton12: TToolButton;
    ToolButton4: TToolButton;
    ToolButton22: TToolButton;
    btnCotas: TToolButton;
    ToolButton14: TToolButton;
    ToolButton9: TToolButton;
    ToolButton15: TToolButton;
    ToolButton16: TToolButton;
    StatusBarMapa: TStatusBar;
    MemoMapaMundi: TMemo;
    ToolBar19: TToolBar;
    CheckBox1: TCheckBox;
    ToolButton18: TToolButton;
    edtGridXFEA: TNumIO2;
    edtGridYFEA: TNumIO2;
    ToolButton57: TToolButton;
    edtXFEA: TNumIO2;
    ToolButton60: TToolButton;
    edtYFEA: TNumIO2;
    RChartMapa: TRChart;
    actSugerirRota: TAction;
    actAlocacaoAutomatica: TAction;
    actCriarRotasAutomaticamente: TAction;
    BitBtn4: TBitBtn;
    actExcluirRoteamentoTODOS: TAction;
    BitBtn6: TBitBtn;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actLocalizarRoteamentosExecute(Sender: TObject);
    procedure DBGridRoteamentoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actInserirRoteamentoExecute(Sender: TObject);
    procedure actEditarRoteamentoExecute(Sender: TObject);
    procedure actExcluirRoteamentoExecute(Sender: TObject);
    procedure actFiltroDestinosExecute(Sender: TObject);
    procedure DBGridProcuraGeralDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actProcuraExecutantesExecute(Sender: TObject);
    procedure actVincularExecute(Sender: TObject);
    procedure edtTipoEtapaServicoKeyPress(Sender: TObject; var Key: Char);
    procedure edtOrigemKeyPress(Sender: TObject; var Key: Char);
    procedure PageControlPrincipalChange(Sender: TObject);
    procedure DBGridProcuraGeralCellClick(Column: TColumn);
    procedure ColorBoxLinhaCloseUp(Sender: TObject);
    procedure actSelProgExec1Execute(Sender: TObject);
    procedure actSelProgExec2Execute(Sender: TObject);
    procedure actRLImprimir_RotaExecute(Sender: TObject);
    procedure actExcelImpressaoExecute(Sender: TObject);
    procedure actImprimirPlanilhaExecute(Sender: TObject);
    procedure PanelTituloAjudaMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure PanelTituloAjudaMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure PanelTituloAjudaMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure SpeedButton4Click(Sender: TObject);
    procedure actKIT_PSExecute(Sender: TObject);
    procedure actLancheExecute(Sender: TObject);
    procedure actFinalGridImpressaoExecute(Sender: TObject);
    procedure actRemoveRowExecute(Sender: TObject);
    procedure actRemoveColExecute(Sender: TObject);
    procedure RLImpressaoFixedCellClick(Sender: TObject; ACol, ARow: Integer);
    procedure MaskEditHoraChange(Sender: TObject);
    procedure ClassificarCrescentepelacolunaselecionada1Click(Sender: TObject);
    procedure actFixarLinhaColunaExecute(Sender: TObject);
    procedure Ajustartamanhodascolunas1Click(Sender: TObject);
    procedure actSalvarRotaExecute(Sender: TObject);
    procedure DBGridRotaExecutantesKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGridRotaExecutantesDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actExcluirVinculoExecuntanteSelecionadoExecute(Sender: TObject);
    procedure actDisponivelExecute(Sender: TObject);
    procedure actExcluirVinculoExecuntanteTodosExecute(Sender: TObject);
    procedure actClassificarRegistrosExecute(Sender: TObject);
    procedure actPreencherRotaExecute(Sender: TObject);
    procedure actSelecionadosExecute(Sender: TObject);
    procedure actMostrarVinculadosExecute(Sender: TObject);
    procedure actExcluirLinhaExecute(Sender: TObject);
    procedure StrGridRotaRowMoved(Sender: TObject; FromIndex, ToIndex: Integer);
    procedure actProcuraProgramadosExecute(Sender: TObject);
    procedure actProcuraGeralExecute(Sender: TObject);
    procedure btnProgramadosClick(Sender: TObject);
    procedure DBGridProcuraProgramadosDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actCalcularExecute(Sender: TObject);
    procedure actAjustarExecute(Sender: TObject);
    procedure actImprimirRotasExecute(Sender: TObject);
    procedure actNumerarAutomaticoExecute(Sender: TObject);
    procedure actProcuraRotaExecutantesExecute(Sender: TObject);
    procedure actLimparFiltrosRotaExecute(Sender: TObject);
    procedure StrGridResumoDblClick(Sender: TObject);
    procedure actExcelDestinosExecute(Sender: TObject);
    procedure StrGridRotaSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure ComboBoxPlataformaCloseUp(Sender: TObject);
    procedure ComboBoxPlataformaMouseLeave(Sender: TObject);
    procedure SpinEditFonteChange(Sender: TObject);
    procedure actCopiarCelulaExecute(Sender: TObject);
    procedure actColarCelulaExecute(Sender: TObject);
    procedure actExcluirCelulaExecute(Sender: TObject);
    procedure ComboBoxOrigemInicialCloseUp(Sender: TObject);
    procedure ComboBoxOrigemInicialMouseLeave(Sender: TObject);
    procedure actTransbordoInternoExecute(Sender: TObject);
    procedure actProcuraRoteamentoExecute(Sender: TObject);
    procedure ToolButton4Click(Sender: TObject);
    procedure DBGridRoteamentoKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DBGridRotaExecutantesCellClick(Column: TColumn);
    procedure actSelRotaExecTODOSExecute(Sender: TObject);
    procedure actSelRotaExecLIMPARExecute(Sender: TObject);
    procedure actSelecionadosRotaExecutantesExecute(Sender: TObject);
    procedure ColorBoxLinhaMouseLeave(Sender: TObject);
    procedure actImprimirExtratoExecute(Sender: TObject);
    procedure actListaRTEmbarqueExecute(Sender: TObject);
    procedure actExtratoOrigemExecute(Sender: TObject);
    procedure actDuplicadosExecute(Sender: TObject);
    procedure actImprimirMapaExecute(Sender: TObject);
    procedure actSalvarImagemExecute(Sender: TObject);
    procedure actSalvarDesenhoExecute(Sender: TObject);
    procedure actAbrirDesenhoExecute(Sender: TObject);
    procedure actMarcarExecute(Sender: TObject);
    procedure actLimparGraficoExecute(Sender: TObject);
    procedure actGerarLinhasExecute(Sender: TObject);
    procedure actGerarLinhasTodasExecute(Sender: TObject);
    procedure actMapaZoomFitExecute(Sender: TObject);
    procedure actMapaZoomDinamicoExecute(Sender: TObject);
    procedure actMapaPanExecute(Sender: TObject);
    procedure actMapaZoomWindExecute(Sender: TObject);
    procedure actMapaMundiExecute(Sender: TObject);
    procedure actLatitudeLongitudeExecute(Sender: TObject);
    procedure actBrasilExecute(Sender: TObject);
    procedure actMapaRoteiroExecute(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure RChartMapaDblClick(Sender: TObject);
    procedure RChartMapaKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure RChartMapaMouseLeave(Sender: TObject);
    procedure RChartMapaMouseMove(Sender: TObject; Shift: TShiftState; X,
      Y: Integer);
    procedure RChartMapaZoomPan(Sender: TObject);
    procedure actSugerirRotaExecute(Sender: TObject);
    procedure actAlocacaoAutomaticaExecute(Sender: TObject);
    procedure actCriarRotasAutomaticamenteExecute(Sender: TObject);
    procedure actExcluirRoteamentoTODOSExecute(Sender: TObject);

  private
    booleanZoomMapa,BooleanIndividual: Boolean;
    FUltimoTickProgresso: Cardinal;
    procedure AtualizarProgressoCriacaoRotas(ATotal, AAtual: Integer; const AMensagem: string);
    { Private declarations }

    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure editarRoteamento(Tipo: Integer);
    function verificaAPLAT(NomeExecutante,Funcao: String): Boolean;
    function MontarRotaSequenciaDoGrid: string;

  public
    { Public declarations }

  end;

var
  FrmGerenciarEmbarcacoes: TFrmGerenciarEmbarcacoes;

implementation
  uses untDataModule,UntPrincipal,untFrmPreview, untFrmTabela;
{$R *.dfm}

procedure TFrmGerenciarEmbarcacoes.actAbrirDesenhoExecute(Sender: TObject);
begin
  OpenDialog1.FileName:= 'Desenho.zuchi';
  OpenDialog1.DefaultExt:= 'zuchi';
  OpenDialog1.Filter:= 'Desenho (*.zuchi)|*.zuchi|Todos os Arquivos (*.*)|*.*';
  if OpenDialog1.Execute then
    RChartMapa.LoadData(OpenDialog1.FileName,true);
end;

procedure TFrmGerenciarEmbarcacoes.actAjustarExecute(Sender: TObject);
begin
  AutoFitGrade(RLImpressao);
end;

procedure TFrmGerenciarEmbarcacoes.actAlocacaoAutomaticaExecute(Sender: TObject);
var
  DataProg: string;
  DistLogistica: TDistribuicaoLogistica;
  MsgResultado: string;
  QtdAlocados: Integer;
begin
  if Application.MessageBox('Deseja que a IA tente alocar automaticamente todos os executantes pendentes nas melhores rotas disponíveis?',
                            'Alocação Automática', MB_YESNO + MB_ICONQUESTION) = IDNO then
    Exit;

  DataProg := DateToStr(DateTimePickerProgramacao.Date);

  DistLogistica := TDistribuicaoLogistica.Create(FrmDataModule.ADOConnectionColibri, FrmDataModule.ADOConnectionConsulta);
  try
    FrmPrincipal.ProgressBarIncializa(100, 'Executando algoritmo de otimização...');

    QtdAlocados := DistLogistica.AlocarLoteAutomaticamente(DataProg, MsgResultado);

    FrmPrincipal.ProgressBarAtualizar;

    ShowMessage(MsgResultado);

    if QtdAlocados > 0 then
    begin
      // Atualizar grids
      FrmDataModule.ADOQueryTM_RotaExecutantes.Active := false;
      FrmDataModule.ADOQueryTM_RotaExecutantes.Active := true;
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := false;
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := true;
    end;

  finally
    DistLogistica.Free;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actBrasilExecute(Sender: TObject);
  var
    i: integer;
    ListaXY: TStringList;
    X1,Y1,X2,Y2: Double;
    Endereco: String;
begin
  Endereco := ExtractFilePath(Application.ExeName) + 'Mapas\Mapa_3.txt';
  MemoMapaMundi.Lines.LoadFromFile(Endereco);
  ListaXY:=TStringList.Create;
  ListaXY.delimiter:= ';';
  ListaXY.StrictDelimiter:=False;
  RChartMapa.LineWidth:= 1;
  RChartMapa.PenStyle:= psSolid;
  RChartMapa.DataColor:= clLime;
  for i := 1 to MemoMapaMundi.Lines.Count-1 do
  begin
    ListaXY.Clear;
    ListaXY.DelimitedText:= MemoMapaMundi.Lines[i];
    X1:= StrToFloat(ListaXY[0]);
    Y1:= StrToFloat(ListaXY[1]);
    X2:= StrToFloat(ListaXY[2]);
    Y2:= StrToFloat(ListaXY[3]);
    RChartMapa.MoveTo(X1,Y1);
    RChartMapa.DrawTo(X2,Y2);
    //RChartMapa.MarkAt(X,Y,24);
  end;
  actMapaZoomFit.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actEditarRoteamentoExecute(Sender: TObject);
begin
  editarRoteamento(0);
  actSalvarRota.Execute;
  FrmDataModule.ADOQueryRoteamento.Refresh;
end;

procedure TFrmGerenciarEmbarcacoes.actExcelDestinosExecute(Sender: TObject);
begin
  ExcelStringGrid(StrGridResumo,'Destinos',
  '','','',false,
    FrmPrincipal.ProgressBarPrincipal,FrmPrincipal.MemoPrincipal);
end;

procedure TFrmGerenciarEmbarcacoes.actExcelImpressaoExecute(Sender: TObject);
begin
  ExcelStringGrid(RLImpressao,'Relatório de Impressão',
  edtTitulo1.Text,edtTitulo2.Text,'',false,
    FrmPrincipal.ProgressBarPrincipal,FrmPrincipal.MemoPrincipal);
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirCelulaExecute(Sender: TObject);
begin
  RLImpressao.Cells[RLImpressao.Col,RLImpressao.Row]:= '';
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirLinhaExecute(Sender: TObject);
begin
  FrmPrincipal.DeleteRow(StrGridRota,StrGridRota.Row);
  actSalvarRota.Execute;
  AutoFitGrade(StrGridRota);
  StrGridRota.ColWidths[0]:= 20;
  StrGridRota.ColWidths[1]:= 80;
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirRoteamentoExecute(Sender: TObject);
begin
  if Application.MessageBox(PChar(
  'Deseja realmente excluir a rota selecionada?'),'.::ATENÇÃO::.',36) = 6 then
  begin
    actExcluirVinculoExecuntanteTodos.Execute;
    FrmDataModule.ADOQueryRoteamento.Delete;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirRoteamentoTODOSExecute(
  Sender: TObject);
begin
  if Application.MessageBox(PChar(
  'Deseja realmente excluir todas as rotas filtradas?'),'.::ATENÇÃO::.',36) = 6 then
  begin
    FrmDataModule.ADOQueryRoteamento.First;
    while not FrmDataModule.ADOQueryRoteamento.Eof do
    begin
      actExcluirVinculoExecuntanteTodos.Execute;
      FrmDataModule.ADOQueryRoteamento.Delete;
    end;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirVinculoExecuntanteSelecionadoExecute(Sender: TObject);
var
  idAux_Rota_Distribuicao, CodigoProgramacaoExecutante, NumRegistros: Integer;
  DistLogistica: TDistribuicaoLogistica;
  MensagemErro: string;
begin
  DistLogistica := TDistribuicaoLogistica.Create(FrmDataModule.ADOConnectionColibri, FrmDataModule.ADOConnectionConsulta);
  try
    NumRegistros := FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount;
    FrmPrincipal.ProgressBarIncializa(NumRegistros, 'Excluindo executantes selecionados...');

    FrmDataModule.ADOQueryTM_RotaExecutantes.First;
    while not FrmDataModule.ADOQueryTM_RotaExecutantes.Eof do
    begin
      if FrmDataModule.DataSourceTM_RotaExecutantes.DataSet.FieldByName('booleanSelecao').AsBoolean then
      begin
        idAux_Rota_Distribuicao := FrmDataModule.DataSourceTM_RotaExecutantes.DataSet.FieldByName('idAux_Rota_Distribuicao').AsInteger;
        CodigoProgramacaoExecutante := FrmDataModule.DataSourceTM_RotaExecutantes.DataSet.FieldByName('CodigoProgramacaoExecutante').AsInteger;

        // Usar a classe de regras de negócio para desvincular
        if DistLogistica.DesvincularExecutanteDeRota(idAux_Rota_Distribuicao, MensagemErro) then
        begin
          // Atualizar status do executante para não inserido
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value := CodigoProgramacaoExecutante;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;

          if not FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.IsEmpty then
          begin
            FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
            FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean := false;
            FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
          end;
        end;
      end;

      FrmDataModule.ADOQueryTM_RotaExecutantes.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
  finally
    DistLogistica.Free;
  end;

  // Atualizar Grids
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active := false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active := true;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := true;

  NumRegistros := FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount;
  StatusBarRotaExecutantes.Panels[0].Text := 'Nº registros: ' + IntToStr(NumRegistros);
  AutoFitStatusBar(StatusBarRotaExecutantes);
end;

procedure TFrmGerenciarEmbarcacoes.actExcluirVinculoExecuntanteTodosExecute(
  Sender: TObject);
  var
    idAux_Rota_Distribuicao,CodigoProgramacaoExecutante,NumRegistros: Integer;
begin
  FrmDataModule.ADOQueryTM_RotaExecutantes.First;
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount,
  'Excluindo todos os executantes vinculados...');
  while not FrmDataModule.ADOQueryTM_RotaExecutantes.Eof do
  begin
    try
      //Excluir vincluo
      idAux_Rota_Distribuicao:= FrmDataModule.DataSourceTM_RotaExecutantes.
      DataSet.FieldByName('idAux_Rota_Distribuicao').AsInteger;
      FrmDataModule.ADOQueryExcluirVinculoExecutante.Active:= false;
      FrmDataModule.ADOQueryExcluirVinculoExecutante.Parameters.Items[0].Value:= idAux_Rota_Distribuicao;
      FrmDataModule.ADOQueryExcluirVinculoExecutante.Active:= true;
      FrmDataModule.ADOQueryExcluirVinculoExecutante.Delete;
      //Atualzar InseridoProgramacaoTransporte
      CodigoProgramacaoExecutante:= FrmDataModule.DataSourceTM_RotaExecutantes.
      DataSet.FieldByName('CodigoProgramacaoExecutante').AsInteger;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= false;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
      CodigoProgramacaoExecutante;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active:= true;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
      FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
      FieldByName('InseridoProgramacaoTransporte').AsBoolean:= false;
      FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;
      FrmPrincipal.ProgressBarIncremento(1);
    except
      FrmDataModule.ADOQueryTM_RotaExecutantes.Next;
    end;
  end;
  //Atualizar de executantes vinculados
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= true;
  NumRegistros:= FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount;
  StatusBarRotaExecutantes.Panels[0].Text:= 'N° registros: '+intToStr(NumRegistros);
  AutoFitStatusBar(StatusBarRotaExecutantes);
  //Atualizar de executantes procurados
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= true;
  NumRegistros:= FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecordCount;
  StatusBarExecutantes.Panels[0].Text:= 'N° registros: '+intToStr(NumRegistros);
  AutoFitStatusBar(StatusBarExecutantes);
  actFiltroDestinos.Execute;
  actDisponivel.Execute;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmGerenciarEmbarcacoes.actProcuraExecutantesExecute(
  Sender: TObject);
begin
  if btnProgramados.Down then
  begin
    actProcuraProgramados.Execute;
  end
  else
  begin
    actProcuraGeral.Execute;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actImprimirExtratoExecute(Sender: TObject);
  var
    DataProcura: String;
begin
  if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;

  actFiltroDestinos.Execute;
  DataProcura:= FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date);
  FrmPreview.ListaTituloRelatorio1:= TStringList.Create;
  FrmPreview.ListaTituloRelatorio1.Add('Extrato de passageiros');
  FrmPreview.ListaTituloRelatorio1.Add(DataProcura);
  FrmPreview.RadioGroupRelatorio.ItemIndex:= 7;
  FrmPreview.GeneratePages;
end;

procedure TFrmGerenciarEmbarcacoes.actImprimirMapaExecute(Sender: TObject);
begin
  {if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;
  try
    FrmPreview.ListaTituloRelatorio:= TStringList.Create;
    FrmPreview.ListaTituloRelatorio.Add('Embarcação: '+
    FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('NomeEmbarcacao').AsString);
    FrmPreview.ListaTituloRelatorio.Add('DATA: '+
    FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('DataRoteamento').AsString);
    if StatusBarMapa.Panels[0].Text<>'' then
      FrmPreview.ListaTituloRelatorio.Add(StatusBarMapa.Panels[0].Text);
    FrmPreview.RadioGroupRelatorio.ItemIndex:= 0;
    FrmPreview.PrintPreview.PrintJobTitle:= 'Embarcação: '+
    FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('NomeEmbarcacao').AsString;
    FrmPreview.GeneratePages;
  finally
    RChartMapa.Align:= alClient;
    RChartMapa.ChartColor:= clGradientInactiveCaption;
  end;}
end;

procedure TFrmGerenciarEmbarcacoes.actImprimirPlanilhaExecute(Sender: TObject);
  var
    DataProcura: String;
begin
  if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;

  FrmPreview.ListaTituloRelatorio1:= TStringList.Create;
  FrmPreview.ListaTituloRelatorio1.Add(edtTitulo1.Text);
  FrmPreview.ListaTituloRelatorio1.Add(edtTitulo2.Text);
  FrmPreview.RadioGroupRelatorio.ItemIndex:= 1;
  DataProcura:= FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date);

  case RadioGroupImpressao.ItemIndex of
    0: FrmPreview.PrintPreview.PrintJobTitle:= edtTitulo1.Text;
    1: FrmPreview.PrintPreview.PrintJobTitle:= edtTitulo1.Text+' - ' + DataProcura;
    2: FrmPreview.PrintPreview.PrintJobTitle:= edtTitulo1.Text+' - ' + DataProcura;
    3: FrmPreview.PrintPreview.PrintJobTitle:= edtTitulo1.Text+' - ' + DataProcura;
  end;
  FrmPreview.GeneratePages;
end;

procedure TFrmGerenciarEmbarcacoes.actImprimirRotasExecute(Sender: TObject);
  var
    DataProcura: String;
begin
  if not Assigned(FrmPreview) then
    FrmPreview:=TFrmPreview.Create(Application)
  else
    FrmPreview.Show;

  FrmPreview.PrintPreview.Clear;
  FrmPreview.btnExcluirAntes.Down:= false;
  FrmPreview.RadioGroupRelatorio.ItemIndex:= 1;
  DataProcura:= FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date);
  FrmDataModule.ADOQueryRoteamento.First;
  while not FrmDataModule.ADOQueryRoteamento.Eof do
  begin
    actRLImprimir_Rota.Execute;
    FrmPreview.ListaTituloRelatorio1:= TStringList.Create;
    FrmPreview.ListaTituloRelatorio1.Add(edtTitulo1.Text);
    FrmPreview.ListaTituloRelatorio1.Add(edtTitulo2.Text);
    //=================================================
    FrmPreview.GeneratePages;
    FrmDataModule.ADOQueryRoteamento.Next;
  end;
  FrmPreview.PrintPreview.PrintJobTitle:= 'LANCHAS - '+DataProcura;
end;

procedure TFrmGerenciarEmbarcacoes.actInserirRoteamentoExecute(Sender: TObject);
var
    numLinha: Integer;
    Plataforma,Latitude,Longitude: string;
    ListaOrigem: TStringList;
begin
  editarRoteamento(1);
  ListaOrigem:= TStringList.Create;
  ListaOrigem.delimiter:= ';';
  ListaOrigem.StrictDelimiter:= true;
  ListaOrigem.DelimitedText:= FrmPrincipal.OrigemPlataformas;
  //Incluir primeira linha com uma origem
  numLinha:= StrGridRota.RowCount;
  StrGridRota.RowCount:= numLinha+1;
  Plataforma:= ListaOrigem[0];
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := false;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Parameters.Items[0].Value:= Plataforma;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := true;
  Latitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
  FieldByName('Latitude').AsString;
  Longitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
  FieldByName('Longitude').AsString;
  StrGridRota.Cells[1,numLinha]:= Plataforma;
  StrGridRota.Cells[5,numLinha]:= Latitude;
  StrGridRota.Cells[6,numLinha]:= Longitude;
  actSalvarRota.Execute;
  AutoFitGrade(StrGridRota);
  StrGridRota.ColWidths[0]:= 20;
  StrGridRota.ColWidths[1]:= 80;
  try
    StrGridRota.FixedRows:= 1;
  except
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actExtratoOrigemExecute(Sender: TObject);
  var
    DataProcura,SQLBase,NomeExecutante,Funcao,Origem,SQLOrigem,
    txtDestino,txtTipoEtapaServico,Empresa: String;
    numLinhas: integer;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  SQLOrigem:= FrmPrincipal.palavraBusca('','tblProgramacaoExecutante.Origem','Exato',
  FrmPrincipal.OrigemPlataformas,'AND');
  if SQLOrigem <> '' then
    SQLOrigem:='AND'+SQLOrigem;
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+
  ')AND(tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem)'+SQLOrigem+
  'AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado"))'+
  ' ORDER BY NomeExecutante,txtTipoEtapaServico;';
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  //==============================================
  edtTitulo1.Text:= 'Extrato de executantes (Origem Cadastro)';
  edtTitulo2.Text:= 'DATA: '+ DataProcura;
  RLImpressao.FixedRows:= 0;
  RLImpressao.RowCount:= 1;
  RLImpressao.ColCount:= 5;
  RLImpressao.Cells[0,0]:= 'Nome do Executante';
  RLImpressao.Cells[1,0]:= 'Tipo de Etapa de Serviço';
  RLImpressao.Cells[2,0]:= 'Empresa';
  RLImpressao.Cells[3,0]:= 'Origem';
  RLImpressao.Cells[4,0]:= 'Destino';
  //====================================================
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount,
  'Calculando...');
  FrmDataModule.ADOQueryTemporarioDBColibri.First;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    NomeExecutante:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('NomeExecutante').AsString;
    Funcao:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('Funcao').AsString;
    Origem:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('Origem').AsString;
    txtDestino:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('txtDestino').AsString;
    txtTipoEtapaServico:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('txtTipoEtapaServico').AsString;
    Empresa:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.
    FieldByName('Empresa').AsString;
    if verificaAPLAT(NomeExecutante,Funcao) then
    begin
      numLinhas:= RLImpressao.RowCount;
      RLImpressao.RowCount:= numLinhas+1;
      RLImpressao.Cells[0,numLinhas]:= NomeExecutante;
      RLImpressao.Cells[1,numLinhas]:= txtTipoEtapaServico;
      RLImpressao.Cells[2,numLinhas]:= Empresa;
      RLImpressao.Cells[3,numLinhas]:= Origem;
      RLImpressao.Cells[4,numLinhas]:= txtDestino;
    end;
    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  StatusBarImpressao.Panels[0].Text:= 'N° registros: '+
  intToStr(RLImpressao.RowCount-1);
  actFinalGridImpressao.Execute;
  AutoFitGrade(RLImpressao);
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmGerenciarEmbarcacoes.actColarCelulaExecute(Sender: TObject);
begin
  RLImpressao.Cells[RLImpressao.Col,RLImpressao.Row]:= Clipboard.AsText;
end;

procedure TFrmGerenciarEmbarcacoes.actClassificarRegistrosExecute(Sender: TObject);
  var
    i,contRegistro,ACol: Integer;
    Incial: String;
begin
  //=============================================================
  FrmPrincipal.ProgressBarIncializa(RLImpressao.RowCount,
  'Numerar registros...');
  Incial:= '1';
  if InputQuery('Numerar registros','Entre com o número inicial:',Incial) then
  begin
    contRegistro:= StrToInt(Incial);
    ACol:= 7;
    for i := 0 to RLImpressao.ColCount-1 do
    begin
      if RLImpressao.Cells[i,0] = 'N°' then
      begin
        ACol:= i;
        break;
      end;
    end;
    for i := RLImpressao.Row to RLImpressao.RowCount-1 do
    begin
      RLImpressao.Cells[ACol,i]:= (IntToStr(contRegistro));
      FrmPrincipal.ProgressBarIncremento(1);
      contRegistro:= contRegistro + 1;
    end;
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actCopiarCelulaExecute(Sender: TObject);
begin
  Clipboard.AsText:= RLImpressao.Cells[RLImpressao.Col,RLImpressao.Row];
end;

procedure TFrmGerenciarEmbarcacoes.actCriarRotasAutomaticamenteExecute(Sender: TObject);
var
  DataProg: string;
  CriadorRotas: TCriacaoRotas;
  MsgResultado: string;
  EmbarcacoesDisponiveis: TStringList;
  EmbarcacoesArray: TArray<string>;
  I: Integer;
  Qry: TADOQuery;
begin
  DataProg := FormatDateTime('dd/mm/yyyy', DateTimePickerProgramacao.Date);

  if Application.MessageBox(
       PChar('Deseja criar rotas automaticamente para a data ' + DataProg + '?'),
       'Criação Automática',
       MB_YESNO + MB_ICONQUESTION
     ) = IDNO then
    Exit;

  Screen.Cursor := crHourGlass;

  EmbarcacoesDisponiveis := TStringList.Create;
  CriadorRotas := TCriacaoRotas.Create(
    FrmDataModule.ADOConnectionColibri,
    FrmDataModule.ADOConnectionConsulta
  );
  Qry := TADOQuery.Create(nil);
  try
    Qry.Connection := FrmDataModule.ADOConnectionConsulta;
    Qry.SQL.Text :=
      'SELECT NomeEmbarcacao ' +
      'FROM tblEmbarcacao ' +
      'WHERE Distribuicao = True ' +
      'ORDER BY NomeEmbarcacao';
    Qry.Open;

    while not Qry.Eof do
    begin
      EmbarcacoesDisponiveis.Add(Trim(Qry.FieldByName('NomeEmbarcacao').AsString));
      Qry.Next;
    end;

    if EmbarcacoesDisponiveis.Count = 0 then
    begin
      ShowMessage('Nenhuma embarcação marcada para distribuição foi encontrada.');
      Exit;
    end;

    SetLength(EmbarcacoesArray, EmbarcacoesDisponiveis.Count);
    for I := 0 to EmbarcacoesDisponiveis.Count - 1 do
      EmbarcacoesArray[I] := EmbarcacoesDisponiveis[I];

    FrmPrincipal.ProgressBarIncializa(100, 'Criando rotas e alocando executantes...');
    Application.ProcessMessages;

    FUltimoTickProgresso := 0;
    CriadorRotas.OnProgresso := AtualizarProgressoCriacaoRotas;

    if CriadorRotas.CriarEAlocarLote(
         DataProg,
         EmbarcacoesArray,
         FrmPrincipal.OrigemPlataformas(),
         MsgResultado
       ) then
    begin
      FrmDataModule.ADOQueryRoteamento.Close;
      FrmDataModule.ADOQueryRoteamento.Open;

      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Close;
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Open;

      ShowMessage(MsgResultado);
    end
    else
      ShowMessage('Erro: ' + MsgResultado);

  finally
    Screen.Cursor := crDefault;
    FrmPrincipal.ProgressBarAtualizar;
    Qry.Free;
    CriadorRotas.Free;
    EmbarcacoesDisponiveis.Free;
    actFiltroDestinos.Execute;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actKIT_PSExecute(Sender: TObject);
  var
    DataProcura,SQLBase: String;
    numLinhas: Integer;
begin
  RadioGroupImpressao.ItemIndex:= 3;
  //Procurar todos os Executantes Programados para a DataProcura
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+') '+
  'AND (tblProgramacaoExecutante.Kit_PS LIKE TRUE)) '+
  'ORDER BY NomeExecutante,Origem,txtDestino;';
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  //==============================================
  edtTitulo1.Text:= 'Lista de KIT Embarque';
  edtTitulo2.Text:= 'DATA: '+ DataProcura;
  RLImpressao.FixedRows:= 0;
  RLImpressao.RowCount:= 1;
  RLImpressao.ColCount:= 8;
  RLImpressao.Cells[0,0]:= 'N°';
  RLImpressao.Cells[1,0]:= 'Origem';
  RLImpressao.Cells[2,0]:= 'Destino';
  RLImpressao.Cells[3,0]:= 'KIT Embarque';
  RLImpressao.Cells[4,0]:= 'Responsável';
  RLImpressao.Cells[5,0]:= 'Tipo de Etapa de Serviço';
  RLImpressao.Cells[6,0]:= 'Retirada';
  RLImpressao.Cells[7,0]:= 'Devolução';
  //====================================================
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount,
  'Carregando Lista distribuição KIT Embarque...');
  FrmDataModule.ADOQueryTemporarioDBColibri.First;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    numLinhas:= RLImpressao.RowCount;
    RLImpressao.RowCount:= numLinhas+1;
    //Preencher valores
    RLImpressao.Cells[0,numLinhas]:= (FormatFloat('00',numLinhas));
    RLImpressao.Cells[1,numLinhas]:= (FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('Origem').AsString);
    RLImpressao.Cells[2,numLinhas]:= (FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('txtDestino').AsString);
    RLImpressao.Cells[3,numLinhas]:= (FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('Origem').AsString);
    RLImpressao.Cells[4,numLinhas]:= (FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('NomeExecutante').AsString);
    RLImpressao.Cells[5,numLinhas]:= (FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('txtTipoEtapaServico').AsString);
    RLImpressao.Cells[6,numLinhas]:= '';
    RLImpressao.Cells[7,numLinhas]:= '';

    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  StatusBarImpressao.Panels[0].Text:= 'N° registros: '+
  intToStr(RLImpressao.RowCount-1);
  actFinalGridImpressao.Execute;
  AutoFitGrade(RLImpressao);
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmGerenciarEmbarcacoes.actCalcularExecute(Sender: TObject);
  var
    i: Integer;
    SQLBase,CodigoRoteamento,Destino: String;
    Long1,Long2,Lat1,Lat2,Distancia,SomaDistancia: Double;
begin
  try
    CodigoRoteamento:= DBEditCodigoRoteamento.Text;
    SomaDistancia:= 0;
    Lat1:= StrToFloat(StrGridRota.Cells[5,1]);
    Long1:= StrToFloat(StrGridRota.Cells[6,1]);
    for i := 1 to StrGridRota.RowCount-1 do
    begin
      try
        //Número de Executantes para o destino selecionado
        Destino:= StrGridRota.Cells[1,i];
        SQLBase:= 'SELECT tblRoteamento.*, tblAux_Rota_Distribuicao.*, tblProgramacaoExecutante.*, tblProgramacaoDiaria.* '+
        'FROM tblProgramacaoDiaria INNER JOIN (tblRoteamento INNER JOIN (tblProgramacaoExecutante INNER JOIN '+
        'tblAux_Rota_Distribuicao ON tblProgramacaoExecutante.idProgramacaoExecutante = '+
        'tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ON tblRoteamento.idRoteamento = '+
        'tblAux_Rota_Distribuicao.CodigoRota) ON tblProgramacaoDiaria.idProgramacaoDiaria = '+
        'tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
        'WHERE ((tblAux_Rota_Distribuicao.CodigoRota LIKE  '+CodigoRoteamento+
        ') AND (txtDestino LIKE '+QuotedStr(Destino)+'))';
        FrmDataModule.ADOQueryTemporarioDBColibri.Close;
        FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
        FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
        FrmDataModule.ADOQueryTemporarioDBColibri.Open;
        StrGridRota.Cells[2,i]:= IntToStr(FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount);
        //Calculo da distancia
        Lat2:= StrToFloat(StrGridRota.Cells[5,i]);
        Long2:= StrToFloat(StrGridRota.Cells[6,i]);
        if i > 1 then
        begin
          Distancia:= RoundTo(TProgramacaoRTUtils.GetDistanceBetween(Lat1,Long1,Lat2,Long2),-3);
          Long1:= Long2;
          Lat1:= Lat2;
          SomaDistancia:= SomaDistancia+Distancia;
          StrGridRota.Cells[3,i]:= FormatFloat('0.00',Distancia);
          StrGridRota.Cells[4,i]:= FormatFloat('0.00',SomaDistancia);
        end
        else
        begin
          StrGridRota.Cells[3,i]:= FormatFloat('0.00',0);
          StrGridRota.Cells[4,i]:= FormatFloat('0.00',0);
        end;
      except

      end;
    end;
  except

  end;
  AutoFitGrade(StrGridRota);
  StrGridRota.ColWidths[0]:= 20;
  StrGridRota.ColWidths[1]:= 80;
  actDisponivel.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actDisponivelExecute(Sender: TObject);
  var
    i: Integer;
    Origem,Destino: String;
begin
  if btnDisponivel.Down then
  begin
    try
      Origem:= StrGridRota.Cells[1,1];
      Destino:= '';
      for I := 2 to StrGridRota.RowCount-1 do
      begin
        if Destino <> '' then
          Destino:= Destino+';'+StrGridRota.Cells[1,i]
        else
          Destino:= StrGridRota.Cells[1,i];
      end;
      FrmPrincipal.buscaFiledGrid1('InseridoProgramacaoTransporte','False','Exato',ColunasLayoutProcuraGeral,4,false);
      FrmPrincipal.buscaFiledGrid1('txtDestino',Destino,'Exato',ColunasLayoutProcuraGeral,4,false);
      FrmPrincipal.buscaFiledGrid1('Origem',Origem,'Exato',ColunasLayoutProcuraGeral,4,false);
      FrmPrincipal.buscaFiledGrid1('StatusProgramacao','Aprovado','Exato',ColunasLayoutProcuraGeral,4,false);
      actProcuraExecutantes.Execute;
    except
    end;
  end
  else
  begin
    FrmPrincipal.buscaFiledGrid1('InseridoProgramacaoTransporte','','',ColunasLayoutProcuraGeral,4,false);
    FrmPrincipal.buscaFiledGrid1('Origem','','',ColunasLayoutProcuraGeral,4,false);
    FrmPrincipal.buscaFiledGrid1('txtDestino','','',ColunasLayoutProcuraGeral,4,false);
    FrmPrincipal.buscaFiledGrid1('StatusProgramacao','','',ColunasLayoutProcuraGeral,4,false);
    FrmPrincipal.buscaFiledGrid1('CodigoSAP','','',ColunasLayoutProcuraGeral,4,false);
    actProcuraExecutantes.Execute;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actDuplicadosExecute(Sender: TObject);
  var
    DataProcura: String;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  edtTitulo1.Text:= 'Lista de Transbordos Internos';
  edtTitulo2.Text:= 'DATA: '+ DataProcura;
  TProgramacaoRTUtils.ProgramacaoTransbordo(RLImpressao,DataProcura);
end;

procedure TFrmGerenciarEmbarcacoes.actLancheExecute(Sender: TObject);
  var
    DataProcura,SQLBase,DestinoAnterior,Destino,Origem,SQLString: String;
    numLinhas,ContLanche,TotalExec: Integer;
begin
  Origem:= '';
  if InputQuery('Filtrar Origem',
  'Entre com as origens a serem filtradas separadas por ";"',Origem) then
  begin
    SQLString:= '';
    if Origem <> '' then
      SQLString:= FrmPrincipal.palavraBusca(SQLString,'Origem','Contem',Origem,'AND');
    if SQLString <> '' then
      SQLString:= 'AND'+SQLString;
    RadioGroupImpressao.ItemIndex:= 2;
    //Procurar todos os Executantes Programados para a DataProcura
    DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
    SQLBase:=
    'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
    'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
    'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
    'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+
    ')AND(tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem)'+SQLString+
    'AND(tblProgramacaoExecutante.StatusProgramacao LIKE '+QuotedStr('Aprovado')+
    ')) ORDER BY tblProgramacaoDiaria.txtDestino, tblProgramacaoExecutante.Origem DESC, '+
    'tblProgramacaoDiaria.txtTipoEtapaServico;';
    FrmDataModule.ADOQueryTemporarioDBColibri.Close;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
    FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
    FrmDataModule.ADOQueryTemporarioDBColibri.Open;
    //==============================================
    edtTitulo1.Text:= 'Lista de Distribuição de Lanches';
    edtTitulo2.Text:= 'DATA: '+ DataProcura;
    RLImpressao.FixedRows:= 0;
    RLImpressao.RowCount:= 1;
    RLImpressao.ColCount:= 5;
    RLImpressao.Cells[0,0]:= 'Origem';
    RLImpressao.Cells[1,0]:= 'Destino';
    RLImpressao.Cells[2,0]:= 'Nome do Executante';
    RLImpressao.Cells[3,0]:= 'Empresa';
    RLImpressao.Cells[4,0]:= 'Assinatura';
    //====================================================
    ContLanche:= 0;
    DestinoAnterior:= FrmDataModule.
    DataSourceTemporarioDBColibri.DataSet.FieldByName('txtDestino').AsString;
    TotalExec:= 0;
    FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTemporarioDBColibri.RecordCount,
    'Gerando Lista de Distribuição de Lanches');
    FrmDataModule.ADOQueryTemporarioDBColibri.First;
    while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
    begin
      numLinhas:= RLImpressao.RowCount;
      RLImpressao.RowCount:= numLinhas+1;
      Destino:= FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('txtDestino').AsString;
      //=====================================
      ContLanche:= ContLanche+1;
      if (DestinoAnterior <> Destino) then
      begin
        RLImpressao.Rows[numLinhas].Clear;
        RLImpressao.Cells[0,numLinhas]:= ('TOTAL');
        RLImpressao.Cells[1,numLinhas]:= (IntToStr(ContLanche-1));
        ContLanche:= 0;
      end
      else//Preencher valores
      begin
        RLImpressao.Cells[0,numLinhas]:= (FrmDataModule.
        DataSourceTemporarioDBColibri.DataSet.FieldByName('Origem').AsString);
        RLImpressao.Cells[1,numLinhas]:= (Destino);
        RLImpressao.Cells[2,numLinhas]:= (FrmDataModule.
        DataSourceTemporarioDBColibri.DataSet.FieldByName('NomeExecutante').AsString);
        RLImpressao.Cells[3,numLinhas]:= (FrmDataModule.
        DataSourceTemporarioDBColibri.DataSet.FieldByName('Empresa').AsString);
        RLImpressao.Cells[4,numLinhas]:=  '';
        TotalExec:= TotalExec+1;
        FrmDataModule.ADOQueryTemporarioDBColibri.Next;
      end;
      DestinoAnterior:= Destino;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
    numLinhas:= RLImpressao.RowCount;
    RLImpressao.RowCount:= numLinhas+1;
    RLImpressao.Cells[0,numLinhas]:= ('TOTAL');
    RLImpressao.Cells[1,numLinhas]:= (IntToStr(ContLanche));
    RLImpressao.Cells[2,numLinhas]:= '';
    RLImpressao.Cells[3,numLinhas]:= '';
    RLImpressao.Cells[4,numLinhas]:= '';
    numLinhas:= RLImpressao.RowCount;
    RLImpressao.RowCount:= numLinhas+1;
    RLImpressao.Cells[0,numLinhas]:= char(931)+' TOTAL';
    RLImpressao.Cells[1,numLinhas]:= (IntToStr(TotalExec));
    RLImpressao.Cells[2,numLinhas]:= '';
    RLImpressao.Cells[3,numLinhas]:= '';
    RLImpressao.Cells[4,numLinhas]:= '';

    StatusBarImpressao.Panels[0].Text:= 'N° registros: '+
    intToStr(TotalExec);
    actFinalGridImpressao.Execute;
    AutoFitGrade(RLImpressao);
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actLatitudeLongitudeExecute(Sender: TObject);
  var
    i: integer;
    ListaXY: TStringList;
    X1,Y1,X2,Y2: Double;
    Endereco: String;
begin
  Endereco := ExtractFilePath(Application.ExeName) + 'Mapas\Mapa_2.txt';
  MemoMapaMundi.Lines.LoadFromFile(Endereco);
  ListaXY:=TStringList.Create;
  ListaXY.delimiter:= ';';
  ListaXY.StrictDelimiter:=False;
  RChartMapa.LineWidth:= 1;
  RChartMapa.PenStyle:= psDashDot;
  RChartMapa.FillColor:= RChartMapa.ChartColor;
  RChartMapa.DataColor:= clGray;
  for i := 1 to MemoMapaMundi.Lines.Count-1 do
  begin
    ListaXY.Clear;
    ListaXY.DelimitedText:= MemoMapaMundi.Lines[i];
    X1:= StrToFloat(ListaXY[0]);
    Y1:= StrToFloat(ListaXY[1]);
    X2:= StrToFloat(ListaXY[2]);
    Y2:= StrToFloat(ListaXY[3]);
    RChartMapa.MoveTo(X1,Y1);
    RChartMapa.DrawTo(X2,Y2);
  end;
  //Escala de Latitude
  RChartMapa.TextBkStyle:= tbClear;
  RChartMapa.TextAlignment:= taCenter;
  RChartMapa.TextBkColor:= RChartMapa.ChartColor;
  RChartMapa.DataColor:= clFuchsia;
  for i := 0 to 12 do
  begin
    RChartMapa.Text(i*1669.793,-7500,10,intToStr(i*15));
    RChartMapa.Text(i*1669.793,12900,10,intToStr(i*15));
  end;
  for i := 0 to 11 do
  begin
    RChartMapa.Text(-i*1669.793,-7500,10,intToStr(-i*15));
    RChartMapa.Text(-i*1669.793,12900,10,intToStr(-i*15));
  end;
  //Escala de Longitude
  RChartMapa.TextAlignment:= taRightJustify;
  RChartMapa.Text(-18735,-5411.9749,10,'-45');
  RChartMapa.Text(-18735,-3430.0606,10,'-30');
  RChartMapa.Text(-18735,-1708.9247,10,'-15');
  RChartMapa.Text(-18735,0,10,'0');
  RChartMapa.Text(-18735,1681.1916,10,'15');
  RChartMapa.Text(-18735,3426.7500,10,'30');
  RChartMapa.Text(-18735,5384.2417,10,'45');
  RChartMapa.Text(-18735,7626.9341,10,'60');
  RChartMapa.Text(-18735,10495.4941,10,'75');
  RChartMapa.TextAlignment:= taLeftJustify;
  RChartMapa.Text(21346,-5411.9749,10,'-45');
  RChartMapa.Text(21346,-3430.0606,10,'-30');
  RChartMapa.Text(21346,-1708.9247,10,'-15');
  RChartMapa.Text(21346,0,10,'0');
  RChartMapa.Text(21346,1681.1916,10,'15');
  RChartMapa.Text(21346,3426.7500,10,'30');
  RChartMapa.Text(21346,5384.2417,10,'45');
  RChartMapa.Text(21346,7626.9341,10,'60');
  RChartMapa.Text(21346,10495.4941,10,'75');
  actMapaZoomFit.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actLimparFiltrosRotaExecute(Sender: TObject);
begin
  FrmPrincipal.LimparColunasFiltro(DBGridRotaExecutantes,ColunasLayoutRotaExecutantes);
  actProcuraRotaExecutantes.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actLimparGraficoExecute(Sender: TObject);
begin
  RChartMapa.ClearGraf;
  RChartMapa.Repaint;
end;

procedure TFrmGerenciarEmbarcacoes.actListaRTEmbarqueExecute(Sender: TObject);
  var
    DataProcura: String;
begin
  RadioGroupImpressao.ItemIndex:= 1;
  TProgramacaoRTUtils.ListaRTEmbarque(RLImpressao, DateTimePickerProgramacao.DateTime,'TMIB;HOTEL');
  //==============================================
  edtTitulo1.Text:= 'Manifesto de Embarque no TMIB';
  edtTitulo2.Text:= 'DATA: '+ DataProcura;
  //==============================================
  StatusBarImpressao.Panels[0].Text:= 'N° registros: '+
    intToStr(RLImpressao.RowCount-1);
  actFinalGridImpressao.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actLocalizarRoteamentosExecute(
  Sender: TObject);
begin
  actProcuraRoteamento.Execute;
  actFiltroDestinos.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaMundiExecute(Sender: TObject);
  var
    i: integer;
    ListaXY: TStringList;
    X1,Y1,X2,Y2: Double;
    Endereco: String;
begin
  Endereco := ExtractFilePath(Application.ExeName) + 'Mapas\Mapa_1.txt';
  MemoMapaMundi.Lines.LoadFromFile(Endereco);
  ListaXY:=TStringList.Create;
  ListaXY.delimiter:= ';';
  ListaXY.StrictDelimiter:=False;
  RChartMapa.LineWidth:= 1;
  RChartMapa.PenStyle:= psSolid;
  RChartMapa.DataColor:= clNavy;
  for i := 1 to MemoMapaMundi.Lines.Count-1 do
  begin
    ListaXY.Clear;
    ListaXY.DelimitedText:= MemoMapaMundi.Lines[i];
    X1:= StrToFloat(ListaXY[0]);
    Y1:= StrToFloat(ListaXY[1]);
    X2:= StrToFloat(ListaXY[2]);
    Y2:= StrToFloat(ListaXY[3]);
    RChartMapa.MoveTo(X1,Y1);
    RChartMapa.DrawTo(X2,Y2);
    //RChartMapa.MarkAt(X,Y,24);
  end;
  actMapaZoomFit.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaPanExecute(Sender: TObject);
begin
  booleanZoomMapa:= false;
  RChartMapa.MouseAction := maPan;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaRoteiroExecute(Sender: TObject);
  var
    ListaRotaSequencia: TStringList;
    Roteiro: String;
    i: Integer;
begin
  if BooleanIndividual then
  begin
    ListaRotaSequencia:= TStringList.Create;
    ListaRotaSequencia.delimiter:= ';';
    ListaRotaSequencia.StrictDelimiter:=False;
    ListaRotaSequencia.DelimitedText:= FrmDataModule.DataSourceRoteamento.DataSet.
    FieldByName('RotaSequencia').AsString;
    Roteiro:= ListaRotaSequencia[0];
    for I := 1 to ListaRotaSequencia.Count-1 do
      Roteiro:= Roteiro+' X '+ListaRotaSequencia[i];
    StatusBarMapa.Panels[0].Text:= 'Roteiro: '+Roteiro;
    AutoFitStatusBar(StatusBarMapa);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaZoomDinamicoExecute(Sender: TObject);
begin
    booleanZoomMapa:= true;
    RChartMapa.MouseAction := maZoomDrag;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaZoomFitExecute(Sender: TObject);
  var
    rangeXLow,rangeXHigh,rangeYHigh,rangeYlow,rangeX,rangeY: Double;
    YLow,YHigh,XLow,XHigh,CentroX,CentroY,Largura,Altura: Double;
begin
  with RChartMapa do
  begin
    AutoRange(0,5);
    rangeXHigh:= Scale1X.RangeHigh;
    rangeXLow:= Scale1X.RangeLow;
    rangeYHigh:= Scale1Y.RangeHigh;
    rangeYlow:= Scale1Y.RangeLow;
    rangeX:= rangeXHigh-rangeXLow;
    rangeY:= rangeYHigh-rangeYLow;
    Largura:= (Width-LRim-RRim);
    Altura:= (Height -TRim-BRim);
    //====================================
    if rangeX > rangeY*(Largura/Altura) then
      rangeY:= rangeX/(Largura/Altura)
    else
      rangeX:= rangeY*(Largura/Altura);
    //=====================================
    CentroX:= ((rangeXHigh-rangeXLow)/2)+rangeXLow;
    CentroY:= ((rangeYHigh-rangeYLow)/2)+rangeYLow;
    XLow:= CentroX-(rangeX/2);
    XHigh:= CentroX+(rangeX/2);
    YLow:= CentroY-(rangeY/2);
    YHigh:= CentroY+(rangeY/2);
    setRange(0,XLow,YLow,XHigh,YHigh);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actMapaZoomWindExecute(Sender: TObject);
begin
  RChartMapa.MouseAction := maZoomWindPos;
end;

procedure TFrmGerenciarEmbarcacoes.actMarcarExecute(Sender: TObject);
  var
    Latitude,Longitude: Double;
    Plataforma: String;
    NoOrigem,NODestino: TPointFloat;
begin
  //Dados da plataforma de Origem (0,0)
  NoOrigem.X:= 0;
  NoOrigem.Y:= 0;
  //======================================================
  RChartMapa.DataColor:= clBlack;
  FrmDataModule.ADOQueryConsultaPlataforma.Active:= false;
  FrmDataModule.ADOQueryConsultaPlataforma.Active:= true;
  FrmDataModule.DataSourceConsultaPlataforma.Enabled:= false;
  FrmDataModule.ADOQueryConsultaPlataforma.First;
  while not(FrmDataModule.ADOQueryConsultaPlataforma.Eof) do
  begin
    Latitude:= FrmDataModule.DataSourceConsultaPlataforma.
    DataSet.FieldByName('Latitude').AsFloat;
    Longitude:= FrmDataModule.DataSourceConsultaPlataforma.
    DataSet.FieldByName('Longitude').AsFloat;
    Plataforma:= FrmDataModule.DataSourceConsultaPlataforma.
    DataSet.FieldByName('Plataforma').AsString;
    NODestino:= FrmPrincipal.LatLong_XY(Latitude,Longitude);
    RChartMapa.MarkAt(NODestino.X,NODestino.Y,24);
    RChartMapa.Text(NODestino.X,NODestino.Y,8,'  '+Plataforma);
    //Gravar coordenadas calculadas X,Y
    FrmDataModule.ADOQueryConsultaPlataforma.Edit;
    FrmDataModule.DataSourceConsultaPlataforma.
    DataSet.FieldByName('CoordX').AsFloat:= RoundTo(NODestino.X,-4);
    FrmDataModule.DataSourceConsultaPlataforma.
    DataSet.FieldByName('CoordY').AsFloat:= RoundTo(NODestino.Y,-4);
    FrmDataModule.ADOQueryConsultaPlataforma.Post;
    FrmDataModule.ADOQueryConsultaPlataforma.Next;
  end;
  FrmDataModule.ADOQueryConsultaPlataforma.First;
  FrmDataModule.DataSourceConsultaPlataforma.Enabled:= true;
  RChartMapa.Repaint;
end;

procedure TFrmGerenciarEmbarcacoes.actMostrarVinculadosExecute(Sender: TObject);
begin
  if btnVincluados.Down then
  begin
    PanelExecutantesVinculados.Visible:= true;
    PanelAnalise.Align:= alTop;
    PanelAnalise.Height:= 365;
    SplitterAnalise.Visible:= true;
    SplitterAnalise.Top:= PanelExecutantesVinculados.Top+5;
  end
  else
  begin
    PanelExecutantesVinculados.Visible:= false;
    SplitterAnalise.Visible:= false;
    PanelAnalise.Align:= alClient;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actNumerarAutomaticoExecute(Sender: TObject);
  var
    i,contRegistro,ACol: Integer;
begin
  contRegistro:= 1;
  ACol:= 7;
  for i := 0 to RLImpressao.ColCount-1 do
  begin
    if RLImpressao.Cells[i,0] = 'N°' then
    begin
      ACol:= i;
      break;
    end;
  end;
  for i := RLImpressao.Row to RLImpressao.RowCount-1 do
  begin
    RLImpressao.Cells[ACol,i]:= (IntToStr(contRegistro));
    contRegistro:= contRegistro + 1;
  end;

end;

procedure TFrmGerenciarEmbarcacoes.actPreencherRotaExecute(Sender: TObject);
  var
    numLinhas: Integer;
begin
  try
    FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount,
    'Gerando rota selecionada');
    FrmDataModule.ADOQueryTM_RotaExecutantes.First;
    while not FrmDataModule.ADOQueryTM_RotaExecutantes.Eof do
    begin
      numLinhas:= RLImpressao.RowCount;
      RLImpressao.RowCount:= numLinhas+1;
      //Preencher valores
      RLImpressao.Cells[0,numLinhas]:= IntToStr(numLinhas);
      RLImpressao.Cells[1,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('Origem').AsString);
      RLImpressao.Cells[2,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('txtDestino').AsString);
      RLImpressao.Cells[3,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('NomeExecutante').AsString);
      RLImpressao.Cells[4,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('txtTipoEtapaServico').AsString);
      RLImpressao.Cells[5,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('Empresa').AsString);
      RLImpressao.Cells[6,numLinhas]:= (FrmDataModule.
      DataSourceTM_RotaExecutantes.DataSet.FieldByName('NomeEmbarcacao').AsString);
      FrmDataModule.ADOQueryTM_RotaExecutantes.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;
    FrmPrincipal.ProgressBarAtualizar;
  except
    FrmPrincipal.ProgressBarAtualizar;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actProcuraGeralExecute(Sender: TObject);
  var
    SQLString,SQLBase,SQL_OrigemDestino,
    DataProcura: String;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  //==========================================================
  SQL_OrigemDestino:= '';
  if CheckBoxOrigemDestino.Checked then
    SQL_OrigemDestino:= 'AND (tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem)';
  //==========================================================
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutProcuraGeral,false);
  //Query de procura
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  //==========================================================
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON '+
  'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
  'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+')'+
  SQL_OrigemDestino+SQLString+' ORDER BY NomeExecutante,txtTipoEtapaServico;';
  //==========================================================
  FrmPrincipal.ProcuraQuery(SQLBase,
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao,StatusBarExecutantes);
  //Numero de selecionados
  actSelecionados.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actProcuraProgramadosExecute(Sender: TObject);
  var
    SQLString,SQLBase,SQL_OrigemDestino,
    DataProcura: String;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  //==========================================================
  SQL_OrigemDestino:= '';
  if CheckBoxOrigemDestino.Checked then
    SQL_OrigemDestino:= 'AND (tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante.Origem)';
  //==========================================================
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutProgramados,false);
  //Query de procura
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  //==========================================================
  SQLBase:=
  'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.*, tblRoteamento.* '+
  'FROM tblRoteamento INNER JOIN ((tblProgramacaoDiaria INNER JOIN '+
  'tblProgramacaoExecutante ON tblProgramacaoDiaria.idProgramacaoDiaria = '+
  'tblProgramacaoExecutante.CodigoProgramacaoDiaria) INNER JOIN tblAux_Rota_Distribuicao '+
  'ON tblProgramacaoExecutante.idProgramacaoExecutante = '+
  'tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ON tblRoteamento.idRoteamento = '+
  'tblAux_Rota_Distribuicao.CodigoRota '+
  'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(Dataprocura)+')'+
  SQLString+SQL_OrigemDestino+' ORDER BY NomeRota,NomeExecutante,txtTipoEtapaServico;';
  //==========================================================
  FrmPrincipal.ProcuraQuery(SQLBase,
  FrmDataModule.ADOQueryProgramados,StatusBarExecutantes);
end;

procedure TFrmGerenciarEmbarcacoes.actProcuraRotaExecutantesExecute(
  Sender: TObject);
 var
    ListaRotaSequencia: TStringList;
    i: Integer;
    Plataforma,Latitude,Longitude,
    SQLString,SQLBase,idRoteamento: String;
begin
  if DBEditCodigoRoteamento.Text <> '' then
  begin
    try
      idRoteamento:= FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('idRoteamento').AsString;
      SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutRotaExecutantes,false);
      if SQLString<>'' then
        SQLString:= 'AND'+SQLString;
      SQLBase:= 'SELECT tblRoteamento.*, tblAux_Rota_Distribuicao.*, tblProgramacaoExecutante.*, '+
      'tblProgramacaoDiaria.* FROM tblProgramacaoDiaria INNER JOIN (tblRoteamento INNER JOIN '+
      '(tblProgramacaoExecutante INNER JOIN tblAux_Rota_Distribuicao ON tblProgramacaoExecutante.'+
      'idProgramacaoExecutante = tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ON '+
      'tblRoteamento.idRoteamento = tblAux_Rota_Distribuicao.CodigoRota) ON tblProgramacaoDiaria.'+
      'idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgramacaoDiaria '+
      'WHERE ((tblAux_Rota_Distribuicao.CodigoRota = '+idRoteamento+'))'+SQLString+
      ' ORDER BY NomeExecutante;';
      FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryTM_RotaExecutantes,StatusBarRotaExecutantes);
      actSelecionadosRotaExecutantes.Execute;
      PanelTituloRotaExecutantes.Caption:= FrmDataModule.DataSourceRoteamento.DataSet.
      FieldByName('NomeRota').AsString+' - '+FrmDataModule.DataSourceRoteamento.DataSet.
      FieldByName('NomeEmbarcacao').AsString+' - '+FrmDataModule.DataSourceRoteamento.DataSet.
      FieldByName('HoraRoteamento').AsString;
    except
    end;
    ListaRotaSequencia:= TStringList.Create;
    ListaRotaSequencia.delimiter:= ';';
    ListaRotaSequencia.StrictDelimiter:=true;
    ListaRotaSequencia.DelimitedText:= FrmDataModule.DataSourceRoteamento.DataSet.
    FieldByName('RotaSequencia').AsString;
    StrGridRota.RowCount:= ListaRotaSequencia.Count+1;
    StrGridRota.Cells[1,1]:= '';
    //Preencher sequencia da rota com Latitude e Longitude
    for I := 0 to ListaRotaSequencia.Count-1 do
    begin
      StrGridRota.Cells[1,i+1]:= '';
      StrGridRota.Cells[2,i+1]:= '';
      StrGridRota.Cells[3,i+1]:= '';
      StrGridRota.Cells[4,i+1]:= '';
      StrGridRota.Cells[5,i+1]:= '';
      StrGridRota.Cells[6,i+1]:= '';
      //===========================
      Plataforma:= ListaRotaSequencia[i];
      FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := false;
      FrmDataModule.ADOQueryConsultaPlataforma_Nome.Parameters.Items[0].Value:= Plataforma;
      FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := true;
      Latitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
      FieldByName('Latitude').AsString;
      Longitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
      FieldByName('Longitude').AsString;
      StrGridRota.Cells[1,i+1]:= Plataforma;
      StrGridRota.Cells[5,i+1]:= Latitude;
      StrGridRota.Cells[6,i+1]:= Longitude;
    end;
    try
      StrGridRota.FixedRows:= 1;
    except
    end;
  end
  else
  begin
    try
      FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= false;
      FrmDataModule.ADOQueryTM_RotaExecutantes.Parameters.Items[0].Value:= 0;
      FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= true;
      StatusBarRotaExecutantes.Panels[0].Text:= 'N° registros: 0';
      AutoFitStatusBar(StatusBarRotaExecutantes);
      StrGridRota.RowCount:= 1;
    except
    end;
  end;
  actDisponivel.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actProcuraRoteamentoExecute(Sender: TObject);
    var
    SQLString,SQLBase,DataProcura: String;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayoutRoteamento,false);
  if SQLString<>'' then
    SQLString:= 'AND'+SQLString;
  SQLBase:=  'SELECT tblRoteamento.* FROM tblRoteamento '+
  'WHERE (DataRoteamento LIKE '+QuotedStr(Dataprocura)+') '+SQLString+
  'ORDER BY NomeRota,HoraRoteamento;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryRoteamento,StatusBarRoteamento);
end;

procedure TFrmGerenciarEmbarcacoes.actRemoveColExecute(Sender: TObject);
begin
  FrmPrincipal.DeleteCol(RLImpressao,RLImpressao.Col);
end;

procedure TFrmGerenciarEmbarcacoes.actRemoveRowExecute(Sender: TObject);
begin
  FrmPrincipal.DeleteRow(RLImpressao,RLImpressao.Row);
end;

procedure TFrmGerenciarEmbarcacoes.actTransbordoInternoExecute(Sender: TObject);
begin
  if Assigned(FrmTabela.RLTabela) then
  begin
    FrmTabela.PanelTitulo.Caption:= 'Transbordo de Executantes';
    FrmTabela.Caption:= 'Transbordo de Executantes';
    TProgramacaoRTUtils.ProgramacaoTransbordosComHora(FrmTabela.RLTabela,
        FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date), false);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actRLImprimir_RotaExecute(Sender: TObject);
  var
    DataProcura,idRoteamento,NomeRota,
    Roteiro,NomeEmbarcacao,DataRoteamento,HoraRoteamento: String;
    ListaRotaSequencia: TStringList;
    i: Integer;
begin
  RadioGroupImpressao.ItemIndex:= 0;
  DataProcura:= FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date);
  idRoteamento:= IntToStr(FrmDataModule.DataSourceRoteamento.
  DataSet.FieldByName('idRoteamento').AsInteger);
  NomeEmbarcacao:= FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('NomeEmbarcacao').AsString;
  NomeRota:= FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('NomeRota').AsString;
  DataRoteamento:= FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('DataRoteamento').AsString;
  HoraRoteamento:= FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('HoraRoteamento').AsString;
  ListaRotaSequencia:= TStringList.Create;
  ListaRotaSequencia.delimiter:= ';';
  ListaRotaSequencia.StrictDelimiter:=False;
  ListaRotaSequencia.DelimitedText:= FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('RotaSequencia').AsString;
  try
    Roteiro:= ListaRotaSequencia[0];
  except

  end;
  for I := 1 to ListaRotaSequencia.Count-1 do
    Roteiro:= Roteiro+' X '+ListaRotaSequencia[i];

  edtTitulo1.Text:= ''+NomeRota+ ': '+NomeEmbarcacao+'|ORIGEM '+
  StrGridRota.Cells[1,1]+' ÀS '+
  HoraRoteamento+ ' ('+DataProcura+')';
  edtTitulo2.Text:= 'ROTEIRO: '+ Roteiro;
  //=============================================================
  RLImpressao.FixedRows:= 0;
  RLImpressao.RowCount:= 1;
  RLImpressao.ColCount:= 8;
  RLImpressao.Cells[0,0]:= 'N°';
  RLImpressao.Cells[1,0]:= 'Origem';
  RLImpressao.Cells[2,0]:= 'Destino';
  RLImpressao.Cells[3,0]:= 'Nome do Executante';
  RLImpressao.Cells[4,0]:= 'Tipo de Etapa de Serviço';
  RLImpressao.Cells[5,0]:= 'Empresa';
  RLImpressao.Cells[6,0]:= 'Lancha';
  RLImpressao.Cells[7,0]:= 'Assinatura';
  //=============================================================
  actPreencherRota.Execute;
  StatusBarImpressao.Panels[0].Text:= 'N° registros: '+
  intToStr(RLImpressao.RowCount-1);
  actFinalGridImpressao.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actSalvarDesenhoExecute(Sender: TObject);
begin
  SaveDialog1.FileName:= 'Desenho.zuchi';
  SaveDialog1.DefaultExt:= 'zuchi';
  SaveDialog1.Filter:= 'Desenho (*.zuchi)|*.zuchi|Todos os Arquivos (*.*)|*.*';
  if SaveDialog1.Execute then
    RChartMapa.SaveData(SaveDialog1.FileName);
end;

procedure TFrmGerenciarEmbarcacoes.actSalvarImagemExecute(Sender: TObject);
begin
  SavePictureDialog1.FileName:= 'Mapa.wmf';
  if SavePictureDialog1.Execute then
    RChartMapa.CopyToWMF(SavePictureDialog1.FileName,false);
end;

procedure TFrmGerenciarEmbarcacoes.actSalvarRotaExecute(Sender: TObject);
begin
  FrmDataModule.DataSourceRoteamento.DataSet.Edit;
  FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('RotaSequencia').AsString :=
  MontarRotaSequenciaDoGrid;
  FrmDataModule.DataSourceRoteamento.DataSet.Post;
end;

procedure TFrmGerenciarEmbarcacoes.actSelecionadosExecute(Sender: TObject);
  var
    NumSelecionados: Integer;
begin
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  //FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL:=
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.SQL;
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  NumSelecionados:= 0;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    if FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('booleanSelecao').AsBoolean then
      NumSelecionados:= NumSelecionados+1;
      FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  end;
  StatusBarExecutantes.Panels[1].Text:= 'Selecionados: '+IntToStr(NumSelecionados);
  AutoFitStatusBar(StatusBarExecutantes);
end;

procedure TFrmGerenciarEmbarcacoes.actSelecionadosRotaExecutantesExecute(
  Sender: TObject);
  var
    NumSelecionados: Integer;
begin
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  //FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL:=
  FrmDataModule.ADOQueryTM_RotaExecutantes.SQL;
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  NumSelecionados:= 0;
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    if FrmDataModule.DataSourceTemporarioDBColibri.DataSet.FieldByName('booleanSelecao').AsBoolean then
      NumSelecionados:= NumSelecionados+1;
      FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  end;
  StatusBarRotaExecutantes.Panels[1].Text:= 'Selecionados: '+IntToStr(NumSelecionados);
  AutoFitStatusBar(StatusBarRotaExecutantes);
end;

procedure TFrmGerenciarEmbarcacoes.actSelProgExec1Execute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecNo;
  FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.First;
  while not FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Next;
  end;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= true;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecNo:= selRegistro;
  FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled:= true;
  //Numero de selecionados
  actSelecionados.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actSelProgExec2Execute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecNo;
  FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled:= false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.First;
  while not FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Next;
  end;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= false;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active:= true;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecNo:= selRegistro;
  FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled:= true;
  //Numero de selecionados
  actSelecionados.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actSelRotaExecLIMPARExecute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryTM_RotaExecutantes.RecNo;
  FrmDataModule.DataSourceTM_RotaExecutantes.Enabled:= false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.First;
  while not FrmDataModule.ADOQueryTM_RotaExecutantes.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceTM_RotaExecutantes.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryTM_RotaExecutantes.Next;
  end;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= true;
  FrmDataModule.ADOQueryTM_RotaExecutantes.RecNo:= selRegistro;
  FrmDataModule.DataSourceTM_RotaExecutantes.Enabled:= true;
  //Numero de selecionados
  actSelecionadosRotaExecutantes.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actSelRotaExecTODOSExecute(Sender: TObject);
  var
    selRegistro,idProgramacaoExecutante: Integer;
begin
  selRegistro:= FrmDataModule.ADOQueryTM_RotaExecutantes.RecNo;
  FrmDataModule.DataSourceTM_RotaExecutantes.Enabled:= false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.First;
  while not FrmDataModule.ADOQueryTM_RotaExecutantes.Eof do
  begin
    idProgramacaoExecutante:= FrmDataModule.DataSourceTM_RotaExecutantes.
    DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
    //Alterar seleção na raiz
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value:=
    idProgramacaoExecutante;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
    FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.
    FieldByName('booleanSelecao').AsBoolean:= true;
    FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

    //=============================================
    FrmDataModule.ADOQueryTM_RotaExecutantes.Next;
  end;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= false;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active:= true;
  FrmDataModule.ADOQueryTM_RotaExecutantes.RecNo:= selRegistro;
  FrmDataModule.DataSourceTM_RotaExecutantes.Enabled:= true;
  //Numero de selecionados
  actSelecionadosRotaExecutantes.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actSugerirRotaExecute(Sender: TObject);
var
  IdExecutante: Integer;
  DataProg: string;
  DistLogistica: TDistribuicaoLogistica;
  RotasAvaliadas: TArray<TRotaAvaliada>;
  I, MaxSugestoes: Integer;
  MsgSugestao: string;
  OpcaoEscolhida: Integer;
  MsgErro: string;
begin
  if FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.IsEmpty then Exit;

  IdExecutante := FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
  DataProg := DateToStr(DateTimePickerProgramacao.Date);

  DistLogistica := TDistribuicaoLogistica.Create(FrmDataModule.ADOConnectionColibri, FrmDataModule.ADOConnectionConsulta);
  try
    FrmPrincipal.ProgressBarIncializa(100, 'Analisando melhores rotas (IA)...');

    RotasAvaliadas := DistLogistica.AvaliarRotasParaExecutante(IdExecutante, DataProg);
    FrmPrincipal.ProgressBarAtualizar;

    if Length(RotasAvaliadas) = 0 then
    begin
      ShowMessage('Nenhuma rota encontrada para esta data.');
      Exit;
    end;

    // Montar mensagem com Top 3
    MsgSugestao := 'Melhores rotas sugeridas pela IA:' + #13#10#13#10;
    MaxSugestoes := Min(3, Length(RotasAvaliadas));

    for I := 0 to MaxSugestoes - 1 do
    begin
      if RotasAvaliadas[I].Valida then
      begin
        MsgSugestao := MsgSugestao + Format('[%d] %s (%s) - Partida: %s' + #13#10 +
                                            '    Score: %.1f | Ocupação: %.1f%% | Viagem: %d min' + #13#10#13#10,
          [I + 1,
           RotasAvaliadas[I].NomeRota,
           RotasAvaliadas[I].NomeEmbarcacao,
           TimeToStr(RotasAvaliadas[I].HoraPartida),
           RotasAvaliadas[I].ScoreTotal,
           RotasAvaliadas[I].OcupacaoPercentual,
           RotasAvaliadas[I].TempoViagemMinutos]);
      end
      else
      begin
        MsgSugestao := MsgSugestao + Format('[X] %s - INVÁLIDA (%s)' + #13#10#13#10,
          [RotasAvaliadas[I].NomeRota, RotasAvaliadas[I].MotivoRejeicao]);
      end;
    end;

    if not RotasAvaliadas[0].Valida then
    begin
      ShowMessage('Nenhuma rota válida disponível para este executante.' + #13#10#13#10 + MsgSugestao);
      Exit;
    end;

    MsgSugestao := MsgSugestao + 'Digite o número da rota desejada (1 a ' + IntToStr(MaxSugestoes) + ') ou 0 para cancelar:';

    OpcaoEscolhida := StrToIntDef(InputBox('Sugestão de Rota', MsgSugestao, '1'), 0);

    if (OpcaoEscolhida > 0) and (OpcaoEscolhida <= MaxSugestoes) and RotasAvaliadas[OpcaoEscolhida - 1].Valida then
    begin
      if DistLogistica.VincularExecutanteARota(IdExecutante, RotasAvaliadas[OpcaoEscolhida - 1].IdRoteamento, MsgErro) then
      begin
        // Atualizar status
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := false;
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value := IdExecutante;
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := true;
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
        FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet.FieldByName('InseridoProgramacaoTransporte').AsBoolean := true;
        FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

        ShowMessage('Executante vinculado com sucesso à rota ' + RotasAvaliadas[OpcaoEscolhida - 1].NomeRota);

        // Atualizar grids
        FrmDataModule.ADOQueryTM_RotaExecutantes.Active := false;
        FrmDataModule.ADOQueryTM_RotaExecutantes.Active := true;
        FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := false;
        FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := true;
      end
      else
        ShowMessage('Erro ao vincular: ' + MsgErro);
    end;

  finally
    DistLogistica.Free;
  end;
end;

function TFrmGerenciarEmbarcacoes.MontarRotaSequenciaDoGrid: string;
var
  I: Integer;
  OrigemDaRota, Destino: string;
  Itens: TStringList;
begin
  Result := '';
  Itens := TStringList.Create;
  try
    OrigemDaRota := Trim(StrGridRota.Cells[1, 1]);

    if OrigemDaRota <> '' then
      Itens.Add(OrigemDaRota);

    for I := 2 to StrGridRota.RowCount - 1 do
    begin
      Destino := Trim(StrGridRota.Cells[1, I]);

      if Destino <> '' then
      begin
        if Itens.IndexOf(Destino) < 0 then
          Itens.Add(Destino);
      end;
    end;

    Result := StringReplace(Itens.Text, sLineBreak, ';', [rfReplaceAll]);

    if Result.EndsWith(';') then
      Delete(Result, Length(Result), 1);
  finally
    Itens.Free;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actVincularExecute(Sender: TObject);
var
  CodigoProgramacaoExecutante, CodigoRota, NumRegistros: Integer;
  DistLogistica: TDistribuicaoLogistica;
  ValidacaoCapacidade: TValidacaoResult;
  NumVinculados, NumFalhas: Integer;
  MensagemErro: string;
  MensagemErroDetalhada: string;
begin
  CodigoRota := FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('idRoteamento').AsInteger;

  if Trim(FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('RotaSequencia').AsString) = '' then
  begin
    ShowMessage('A rota selecionada não possui sequência definida. Salve ou recalcule o roteiro antes de vincular executantes.');
    Exit;
  end;

  DistLogistica := TDistribuicaoLogistica.Create(
    FrmDataModule.ADOConnectionColibri,
    FrmDataModule.ADOConnectionConsulta
  );
  try
    MensagemErroDetalhada := '';

    NumRegistros := 0;
    FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.First;
    while not FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Eof do
    begin
      if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet
           .FieldByName('booleanSelecao').AsBoolean then
        Inc(NumRegistros);

      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Next;
    end;

    if NumRegistros = 0 then
    begin
      ShowMessage('Nenhum executante selecionado para vinculação.');
      Exit;
    end;

    ValidacaoCapacidade := DistLogistica.ValidarCapacidadeEmbarcacao(CodigoRota, NumRegistros);
    if not ValidacaoCapacidade.Valido then
    begin
      if Application.MessageBox(
           PChar(ValidacaoCapacidade.Mensagem + #13#10#13#10 + 'Deseja continuar mesmo assim?'),
           'Aviso de Capacidade',
           MB_YESNO + MB_ICONWARNING
         ) = IDNO then
        Exit;
    end;

    FrmDataModule.ADOQueryAux_Rota_Distribuicao.Active := True;
    FrmPrincipal.ProgressBarIncializa(
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecordCount,
      'Vinculando executantes à rota selecionada...'
    );

    FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled := False;
    FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.First;

    NumVinculados := 0;
    NumFalhas := 0;

    while not FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Eof do
    begin
      if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet
           .FieldByName('booleanSelecao').AsBoolean then
      begin
        CodigoProgramacaoExecutante :=
          FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet
            .FieldByName('idProgramacaoExecutante').AsInteger;

        if DistLogistica.VincularExecutanteARota(
             CodigoProgramacaoExecutante,
             CodigoRota,
             MensagemErro
           ) then
        begin
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := False;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Parameters.Items[0].Value :=
            CodigoProgramacaoExecutante;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Active := True;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Edit;
          FrmDataModule.DataSourceConsultaProgramacaoExecutante_ID.DataSet
            .FieldByName('InseridoProgramacaoTransporte').AsBoolean := True;
          FrmDataModule.ADOQueryConsultaProgramacaoExecutante_ID.Post;

          Inc(NumVinculados);
        end
        else
        begin
          Inc(NumFalhas);
          MensagemErroDetalhada := MensagemErroDetalhada +
            Format('Executante %d: %s', [CodigoProgramacaoExecutante, MensagemErro]) +
            sLineBreak;
        end;
      end;

      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Next;
      FrmPrincipal.ProgressBarIncremento(1);
    end;

  finally
    DistLogistica.Free;
    FrmPrincipal.ProgressBarAtualizar;
  end;

  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.First;
  FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.Enabled := True;

  FrmDataModule.ADOQueryTM_RotaExecutantes.Active := False;
  FrmDataModule.ADOQueryTM_RotaExecutantes.Active := True;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := False;
  FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Active := True;

  NumRegistros := FrmDataModule.ADOQueryTM_RotaExecutantes.RecordCount;
  StatusBarRotaExecutantes.Panels[0].Text := 'Nº registros: ' + IntToStr(NumRegistros);
  AutoFitStatusBar(StatusBarRotaExecutantes);

  NumRegistros := FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.RecordCount;
  StatusBarExecutantes.Panels[0].Text := 'Nº registros: ' + IntToStr(NumRegistros);

  if NumFalhas > 0 then
    ShowMessage(
      Format('%d executantes vinculados com sucesso. %d falharam.' + sLineBreak + sLineBreak + '%s',
        [NumVinculados, NumFalhas, MensagemErroDetalhada])
    );
  //else if NumVinculados > 0 then
   // ShowMessage(Format('%d executantes vinculados com sucesso.', [NumVinculados]));
end;

procedure TFrmGerenciarEmbarcacoes.Ajustartamanhodascolunas1Click(
  Sender: TObject);
begin
  AutoFitGrade(RLImpressao);
end;

procedure TFrmGerenciarEmbarcacoes.AtualizarProgressoCriacaoRotas(
  ATotal, AAtual: Integer; const AMensagem: string);
var
  Agora: Cardinal;
begin
  if ATotal <= 0 then
    ATotal := 1;

  FrmPrincipal.ProgressBarPrincipal.Max := ATotal;
  FrmPrincipal.ProgressBarPrincipal.value := AAtual;

  // se você tiver label/status text:
  FrmPrincipal.StatusBar1.Panels[0].Text := AMensagem;

  // evita repintar em excesso
  Agora := GetTickCount;
  if (Agora - FUltimoTickProgresso >= 80) or (AAtual >= ATotal) then
  begin
    FrmPrincipal.ProgressBarPrincipal.Update;
    FrmPrincipal.StatusBar1.Update;
    Application.ProcessMessages;
    FUltimoTickProgresso := Agora;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.btnProgramadosClick(Sender: TObject);
begin
  if btnProgramados.Down then
  begin
    DBGridProcuraProgramados.Visible:= true;
    DBGridProcuraGeral.Visible:= false;
    DBGridProcuraProgramados.Align:= alClient;
    actVincular.Enabled:= false;
    actDisponivel.Enabled := false;
    actProcuraProgramados.Execute;
    PanelTituloExecutantes.Caption:= 'Consulta de Executantes Programados';
    btnFiltroClearProcuraGeral.Visible:= false;
    btnFiltroClearProgramados.Visible:= true;
  end
  else
  begin
    DBGridProcuraProgramados.Visible:= false;
    DBGridProcuraGeral.Visible:= true;
    DBGridProcuraGeral.Align:= alClient;
    actVincular.Enabled:= true;
    actDisponivel.Enabled := true;
    actProcuraGeral.Execute;
    PanelTituloExecutantes.Caption:= 'Seleção de Executantes para vinculo ao roteamento';
    btnFiltroClearProcuraGeral.Visible:= true;
    btnFiltroClearProgramados.Visible:= false;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.SpeedButton4Click(Sender: TObject);
begin
  PanelAjuda.Visible := false;
end;

procedure TFrmGerenciarEmbarcacoes.SpinEditFonteChange(Sender: TObject);
begin
  RLImpressao.Font.Size:= SpinEditFonte.Value;
end;

procedure TFrmGerenciarEmbarcacoes.StrGridRotaRowMoved(Sender: TObject;
  FromIndex, ToIndex: Integer);
begin
  actSalvarRota.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.StrGridRotaSelectCell(Sender: TObject; ACol,
  ARow: Integer; var CanSelect: Boolean);
  var
    R: TRect;
begin
  if ((ACol = 1) AND(ARow > 1)) then
  begin
    R := StrGridRota.CellRect(ACol, ARow);
    R.Left := R.Left + StrGridRota.Left;
    R.Right := R.Right + StrGridRota.Left;
    R.Top := R.Top + StrGridRota.Top;
    R.Bottom := R.Bottom + StrGridRota.Top;
    ComboBoxPlataforma.Left := R.Left + 1;
    ComboBoxPlataforma.Top := R.Top + 1;
    ComboBoxPlataforma.Width := (R.Right + 1) - R.Left;
    ComboBoxPlataforma.Height := (R.Bottom + 1) - R.Top;
    ComboBoxPlataforma.Visible := True;
    FrmPrincipal.selComboBox(ComboBoxPlataforma,
    (StrGridRota.Cells[1,ARow]));
    ComboBoxPlataforma.SetFocus;
  end
  else if ((ACol = 1) AND(ARow = 1)) then
  begin
    R := StrGridRota.CellRect(ACol, ARow);
    R.Left := R.Left + StrGridRota.Left;
    R.Right := R.Right + StrGridRota.Left;
    R.Top := R.Top + StrGridRota.Top;
    R.Bottom := R.Bottom + StrGridRota.Top;
    ComboBoxOrigemInicial.Left := R.Left + 1;
    ComboBoxOrigemInicial.Top := R.Top + 1;
    ComboBoxOrigemInicial.Width := (R.Right + 1) - R.Left;
    ComboBoxOrigemInicial.Height := (R.Bottom + 1) - R.Top;
    ComboBoxOrigemInicial.Visible := True;
    FrmPrincipal.selComboBox(ComboBoxOrigemInicial,
    (StrGridRota.Cells[1,ARow]));
    ComboBoxOrigemInicial.SetFocus;
  end;
  CanSelect := True;
end;

procedure TFrmGerenciarEmbarcacoes.ToolButton4Click(Sender: TObject);
begin
  BooleanIndividual:= true;
  actGerarLinhas.Execute;
end;

function TFrmGerenciarEmbarcacoes.verificaAPLAT(NomeExecutante,Funcao: String): Boolean;
  var
    i: Integer;
begin
  Result:= false;
  for i := 0 to High(FrmPrincipal.MatrizExecutanteAPLAT[0]) do
    if ((FrmPrincipal.MatrizExecutanteAPLAT[3,i] = NomeExecutante)AND
    (FrmPrincipal.MatrizExecutanteAPLAT[1,i] = Funcao)) then
    begin
      Result:= true;
      break
    end;
end;

procedure TFrmGerenciarEmbarcacoes.StrGridResumoDblClick(Sender: TObject);
  var
    numLinha: Integer;
    Plataforma,Latitude,Longitude: string;
begin
  numLinha:= StrGridRota.RowCount;
  StrGridRota.RowCount:= numLinha+1;
  Plataforma:= StrGridResumo.Cells[0,StrGridResumo.Row];
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := false;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Parameters.Items[0].Value:= Plataforma;
  FrmDataModule.ADOQueryConsultaPlataforma_Nome.Active := true;
  Latitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
  FieldByName('Latitude').AsString;
  Longitude:= FrmDataModule.DataSourceConsultaPlataforma_Nome.DataSet.
  FieldByName('Longitude').AsString;
  StrGridRota.Cells[1,numLinha]:= Plataforma;
  StrGridRota.Cells[5,numLinha]:= Latitude;
  StrGridRota.Cells[6,numLinha]:= Longitude;
  actSalvarRota.Execute;
  AutoFitGrade(StrGridRota);
  StrGridRota.ColWidths[0]:= 20;
  StrGridRota.ColWidths[1]:= 80;
end;

procedure TFrmGerenciarEmbarcacoes.MaskEditHoraChange(Sender: TObject);
begin
  try
    if not (DBGridRoteamento.DataSource.State in [dsEdit, dsInsert]) then
      DBGridRoteamento.DataSource.Edit;

    DBGridRoteamento.DataSource.DataSet.FieldByName('HoraRoteamento').AsString:=
    MaskEditHora.Text;
  except
  end;
end;

procedure TFrmGerenciarEmbarcacoes.actFiltroDestinosExecute(Sender: TObject);
 var
    DataProcura,Destino1,Destino2,SQLBase,SQLOrigem,TotalOrigem: String;
    ListaOrigem: TStringList;
    i,Linha: Integer;
begin
  DataProcura:= DateToStr(DateTimePickerProgramacao.Date);
  ListaOrigem:= TStringList.Create;
  ListaOrigem.delimiter:= ';';
  ListaOrigem.StrictDelimiter:= true;
  ListaOrigem.DelimitedText:= FrmPrincipal.OrigemPlataformas;
  if btnPendentes.Down then
  begin
    SQLOrigem:= FrmPrincipal.palavraBusca('','tblProgramacaoExecutante.Origem','Exato',
    FrmPrincipal.OrigemPlataformas,'AND');
    if SQLOrigem <> '' then
      SQLOrigem:='AND'+SQLOrigem;
    SQLBase:=
    'SELECT tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao, '+
    'tblProgramacaoExecutante.StatusProgramacao, tblProgramacaoExecutante.'+
    'InseridoProgramacaoTransporte, tblProgramacaoExecutante.Origem '+
    'FROM tblProgramacaoDiaria INNER JOIN '+
    'tblProgramacaoExecutante ON tblProgramacaoDiaria.idProgramacaoDiaria = '+
    'tblProgramacaoExecutante.CodigoProgramacaoDiaria GROUP BY '+
    'tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao, '+
    'tblProgramacaoExecutante.StatusProgramacao, tblProgramacaoExecutante.'+
    'InseridoProgramacaoTransporte, tblProgramacaoExecutante.Origem '+
    'HAVING (((tblProgramacaoDiaria.DataProgramacao Like '+QuotedStr(DataProcura)+
    ')AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado")'+SQLOrigem+
    'AND(tblProgramacaoExecutante.InseridoProgramacaoTransporte LIKE False)));';
  end
  else
  begin
    SQLOrigem:= FrmPrincipal.palavraBusca('','tblProgramacaoExecutante.Origem','Exato',
    FrmPrincipal.OrigemPlataformas,'AND');
    if SQLOrigem <> '' then
      SQLOrigem:='AND'+SQLOrigem;
    SQLBase:=
    'SELECT tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao, '+
    'tblProgramacaoExecutante.StatusProgramacao, tblProgramacaoExecutante.Origem '+
    'FROM tblProgramacaoDiaria INNER JOIN '+
    'tblProgramacaoExecutante ON tblProgramacaoDiaria.idProgramacaoDiaria = '+
    'tblProgramacaoExecutante.CodigoProgramacaoDiaria GROUP BY '+
    'tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.DataProgramacao, '+
    'tblProgramacaoExecutante.StatusProgramacao, tblProgramacaoExecutante.Origem '+
    'HAVING (((tblProgramacaoDiaria.DataProgramacao LIKE '+QuotedStr(DataProcura)+
    ')AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado")'+SQLOrigem+'));';
  end;
  FrmDataModule.ADOQueryTemporarioDBColibri.Close;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Clear;
  FrmDataModule.ADOQueryTemporarioDBColibri.SQL.Add(SQLBase);
  FrmDataModule.ADOQueryTemporarioDBColibri.Open;
  StrGridResumo.FixedRows:= 0;
  StrGridResumo.RowCount:= 1;
  Linha:= StrGridResumo.RowCount;
  StrGridResumo.ColCount:= ListaOrigem.Count+1;
  StrGridResumo.Cells[0,0]:= 'Plataforma';
  for I := 0 to ListaOrigem.Count-1 do
    StrGridResumo.Cells[i+1,0]:= ListaOrigem[i];
  //=======================================
  FrmDataModule.ADOQueryTemporarioDBColibri.First;
  Destino1:= 'aaaa';
  while not FrmDataModule.ADOQueryTemporarioDBColibri.Eof do
  begin
    Destino2:= FrmDataModule.DataSourceTemporarioDBColibri.
    DataSet.FieldByName('txtDestino').AsString;
    if Destino1<>Destino2 then
    begin
      for I := 0 to ListaOrigem.Count-1 do
      begin
        if btnTotal.Down = false then
        begin
          FrmDataModule.ADOQueryPassageirosTM_SIM.Active := false;
          FrmDataModule.ADOQueryPassageirosTM_SIM.Parameters.Items[0].Value:= Dataprocura;
          FrmDataModule.ADOQueryPassageirosTM_SIM.Parameters.Items[1].Value:= ListaOrigem[i];
          FrmDataModule.ADOQueryPassageirosTM_SIM.Parameters.Items[2].Value:= Destino2;
          FrmDataModule.ADOQueryPassageirosTM_SIM.Active := true;
          TotalOrigem:= IntToStr(FrmDataModule.ADOQueryPassageirosTM_SIM.RecordCount);
        end
        else
        begin
          FrmDataModule.ADOQueryPassageirosTM_NAO.Active := false;
          FrmDataModule.ADOQueryPassageirosTM_NAO.Parameters.Items[0].Value:= Dataprocura;
          FrmDataModule.ADOQueryPassageirosTM_NAO.Parameters.Items[1].Value:= ListaOrigem[i];
          FrmDataModule.ADOQueryPassageirosTM_NAO.Parameters.Items[2].Value:= Destino2;
          FrmDataModule.ADOQueryPassageirosTM_NAO.Active := true;
          TotalOrigem:= IntToStr(FrmDataModule.ADOQueryPassageirosTM_NAO.RecordCount);
        end;
        StrGridResumo.Cells[0,Linha]:= Destino2;
        StrGridResumo.Cells[i+1,Linha]:= TotalOrigem;
      end;
      StrGridResumo.RowCount:= StrGridResumo.RowCount+1;
      Linha:= StrGridResumo.RowCount;
      Destino1:= Destino2;
    end;
    FrmDataModule.ADOQueryTemporarioDBColibri.Next;
  end;
  try
    StrGridResumo.FixedRows:= 1;
  except
  end;
  AutoFitGrade(StrGridResumo);
end;

procedure TFrmGerenciarEmbarcacoes.actFinalGridImpressaoExecute(
  Sender: TObject);
begin
  AutoFitStatusBar(StatusBarImpressao);
  actFixarLinhaColuna.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actFixarLinhaColunaExecute(Sender: TObject);
begin
  if Fixar1Coluna1.Checked then
  begin
    try
      RLImpressao.FixedCols:= 1;
    except

    end;
  end
  else
    RLImpressao.FixedCols:= 0;

  if Fixar1Linha1.Checked then
  begin
    try
      RLImpressao.FixedRows:= 1;
    except

    end;
  end
  else
    RLImpressao.FixedRows:= 0;
end;

procedure TFrmGerenciarEmbarcacoes.actGerarLinhasExecute(Sender: TObject);
const
  VelocidadeKmH = 40.0; // lancha/embarcação
  MinTravelMin = 15;
  DefaultNoDistMin = 60;
var
  NoOrigem, NoDestino: TPointFloat;
  POrigem, PDestino: string;
  i, TravelMin: Integer;
  DistKm: Double;
  PlatOrigem, PlatDestino: TDadosPlataforma;

  procedure SetupChart;
  begin
    RChartMapa.TextBkStyle   := tbClear;
    RChartMapa.TextAlignment := taLeftJustify;
    RChartMapa.LineWidth     := 1;
    RChartMapa.PenStyle      := psSolid;

    // evita AV se dataset estiver fechado
    if (FrmDataModule.DataSourceRoteamento <> nil) and
       (FrmDataModule.DataSourceRoteamento.DataSet <> nil) and
       (FrmDataModule.DataSourceRoteamento.DataSet.Active) then
      RChartMapa.DataColor := FrmDataModule.DataSourceRoteamento.DataSet.FieldByName('linhaCor').AsInteger;

    RChartMapa.FillColor := clWhite;
  end;

  procedure MarkPoint(const Pt: TPointFloat; const LabelText: string);
  begin
    RChartMapa.MarkAt(Pt.X, Pt.Y, 9);
    if LabelText <> '' then
      RChartMapa.Text(Pt.X, Pt.Y, 8, '  ' + LabelText);
  end;

  function SafeGetPlat(const Nome: string; out P: TDadosPlataforma): Boolean;
  begin
    Result := False;
    if Trim(Nome) = '' then Exit;

    P := TProgramacaoRTUtils.DadosPlataforma_RT(Nome);

    // proteção básica: coordenadas do mapa precisam existir
    if (Abs(P.CoordX) < 1e-8) and (Abs(P.CoordY) < 1e-8) then
      Exit(False);

    Result := True;
  end;

  function FormatDistMin(const Dist: Double; const Min: Integer): string;
  begin
    // exemplo: "12.34 km | 19 min"
    Result := FormatFloat('0.00', Dist) + ' km | ' + IntToStr(Min) + ' min';
  end;

begin
  if StrGridRota.RowCount < 2 then
    Exit;

  // Origem (linha 1)
  POrigem := Trim(StrGridRota.Cells[1, 1]);
  if not SafeGetPlat(POrigem, PlatOrigem) then
  begin
    ShowMessage('Origem inválida ou sem coordenadas: ' + POrigem);
    Exit;
  end;

  NoOrigem.X := PlatOrigem.CoordX;
  NoOrigem.Y := PlatOrigem.CoordY;

  SetupChart;

  // inicia traçado
  RChartMapa.MoveTo(NoOrigem.X, NoOrigem.Y);
  MarkPoint(NoOrigem, POrigem);

  // percorre destinos
  for i := 2 to StrGridRota.RowCount - 1 do
  begin
    PDestino := Trim(StrGridRota.Cells[1, i]);
    if PDestino = '' then
      Continue;

    if not SafeGetPlat(PDestino, PlatDestino) then
    begin
      // se não tem coord, só pula (ou desenha só o texto no ponto atual)
      RChartMapa.Text(NoOrigem.X, NoOrigem.Y, 8, '  ' + PDestino + ' (sem coord)');
      Continue;
    end;

    NoDestino.X := PlatDestino.CoordX;
    NoDestino.Y := PlatDestino.CoordY;

    // Calcula distância e tempo (usando coords geográficas)
    if TProgramacaoRTUtils.GetDistanceAndTravelMin(
      PlatOrigem.Latitude, PlatOrigem.Longitude,
      PlatDestino.Latitude, PlatDestino.Longitude,
      DistKm,
      TravelMin
    ) then
      MarkPoint(NoDestino, PDestino + '  ' + FormatDistMin(DistKm, TravelMin))
    else
      MarkPoint(NoDestino, PDestino + '  N/D');

    // desenha seta entre os pontos do mapa
    RChartMapa.Arrow(NoOrigem.X, NoOrigem.Y, NoDestino.X, NoDestino.Y, 12);

    // avança origem para o próximo trecho
    NoOrigem := NoDestino;
    PlatOrigem := PlatDestino;
  end;

  RChartMapa.Repaint;
  actMapaRoteiro.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.actGerarLinhasTodasExecute(Sender: TObject);
begin
  BooleanIndividual:= false;
  StatusBarMapa.Panels[0].Text:= '';
  AutoFitStatusBar(StatusBarMapa);
  FrmDataModule.ADOQueryRoteamento.First;
  while not FrmDataModule.ADOQueryRoteamento.Eof do
  begin
    actGerarLinhas.Execute;
    FrmDataModule.ADOQueryRoteamento.Next;
  end;
  FrmDataModule.ADOQueryRoteamento.First;
end;

procedure TFrmGerenciarEmbarcacoes.CheckBox1Click(Sender: TObject);
begin
  if CheckBox1.Checked=true then
    RChartMapa.GridStyle:= gsDotLines
  else
    RChartMapa.GridStyle:= gsNone;
end;

procedure TFrmGerenciarEmbarcacoes.ClassificarCrescentepelacolunaselecionada1Click(
  Sender: TObject);
begin
  FrmPrincipal.clasifica(RLImpressao,RLImpressao.Col,true);
end;

procedure TFrmGerenciarEmbarcacoes.ColorBoxLinhaCloseUp(Sender: TObject);
begin
  try
    if not (DBGridRoteamento.DataSource.State in [dsEdit, dsInsert]) then
      DBGridRoteamento.DataSource.Edit;

    DBGridRoteamento.DataSource.DataSet.FieldByName('linhaCor').AsInteger:=
    ColorBoxLinha.Selected;
  except

  end;
end;

procedure TFrmGerenciarEmbarcacoes.ColorBoxLinhaMouseLeave(Sender: TObject);
begin
  ColorBoxLinha.Visible:= false;
end;

procedure TFrmGerenciarEmbarcacoes.ComboBoxOrigemInicialCloseUp(
  Sender: TObject);
begin
  StrGridRota.Cells[1,StrGridRota.Row]:= (ComboBoxOrigemInicial.Text);
  ComboBoxOrigemInicial.Visible := False;
end;

procedure TFrmGerenciarEmbarcacoes.ComboBoxOrigemInicialMouseLeave(
  Sender: TObject);
begin
  ComboBoxOrigemInicial.Visible := False;
end;

procedure TFrmGerenciarEmbarcacoes.ComboBoxPlataformaCloseUp(Sender: TObject);
begin
  StrGridRota.Cells[1,StrGridRota.Row]:= (ComboBoxPlataforma.Text);
  ComboBoxPlataforma.Visible := False;
end;

procedure TFrmGerenciarEmbarcacoes.ComboBoxPlataformaMouseLeave(
  Sender: TObject);
begin
  ComboBoxPlataforma.Visible := False;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridProcuraGeralCellClick(
  Column: TColumn);
begin
  try
    if (Self.DBGridProcuraGeral.SelectedField.DataType = ftBoolean)AND
    (Column.Field.FieldName = 'booleanSelecao') then
    begin
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Edit;
      FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
      FieldByName('booleanSelecao').AsBoolean:=
      not Self.DBGridProcuraGeral.SelectedField.AsBoolean;
      FrmDataModule.ADOQueryConsultaExecutantes_DataProgramacao.Post;
      //Numero de selecionados
      actSelecionados.Execute;
    end;
  except

  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridProcuraGeralDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridProcuraGeral,ColunasLayoutProcuraGeral,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'InseridoProgramacaoTransporte') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
    FieldByName('InseridoProgramacaoTransporte').AsBoolean then
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +10,Rect.Top + 1, 92);
    end
    else
    begin
      TDBGrid(Sender).Canvas.FillRect(Rect);
      FrmPrincipal.ImageList1.Draw(TDBGrid(Sender).Canvas, Rect.Left +10,Rect.Top + 1, 93);
    end;
  end
  else if (Column.Field.FieldName = 'RequisitantePT') then
  begin
    if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
    FieldByName('RequisitantePT').AsBoolean then
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
    if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Aprovado' then
    begin
      DBGridProcuraGeral.Canvas.Brush.Color:= clLime;
      DBGridProcuraGeral.Font.Color:= clBlack;
      DBGridProcuraGeral.Canvas.FillRect(Rect);
      DBGridProcuraGeral.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Cancelado' then
    begin
      DBGridProcuraGeral.Canvas.Brush.Color:= clRed;
      DBGridProcuraGeral.Font.Color:= clBlack;
      DBGridProcuraGeral.Canvas.FillRect(Rect);
      DBGridProcuraGeral.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceConsultaExecutantes_DataProgramacao.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Mudança' then
    begin
      DBGridProcuraGeral.Canvas.Brush.Color:= clYellow;
      DBGridProcuraGeral.Font.Color:= clBlack;
      DBGridProcuraGeral.Canvas.FillRect(Rect);
      DBGridProcuraGeral.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'booleanSelecao') then   //  mudança no código
  begin
    Self.DBGridProcuraGeral.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridProcuraGeral.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridProcuraProgramadosDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridProcuraProgramados,ColunasLayoutProgramados,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'RequisitantePT') then
  begin
    if FrmDataModule.DataSourceProgramados.DataSet.
    FieldByName('RequisitantePT').AsBoolean then
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
    if FrmDataModule.DataSourceProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Aprovado' then
    begin
      DBGridProcuraProgramados.Canvas.Brush.Color:= clLime;
      DBGridProcuraProgramados.Font.Color:= clBlack;
      DBGridProcuraProgramados.Canvas.FillRect(Rect);
      DBGridProcuraProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceProgramados.DataSet.
    FieldByName('StatusProgramacao').AsString = 'Cancelado' then
    begin
      DBGridProcuraProgramados.Canvas.Brush.Color:= clRed;
      DBGridProcuraProgramados.Font.Color:= clBlack;
      DBGridProcuraProgramados.Canvas.FillRect(Rect);
      DBGridProcuraProgramados.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end;
  end
  else if (Column.Field.FieldName = 'booleanSelecao') then   //  mudança no código
  begin
    Self.DBGridProcuraProgramados.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridProcuraProgramados.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridRoteamentoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridRoteamento,ColunasLayoutRoteamento,State,Rect,DataCol,Column);
  if (gdFocused in State) then
  begin
    //Cor da linha
    if (Column.Field.FieldName= 'linhaCor') then
    begin
      with ColorBoxLinha do
      begin
        Left := Rect.Left + DBGridRoteamento.Left + 1;
        Top := Rect.Top + DBGridRoteamento.Top + 1;
        Width := Rect.Right - Rect.Left + 2;
        Width := Rect.Right - Rect.Left + 2;
        Height := Rect.Bottom - Rect.Top + 2;
        Selected:= DBGridRoteamento.DataSource.
        DataSet.FieldByName('linhaCor').AsInteger;
        Visible := True;
      end;
    end
    else
      ColorBoxLinha.Visible := False;
    //hora 1° viagem
    if (Column.Field.FieldName= 'HoraRoteamento') then
    begin
      with MaskEditHora do
      begin
        Left := Rect.Left + DBGridRoteamento.Left + 1;
        Top := Rect.Top + DBGridRoteamento.Top + 1;
        Width := Rect.Right - Rect.Left + 2;
        Width := Rect.Right - Rect.Left + 2;
        Height := Rect.Bottom - Rect.Top + 2;
        MaskEditHora.Text:= DBGridRoteamento.DataSource.
        DataSet.FieldByName('HoraRoteamento').AsString;
        Visible := True;
      end;
    end
    else
      MaskEditHora.Visible := False;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridRoteamentoKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridRoteamento.DataSource.DataSet.RecNo=
  DBGridRoteamento.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridRotaExecutantesCellClick(
  Column: TColumn);
begin
  try
    if (Self.DBGridRotaExecutantes.SelectedField.DataType = ftBoolean)AND
    (Column.Field.FieldName = 'booleanSelecao') then
    begin
      DBGridRotaExecutantes.Options:=
      [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
      FrmDataModule.ADOQueryTM_RotaExecutantes.Edit;
      FrmDataModule.DataSourceTM_RotaExecutantes.DataSet.
      FieldByName('booleanSelecao').AsBoolean:= not Self.DBGridRotaExecutantes.SelectedField.AsBoolean;
      FrmDataModule.ADOQueryTM_RotaExecutantes.Post;
      actSelecionadosRotaExecutantes.Execute;
    end
    else
      DBGridRotaExecutantes.Options:=
      [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
  except
  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridRotaExecutantesDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
  Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
  var
    CheckBoxRectangle : TRect;
begin
  FrmPrincipal.GridZebrado(DBGridRotaExecutantes,ColunasLayoutRotaExecutantes,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'booleanSelecao') then   //  mudança no código
  begin
    Self.DBGridRotaExecutantes.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridRotaExecutantes.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.DBGridRotaExecutantesKeyDown(
  Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key in [VK_DOWN]) AND (DBGridRotaExecutantes.DataSource.DataSet.RecNo=
  DBGridRotaExecutantes.DataSource.DataSet.RecordCount) then
    Abort;
end;

procedure TFrmGerenciarEmbarcacoes.editarRoteamento(Tipo: Integer);
  var
    CapacidadePAX,TempoTransbordo: Integer;
    VelocidadeLancha: Double;
begin
  try
  case Tipo of
    0: FrmDataModule.ADOQueryRoteamento.Edit;
    1: FrmDataModule.ADOQueryRoteamento.Insert;
  end;
  CapacidadePAX:= 0;
  VelocidadeLancha:= 0;
  TempoTransbordo:= 1;
  FrmDataModule.ADOQueryTM_Embarcacao.First;
  while not FrmDataModule.ADOQueryTM_Embarcacao.Eof do
  begin
    if FrmDataModule.DataSourceTM_Embarcacao.DataSet.
    FieldByName('NomeEmbarcacao').AsString = FrmDataModule.DataSourceRoteamento.DataSet.
    FieldByName('NomeEmbarcacao').AsString then
    begin
      CapacidadePAX:= FrmDataModule.DataSourceTM_Embarcacao.DataSet.
      FieldByName('CapacidadePAX').AsInteger;
      VelocidadeLancha:= FrmDataModule.DataSourceTM_Embarcacao.DataSet.
      FieldByName('VelocidadeLancha').AsFloat;
      TempoTransbordo:= FrmDataModule.DataSourceTM_Embarcacao.DataSet.
      FieldByName('TempoTransbordo').AsInteger;
    end;
    FrmDataModule.ADOQueryTM_Embarcacao.Next;
  end;
  //=============================================
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('CapacidadePAX').AsInteger:= CapacidadePAX;
  //=============================================
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('VelocidadeLancha').AsFloat:= VelocidadeLancha;
  //=============================================
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('DataRoteamento').AsString:=
  FrmPrincipal.corrigirData(DateTimePickerProgramacao.Date);
  //=============================================
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('TempoTransbordo').AsInteger:= TempoTransbordo;
  //=============================================
  FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('linhaCor').AsInteger:= FrmPrincipal.CoresAuto[Random(12)];
  //=============================================
  if (FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('HoraRoteamento').AsString = '')OR(FrmDataModule.DataSourceRoteamento.DataSet.
  FieldByName('HoraRoteamento').AsString = '  :  ') then
    FrmDataModule.DataSourceRoteamento.DataSet.
    FieldByName('HoraRoteamento').AsString:= '06:30';
  FrmDataModule.ADOQueryRoteamento.Post;
  except

  end;
end;

procedure TFrmGerenciarEmbarcacoes.edtOrigemKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    actProcuraExecutantes.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.edtTipoEtapaServicoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if key = #13 then
    actProcuraExecutantes.Execute;
end;

procedure TFrmGerenciarEmbarcacoes.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmGerenciarEmbarcacoes:=nil;
end;

procedure TFrmGerenciarEmbarcacoes.FormCreate(Sender: TObject);
begin
  FrmPrincipal.ProgressBarIncializa(7,'Inicializando....');
  FrmDataModule.setFilterDBGrid(DBGridRoteamento);
  FrmDataModule.setFilterDBGrid(DBGridRotaExecutantes);
  FrmDataModule.setFilterDBGrid(DBGridProcuraGeral);
  FrmDataModule.setFilterDBGrid(DBGridProcuraProgramados);

  PageControlPrincipal.TabIndex:= 0;
  StrGridRota.Cells[1,0]:= 'Plataforma';
  StrGridRota.Cells[2,0]:= 'N° Exec.';
  StrGridRota.Cells[3,0]:= 'Distância (km)';
  StrGridRota.Cells[4,0]:= '∑ Dist. (km)';
  StrGridRota.Cells[5,0]:= 'Latitude';
  StrGridRota.Cells[6,0]:= 'Longitude';
  //Configurações iniciais
  DBGridProcuraProgramados.Visible:= false;
  DBGridProcuraGeral.Visible:= true;
  DBGridProcuraGeral.Align:= alClient;
  FrmPrincipal.ProgressBarIncremento(1);
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  FrmDataModule.ADOQueryTM_Embarcacao.Active:= false;
  FrmDataModule.ADOQueryTM_Embarcacao.Active:= true;
  FrmDataModule.ADOQueryPlataforma.Active:= false;
  FrmDataModule.ADOQueryPlataforma.Active:= true;
  FrmDataModule.ADOQueryAux_Rota_Distribuicao.Active:= false;
  FrmDataModule.ADOQueryAux_Rota_Distribuicao.Active:= true;
  FrmPrincipal.ProgressBarIncremento(1);
  DateTimePickerProgramacao.Date:= IncDay(now,1);
  FrmPrincipal.SetupGridFilterPickListSQL(FrmDataModule.ADOConnectionConsulta,'NomeEmbarcacao',
  'SELECT NomeEmbarcacao FROM tblEmbarcacao '+
  'WHERE ((TipoEmbarcacao = "Lancha de passageiros")OR(TipoEmbarcacao = "Mergulho"))'+
  'AND((StatusEmbarcacao = "Operando"))'+
  'ORDER BY NomeEmbarcacao;',DBGridRoteamento,true);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.carregarComboBox(FrmDataModule.ADOConnectionConsulta,'Plataforma',
  'SELECT tblPlataforma.* FROM tblPlataforma ORDER BY Plataforma;',ComboBoxPlataforma);
  ComboBoxPlataforma.Height:= StrGridRota.DefaultRowHeight;
  //====================================
  StrGridRota.ColCount:= 7;
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO)) then
  begin
    //Roteamento
    DBGridRoteamento.ReadOnly:= false;
    btnExcluirRota.Enabled:= true;
    btnInserir.Enabled:= true;
    btnSalvar.Enabled:= true;
    actVincular.Enabled:= true;
    btnExcluirVinculos.Enabled:= true;
    actNumerar.Enabled:= true;
    StrGridResumo.Enabled:= true;
    StrGridRota.Enabled:= true;
  end
  else
  begin
    //Roteamento
    DBGridRoteamento.ReadOnly:= true;
    btnExcluirRota.Enabled:= false;
    btnInserir.Enabled:= false;
    btnSalvar.Enabled:= false;
    actVincular.Enabled:= false;
    btnExcluirVinculos.Enabled:= false;
    actNumerar.Enabled:= false;
    StrGridResumo.Enabled:= false;
    StrGridRota.Enabled:= false;
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  FrmDataModule.setFilterDBGrid(DBGridProcuraProgramados);
  FrmDataModule.setFilterDBGrid(DBGridProcuraGeral);
  FrmDataModule.setFilterDBGrid(DBGridRotaExecutantes);
  FrmDataModule.setFilterDBGrid(DBGridRoteamento);
  FrmPrincipal.ProgressBarIncremento(1);
  actLocalizarRoteamentos.Execute;
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.actMatrizExecutanteAPLAT.Execute;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmGerenciarEmbarcacoes.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmGerenciarEmbarcacoes.RChartMapaDblClick(Sender: TObject);
begin
  if booleanZoomMapa = true then
  begin
    booleanZoomMapa := false;
    actMapaPan.Execute;
    RChartMapa.Repaint;
  end
  else
  begin
    booleanZoomMapa := true;
    actMapaZoomDinamico.Execute;
    RChartMapa.Repaint;
  end;
end;

procedure TFrmGerenciarEmbarcacoes.RChartMapaKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
    var
      ptMouse: TPointFloat;
begin
  //Coordenadas X,Y do mouse
  ptMouse.X:= RChartMapa.MousePosX;
  ptMouse.Y:= RChartMapa.MousePosY;
  if btnCotas.Down then
  begin
    FrmPrincipal.selecaoPlataforma(RChartMapa,ptMouse,StatusBarMapa);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.RChartMapaMouseLeave(Sender: TObject);
begin
  screen.Cursor:= crDefault;
end;

procedure TFrmGerenciarEmbarcacoes.RChartMapaMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  edtXFEA.Value := RChartMapa.MousePosX;
  edtYFEA.Value  := RChartMapa.MousePosY;
end;

procedure TFrmGerenciarEmbarcacoes.RChartMapaZoomPan(Sender: TObject);
  var
    rangeXLow,rangeXHigh,rangeYHigh,rangeYlow,rangeX,rangeY: Double;
    YLow,YHigh,XLow,XHigh,CentroX,CentroY,Largura,Altura: Double;
begin
  with RChartMapa do
  begin
    rangeXHigh:= Scale1X.RangeHigh;
    rangeXLow:= Scale1X.RangeLow;
    rangeYHigh:= Scale1Y.RangeHigh;
    rangeYlow:= Scale1Y.RangeLow;
    rangeX:= rangeXHigh-rangeXLow;
    rangeY:= rangeYHigh-rangeYLow;
    Largura:= (Width-LRim-RRim);
    Altura:= (Height -TRim-BRim);
    //====================================
    if rangeX > rangeY*(Largura/Altura) then
      rangeY:= rangeX/(Largura/Altura)
    else
      rangeX:= rangeY*(Largura/Altura);
    //=====================================
    CentroX:= ((rangeXHigh-rangeXLow)/2)+rangeXLow;
    CentroY:= ((rangeYHigh-rangeYLow)/2)+rangeYLow;
    XLow:= CentroX-(rangeX/2);
    XHigh:= CentroX+(rangeX/2);
    YLow:= CentroY-(rangeY/2);
    YHigh:= CentroY+(rangeY/2);
    setRange(0,XLow,YLow,XHigh,YHigh);
  end;

end;

procedure TFrmGerenciarEmbarcacoes.RLImpressaoFixedCellClick(Sender: TObject;
  ACol, ARow: Integer);
  var
    i: Integer;
begin
  FrmPrincipal.clasifica(RLImpressao,ACol,true);
  if RLImpressao.Cells[0,0] = 'N°' then
  begin
    for i := 1 to RLImpressao.RowCount-1 do
      RLImpressao.Cells[0,i]:= IntToStr(i);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.PageControlPrincipalChange(Sender: TObject);
begin
  if PageControlPrincipal.TabIndex = 0 then
    actProcuraExecutantes.Execute
end;

procedure TFrmGerenciarEmbarcacoes.PanelTituloAjudaMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  FrmPrincipal.Capturing := true;
  FrmPrincipal.MouseDownSpot.X := X;
  FrmPrincipal.MouseDownSpot.Y := Y;
end;

procedure TFrmGerenciarEmbarcacoes.PanelTituloAjudaMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.PanelTituloAjudaMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  if FrmPrincipal.Capturing then
  begin
    FrmPrincipal.Capturing := false;
    PanelAjuda.Left := PanelAjuda.Left - (FrmPrincipal.MouseDownSpot.X - X);
    PanelAjuda.Top := PanelAjuda.Top - (FrmPrincipal.MouseDownSpot.Y - Y);
  end;
end;

procedure TFrmGerenciarEmbarcacoes.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
