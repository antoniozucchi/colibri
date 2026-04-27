unit untFrmPrioridadeDistribuicao;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Data.Win.ADODB, Vcl.ExtCtrls,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Grids, Vcl.DBGrids, untDBGridFilter,
  Vcl.StdCtrls, Vcl.Buttons, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan;

type
  TFrmPrioridadeDistribuicao = class(TForm)
    qryPrioridades: TADOQuery;
    dsPrioridades: TDataSource;
    PanelTitulo: TPanel;
    GridPanel1: TGridPanel;
    btnSalvar: TBitBtn;
    btnFechar: TBitBtn;
    DBGrid1: TFilterDBGrid;
    ToolBar1: TToolBar;
    btnCarregar: TBitBtn;
    DateTimePickerData: TDateTimePicker;
    ActionManager1: TActionManager;
    actCarregar: TAction;
    actSalvar: TAction;
    actFechar: TAction;

    btnExcel: TToolButton;
    btnClearFiltro: TToolButton;
    actProcurar: TAction;
    procedure FormCreate(Sender: TObject);
    procedure qryPrioridadesBeforePost(DataSet: TDataSet);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actFecharExecute(Sender: TObject);
    procedure actSalvarExecute(Sender: TObject);
    procedure actCarregarExecute(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
  private
    FConnColibri: TADOConnection;
    FConnConsulta: TADOConnection;
    procedure CarregarPrioridades;
    { Private declarations }
  public
    { Public declarations }
    procedure Inicializar(
      AConnColibri, AConnConsulta: TADOConnection; AData: TDateTime);
  end;

var
  FrmPrioridadeDistribuicao: TFrmPrioridadeDistribuicao;

implementation
  uses untPrincipal, untDataModule;

{$R *.dfm}

procedure TFrmPrioridadeDistribuicao.Inicializar(
  AConnColibri, AConnConsulta: TADOConnection; AData: TDateTime);
begin
  FConnColibri := AConnColibri;
  FConnConsulta := AConnConsulta;

  qryPrioridades.Connection := FConnConsulta;
  DateTimePickerData.Date := AData;

  CarregarPrioridades;
end;

procedure TFrmPrioridadeDistribuicao.qryPrioridadesBeforePost(
  DataSet: TDataSet);
var
  V: Variant;
begin
  V := DataSet.FieldByName('PrioridadeDistribuicao').Value;

  if not VarIsNull(V) and not VarIsEmpty(V) then
  begin
    if DataSet.FieldByName('PrioridadeDistribuicao').AsInteger < 1 then
      raise Exception.Create('A PrioridadeDistribuicao deve ser maior ou igual a 1.');
  end;
end;

procedure TFrmPrioridadeDistribuicao.actCarregarExecute(Sender: TObject);
begin
  CarregarPrioridades;
end;

procedure TFrmPrioridadeDistribuicao.actFecharExecute(Sender: TObject);
begin
  Close;
end;

procedure TFrmPrioridadeDistribuicao.actProcurarExecute(Sender: TObject);
begin
  CarregarPrioridades;
end;

procedure TFrmPrioridadeDistribuicao.actSalvarExecute(Sender: TObject);
begin
  if qryPrioridades.Active and (qryPrioridades.State in [dsEdit, dsInsert]) then
    qryPrioridades.Post;

  ShowMessage('Prioridades salvas com sucesso.');
end;

procedure TFrmPrioridadeDistribuicao.CarregarPrioridades;
var
  QryProg: TADOQuery;
  ListaDestinos: TStringList;
  SQLIn,SQLString: string;
  I: Integer;
begin
  qryPrioridades.Close;

  QryProg := TADOQuery.Create(nil);
  ListaDestinos := TStringList.Create;
  try

  SQLString:= BuildFilterSQL(DBGrid1,false);
    //Query de procura
    if SQLString <> '' then
      SQLString:= ' AND '+SQLString;
    // 1) Busca os destinos programados no banco COLIBRI
    QryProg.Connection := FConnColibri;
    QryProg.SQL.Text :=
      'SELECT DISTINCT pd.txtDestino ' +
      'FROM tblProgramacaoDiaria pd ' +
      'INNER JOIN tblProgramacaoExecutante pe ' +
      '  ON pd.idProgramacaoDiaria = pe.CodigoProgramacaoDiaria ' +
      'WHERE pd.DataProgramacao = :DataProg ' +
      '  AND pd.txtDestino IS NOT NULL ' +
      '  AND Trim(pd.txtDestino) <> '''' ' +
      '  AND pe.StatusProgramacao = ''Aprovado'' ' +
      'ORDER BY pd.txtDestino';

    QryProg.Parameters.ParamByName('DataProg').Value :=
      FormatDateTime('dd/mm/yyyy', DateTimePickerData.Date);

    QryProg.Open;

    while not QryProg.Eof do
    begin
      if Trim(QryProg.FieldByName('txtDestino').AsString) <> '' then
        ListaDestinos.Add(Trim(QryProg.FieldByName('txtDestino').AsString));
      QryProg.Next;
    end;

    if ListaDestinos.Count = 0 then
    begin
      qryPrioridades.Close;
      Exit;
    end;

    // 2) Monta IN (...) para consultar tblPlataforma no banco CONSULTA
    SQLIn := '';
    for I := 0 to ListaDestinos.Count - 1 do
    begin
      if SQLIn <> '' then
        SQLIn := SQLIn + ',';
      SQLIn := SQLIn + QuotedStr(StringReplace(ListaDestinos[I], '''', '''''', [rfReplaceAll]));
    end;

    qryPrioridades.Connection := FConnConsulta;
    qryPrioridades.SQL.Clear;
    qryPrioridades.SQL.Add(
      'SELECT ' +
      '    p.idPlataforma, ' +
      '    p.Plataforma, ' +
      '    p.NomeSAP, ' +
      '    p.RT_Modal, ' +
      '    p.PrioridadeDistribuicao ' +
      'FROM tblPlataforma p ' +
      'WHERE p.Plataforma IN (' + SQLIn + ') ' + SQLString +
      'ORDER BY IIf(p.PrioridadeDistribuicao Is Null, 0, 1), ' +
      '         p.PrioridadeDistribuicao, ' +
      '         p.Plataforma'
    );

    qryPrioridades.Open;

  finally
    ListaDestinos.Free;
    QryProg.Free;
  end;
end;

procedure TFrmPrioridadeDistribuicao.DBGrid1DrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  if qryPrioridades.Active then
  begin
    if qryPrioridades.FieldByName('PrioridadeDistribuicao').IsNull then
      DBGrid1.Canvas.Brush.Color := $00D9FFFF; // amarelo claro
  end;

  DBGrid1.DefaultDrawColumnCell(Rect, DataCol, Column, State);
end;

procedure TFrmPrioridadeDistribuicao.FormCreate(Sender: TObject);
begin
  DBGrid1.DataSource := dsPrioridades;
  dsPrioridades.DataSet := qryPrioridades;
end;

end.


