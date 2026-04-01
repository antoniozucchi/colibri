unit untConsultaServicosProgramados;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.StdCtrls, Vcl.Buttons,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.ExtCtrls, Vcl.Grids, Vcl.DBGrids,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,
  Vcl.CheckLst,DateUtils, untDBGridFilter;

type
  TFrmConsultaServicosProgramados = class(TForm)
    DBGridServicosProgramados: TFilterDBGrid;
    Panel2: TPanel;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    StatusBar1: TStatusBar;
    ColunasLayout: TStringGrid;
    ToolBar1: TToolBar;
    btnClearFiltro: TToolButton;
    btnLayout: TToolButton;
    btnExcel: TToolButton;
    Panel8: TPanel;
    dataInicio: TDateTimePicker;
    Panel24: TPanel;
    dataFim: TDateTimePicker;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridServicosProgramadosDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actProcurarExecute(Sender: TObject);
  private
    { Private declarations }
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
  public
    { Public declarations }
  end;

var
  FrmConsultaServicosProgramados: TFrmConsultaServicosProgramados;

implementation
  uses untPrincipal,untDataModule;
{$R *.dfm}

procedure TFrmConsultaServicosProgramados.actProcurarExecute(
  Sender: TObject);
  var
    SQLString,SQLBase,DataProcuraIncio,DataProcuraFim: String;
begin
  DataProcuraIncio:= (FormatDateTime('mm/dd/yyyy',dataInicio.Date));
  DataProcuraFim:= (FormatDateTime('mm/dd/yyyy',dataFim.Date));
  //=======================================================
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,false);
  if SQLString <> '' then
    SQLString:= ' AND '+SQLString;
  SQLBase:= 'SELECT tblProgramacaoDiaria.*, tblProgramacaoServico.* '+
  'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoServico '+
  'ON tblProgramacaoDiaria.idProgramacaoDiaria = '+
  'tblProgramacaoServico.CodigoProgramacaoDiaria '+
  'WHERE (DataProgramacao BETWEEN #'+DataProcuraIncio+'# and #'+DataProcuraFim+'#)'+
  SQLString+' ORDER BY DataProgramacao;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryConsultaServicosProgramados,
  StatusBar1);
end;

procedure TFrmConsultaServicosProgramados.DBGridServicosProgramadosDrawColumnCell(
  Sender: TObject; const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridServicosProgramados,ColunasLayout,State,Rect,DataCol,Column);
end;

procedure TFrmConsultaServicosProgramados.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmConsultaServicosProgramados:=nil;
end;

procedure TFrmConsultaServicosProgramados.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  dataInicio.Date:= IncDay(now,1);
  dataFim.Date:= IncDay(now,1);
  //Incicialização
  FrmDataModule.setFilterDBGrid(DBGridServicosProgramados);
  actProcurar.Execute;
end;

procedure TFrmConsultaServicosProgramados.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmConsultaServicosProgramados.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
