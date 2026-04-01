unit untControleGeradores;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ExtCtrls, Vcl.DBCtrls, Vcl.ToolWin, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.Mask,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,
  Vcl.Buttons, untDBGridFilter;

type
  TFrmControleGeradores = class(TForm)
    ToolBar1: TToolBar;
    DBNavigatorGerador: TDBNavigator;
    PanelTitulo: TPanel;
    DBGridGerador: TFilterDBGrid;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    StatusBar1: TStatusBar;
    ColunasLayout: TStringGrid;
    Splitter1: TSplitter;
    Panel2: TPanel;
    PanelTituloPltaformas: TPanel;
    DBMemo1: TDBMemo;
    btnClearFiltro: TToolButton;
    btnLayput: TToolButton;
    btnExcel: TToolButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridGeradorDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGridGeradorKeyPress(Sender: TObject; var Key: Char);
    procedure actProcurarExecute(Sender: TObject);
  private
    { Private declarations }
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
  public
    { Public declarations }
  end;

var
  FrmControleGeradores: TFrmControleGeradores;

implementation
  uses untDataModule,untPrincipal;
{$R *.dfm}

procedure TFrmControleGeradores.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,true);
  SQLBase:= 'SELECT tblGerador.* FROM tblGerador '+
  SQLString+' ORDER BY Plataforma;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryGeradores,StatusBar1);
end;

procedure TFrmControleGeradores.DBGridGeradorDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridGerador,ColunasLayout,State,Rect,DataCol,Column);
  if (Column.Field.FieldName = 'Status') then
  begin
    if FrmDataModule.DataSourceGeradores.DataSet.
    FieldByName('Status').AsString = 'Operando' then
    begin
      DBGridGerador.Canvas.Brush.Color:= clLime;
      DBGridGerador.Font.Color:= clBlack;
      DBGridGerador.Canvas.FillRect(Rect);
      DBGridGerador.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if FrmDataModule.DataSourceGeradores.DataSet.
    FieldByName('Status').AsString = 'Fora de operação' then
    begin
      DBGridGerador.Canvas.Brush.Color:= clRed;
      DBGridGerador.Font.Color:= clBlack;
      DBGridGerador.Canvas.FillRect(Rect);
      DBGridGerador.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
  end;
end;

procedure TFrmControleGeradores.DBGridGeradorKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (DBGridGerador.SelectedField.FieldName = 'Plataforma')OR
  (DBGridGerador.SelectedField.FieldName = 'Status') then
  begin
    Key:= #0;
  end;
end;

procedure TFrmControleGeradores.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmControleGeradores:=nil;
end;

procedure TFrmControleGeradores.FormCreate(Sender: TObject);
begin
  FrmPrincipal.MDIChildCreated(self.Handle);

  if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) then
  begin
    DBNavigatorGerador.Enabled:= true;
    DBGridGerador.ReadOnly:= false;
  end
  else
  begin
    DBNavigatorGerador.Enabled:= false;
    DBGridGerador.ReadOnly:= true;
  end;
  //Incicialização
  FrmDataModule.setFilterDBGrid(DBGridGerador);
  FrmPrincipal.SetupGridFilterPickListSQL(FrmDataModule.ADOConnectionConsulta,'Plataforma',
  'SELECT Plataforma FROM tblPlataforma ORDER BY Plataforma',
  DBGridGerador,false);
  actProcurar.Execute;
end;

procedure TFrmControleGeradores.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmControleGeradores.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
