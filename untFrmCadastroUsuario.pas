unit untFrmCadastroUsuario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Mask, DBCtrls, StdCtrls, Buttons, ToolWin, ComCtrls, ExtCtrls, Grids,
  DBGrids, ActnList, PlatformDefaultStyleActnCtrls, ActnMan, Data.DB,
  System.Actions, Vcl.CheckLst;

type
  TFrmCadastroUsuario = class(TForm)
    Panel2: TPanel;
    ToolBar1: TToolBar;
    DBNavigatorComissao: TDBNavigator;
    BitBtn4: TBitBtn;
    ActionManager1: TActionManager;
    actExcel: TAction;
    ColunasLayout: TStringGrid;
    actFiltroInserir: TAction;
    actProcurar: TAction;
    actGridASC: TAction;
    actGridDESC: TAction;
    actSubstituirPor: TAction;
    actLimparFiltros: TAction;
    actFiltrosTabela: TAction;
    actProcuraFiltrosTabela: TAction;
    StatusBar1: TStatusBar;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    DBGridUsuarios: TDBGrid;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridUsuariosKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridUsuariosDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actExcelExecute(Sender: TObject);
    procedure actFiltroInserirExecute(Sender: TObject);
    procedure actGridASCExecute(Sender: TObject);
    procedure actGridDESCExecute(Sender: TObject);
    procedure actSubstituirPorExecute(Sender: TObject);
    procedure actLimparFiltrosExecute(Sender: TObject);
    procedure actFiltrosTabelaExecute(Sender: TObject);
    procedure actProcuraFiltrosTabelaExecute(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
    procedure DBGridUsuariosTitleClick(Column: TColumn);
  private
    { Private declarations }
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
  public
    { Public declarations }
  end;

var
  FrmCadastroUsuario: TFrmCadastroUsuario;

implementation
uses untPrincipal,untFrmLogin,untDataModule;
{$R *.dfm}

procedure TFrmCadastroUsuario.actExcelExecute(Sender: TObject);
begin
  FrmPrincipal.GerarExcel(DBGridUsuarios,'Usuarios');
end;

procedure TFrmCadastroUsuario.actFiltroInserirExecute(Sender: TObject);
begin
  FrmPrincipal.inserirProcura(DBGridUsuarios,ColunasLayout);
  actProcurar.Execute;
  if (FrmPrincipal.PanelFiltrosTabela.Visible)AND(FrmPrincipal.PanelAjuda1.Visible) then
    actFiltrosTabela.Execute;
end;

procedure TFrmCadastroUsuario.actFiltrosTabelaExecute(Sender: TObject);
begin
  FrmPrincipal.btnProcurarFiltrosTabela.Action:= actProcuraFiltrosTabela;
  FrmPrincipal.FiltrosTabela(DBGridUsuarios,ColunasLayout);
end;

procedure TFrmCadastroUsuario.actGridASCExecute(Sender: TObject);
begin
  FrmPrincipal.ClassificaDBGrid(DBGridUsuarios,FrmDataModule.ADOQueryCadastroUsuario,0);
end;

procedure TFrmCadastroUsuario.actGridDESCExecute(Sender: TObject);
begin
  FrmPrincipal.ClassificaDBGrid(DBGridUsuarios,FrmDataModule.ADOQueryCadastroUsuario,1);
end;

procedure TFrmCadastroUsuario.actLimparFiltrosExecute(Sender: TObject);
begin
  FrmPrincipal.LimparColunasFiltro(DBGridUsuarios,ColunasLayout);
  actProcurar.Execute;
  if (FrmPrincipal.PanelFiltrosTabela.Visible)AND(FrmPrincipal.PanelAjuda1.Visible) then
    actFiltrosTabela.Execute;
end;

procedure TFrmCadastroUsuario.actProcuraFiltrosTabelaExecute(Sender: TObject);
begin
  frmPrincipal.CarregaFiltrosProcura(ColunasLayout);
  actProcurar.Execute;
end;

procedure TFrmCadastroUsuario.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,true);
  SQLBase:= 'SELECT tblUsuario.* FROM tblUsuario '+
  SQLString+' ORDER BY Nome;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryCadastroUsuario,StatusBar1);
end;

procedure TFrmCadastroUsuario.actSubstituirPorExecute(Sender: TObject);
begin
  FrmPrincipal.SubstituirPor(DBGridUsuarios,FrmDataModule.ADOQueryCadastroUsuario,
  FrmDataModule.DataSourceCadastroUsuario);
end;

procedure TFrmCadastroUsuario.DBGridUsuariosDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  FrmPrincipal.GridZebrado(DBGridUsuarios,ColunasLayout,State,Rect,DataCol,Column);
end;

procedure TFrmCadastroUsuario.DBGridUsuariosKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (DBGridUsuarios.SelectedField.FieldName = 'Perfil') then
    Key := #0;
end;

procedure TFrmCadastroUsuario.DBGridUsuariosTitleClick(Column: TColumn);
begin
  FrmPrincipal.configurarFiltro(1,Column.FieldName,IntToStr(Column.Index),
  Column.ReadOnly,actFiltroInserir,actGridASC,actGridDESC,actSubstituirPor);
  //======================================================
  FrmPrincipal.titleGrid(ColunasLayout,'Colibri',FrmDataModule.ADOQueryCadastroUsuario.SQL.Text);
end;

procedure TFrmCadastroUsuario.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmCadastroUsuario:=nil;
end;

procedure TFrmCadastroUsuario.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  //Incicialização
  FrmPrincipal.SetUpColunasLayout(DBGridUsuarios, ColunasLayout);
  actProcurar.Execute;
end;

procedure TFrmCadastroUsuario.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmCadastroUsuario.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
