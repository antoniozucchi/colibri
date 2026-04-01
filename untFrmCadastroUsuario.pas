unit untFrmCadastroUsuario;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, Mask, DBCtrls, StdCtrls, Buttons, ToolWin, ComCtrls, ExtCtrls, Grids,
  DBGrids, ActnList, PlatformDefaultStyleActnCtrls, ActnMan, Data.DB,
  System.Actions, Vcl.CheckLst, untDBGridFilter;

type
  TFrmCadastroUsuario = class(TForm)
    Panel2: TPanel;
    ToolBar1: TToolBar;
    DBNavigatorComissao: TDBNavigator;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    StatusBarRegrasRecolhimento: TStatusBar;
    DBGridUsuarios: TFilterDBGrid;
    btnClearFiltro: TToolButton;
    btnLayout: TToolButton;
    btnExcel: TToolButton;
    ColunasLayout: TStringGrid;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridUsuariosKeyPress(Sender: TObject; var Key: Char);
    procedure DBGridUsuariosDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure actProcurarExecute(Sender: TObject);
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

procedure TFrmCadastroUsuario.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,true);
  SQLBase:= 'SELECT tblUsuario.* FROM tblUsuario '+
  SQLString+' ORDER BY Nome;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryCadastroUsuario,StatusBarRegrasRecolhimento);
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
  FrmDataModule.setFilterDBGrid(DBGridUsuarios);
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
