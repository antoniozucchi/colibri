unit untSituacaoEquipamentoAcesso;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ExtCtrls, Vcl.DBCtrls, Vcl.ToolWin, Vcl.ComCtrls, Vcl.StdCtrls, Vcl.Mask,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,
  Vcl.Buttons, untDBGridFilter;

type
  TFrmSituacaoEquipamentoAcesso = class(TForm)
    ToolBar1: TToolBar;
    PanelTitulo: TPanel;
    DBGridSituacaoEquipamentoAcesso: TFilterDBGrid;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    StatusBar1: TStatusBar;
    Splitter1: TSplitter;
    Panel2: TPanel;
    PanelTituloNotas: TPanel;
    DBMemo1: TDBMemo;
    DBNavigator: TDBNavigator;
    btnClearFiltro: TToolButton;
    btnLayout: TToolButton;
    btnExcel: TToolButton;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridSituacaoEquipamentoAcessoDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGridSituacaoEquipamentoAcessoKeyPress(Sender: TObject; var Key: Char);
    procedure actProcurarExecute(Sender: TObject);
  private
    { Private declarations }
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
    procedure DrawCollumCellDBGrid(Tabela: TFilterDBGrid;
      Const Rect: TRect; DataCol: Integer;
      Column: TColumn; State: TGridDrawState; FieldName: String);
  public
    { Public declarations }
  end;

var
  FrmSituacaoEquipamentoAcesso: TFrmSituacaoEquipamentoAcesso;

implementation
  uses untDataModule,untPrincipal;
{$R *.dfm}

procedure TFrmSituacaoEquipamentoAcesso.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= BuildFilterSQL(DBGridSituacaoEquipamentoAcesso,false);
  if SQLString <> '' then
    SQLString:= ' AND ' + SQLString;

  SQLBase:= 'SELECT tblPlataforma.* FROM tblPlataforma '+
  'WHERE (BooleanPlataforma = True) '+
  SQLString+' ORDER BY Plataforma;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryPlataformaControle,StatusBar1);
end;

procedure TFrmSituacaoEquipamentoAcesso.DrawCollumCellDBGrid(Tabela: TFilterDBGrid;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState;
  FieldName: String);
begin
  if (Column.Field.FieldName = FieldName) then
  begin
    if Tabela.DataSource.DataSet.
    FieldByName(FieldName).AsString = 'OK' then
    begin
      Tabela.Canvas.Brush.Color:= clLime;
      Tabela.Font.Color:= clBlack;
      Tabela.Canvas.FillRect(Rect);
      Tabela.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if Tabela.DataSource.DataSet.
    FieldByName(FieldName).AsString = 'FO' then
    begin
      Tabela.Canvas.Brush.Color:= clRed;
      Tabela.Font.Color:= clBlack;
      Tabela.Canvas.FillRect(Rect);
      Tabela.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if Tabela.DataSource.DataSet.
    FieldByName(FieldName).AsString = 'RO' then
    begin
      Tabela.Canvas.Brush.Color:= clYellow;
      Tabela.Font.Color:= clBlack;
      Tabela.Canvas.FillRect(Rect);
      Tabela.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
    else if Tabela.DataSource.DataSet.
    FieldByName(FieldName).AsString = 'N/A' then
    begin
      Tabela.Canvas.Brush.Color:= clSilver;
      Tabela.Font.Color:= clBlack;
      Tabela.Canvas.FillRect(Rect);
      Tabela.DefaultDrawColumnCell(Rect, DataCol,Column, State);
    end
  end
end;

procedure TFrmSituacaoEquipamentoAcesso.DBGridSituacaoEquipamentoAcessoDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoGD');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoBCI');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoLinhaBCI');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoBalsa');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoSurfer');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoAqua');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoSOV');
  DrawCollumCellDBGrid(DBGridSituacaoEquipamentoAcesso,Rect,DataCol,Column,State,'SituacaoDegraus');
end;

procedure TFrmSituacaoEquipamentoAcesso.DBGridSituacaoEquipamentoAcessoKeyPress(Sender: TObject;
  var Key: Char);
begin
  if (DBGridSituacaoEquipamentoAcesso.SelectedField.FieldName = 'Plataforma')OR
  (DBGridSituacaoEquipamentoAcesso.SelectedField.FieldName = 'Status') then
  begin
    Key:= #0;
  end;
end;

procedure TFrmSituacaoEquipamentoAcesso.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action:= caFree;
  FrmSituacaoEquipamentoAcesso:=nil;
end;

procedure TFrmSituacaoEquipamentoAcesso.FormCreate(Sender: TObject);
begin
  FrmPrincipal.MDIChildCreated(self.Handle);

  if (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) then
  begin
    DBNavigator.Enabled:= true;
    DBGridSituacaoEquipamentoAcesso.ReadOnly:= false;
  end
  else
  begin
    DBNavigator.Enabled:= false;
    DBGridSituacaoEquipamentoAcesso.ReadOnly:= true;
  end;
  //Inicializacao
  actProcurar.Execute;
end;

procedure TFrmSituacaoEquipamentoAcesso.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmSituacaoEquipamentoAcesso.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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

