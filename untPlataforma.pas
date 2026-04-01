unit untPlataforma;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Vcl.Buttons, Vcl.CheckLst, Vcl.ExtCtrls, Vcl.ComCtrls,
  Vcl.Samples.Spin, Vcl.OleCtrls, SHDocVw, Vcl.DBCtrls,
  Vcl.ToolWin, Vcl.Grids, Vcl.DBGrids,ComOBJ, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, Vcl.Mask, Menus, Data.DB,
  SDL_rchart, Vcl.ExtDlgs,Math,SDL_sdlbase, SDL_NumIO, Data.Win.ADODB,System.Types,DateUtils,
  untDBGridFilter;


type
  TFrmPlataforma = class(TForm)
    ActionManager1: TActionManager;
    actPanTo: TAction;
    actImprimir: TAction;
    actSalvar: TAction;
    PanelTitulo: TPanel;
    actAtualizarTituo: TAction;
    actMarcar: TAction;
    actProcurar: TAction;
    SavePictureDialog1: TSavePictureDialog;
    actCalcularXY: TAction;
    actZoomFit: TAction;
    actPan: TAction;
    actZoomDinamico: TAction;
    actZoomWindow: TAction;
    actClearGrafico: TAction;
    PopupMenuMapa: TPopupMenu;
    actMapaMundi: TAction;
    MapaMundi1: TMenuItem;
    actBrasil: TAction;
    Brasil1: TMenuItem;
    actLatitudeLongitude: TAction;
    LatitudeseLongitudes1: TMenuItem;
    actCarregarLinhas: TAction;
    OpenDialog1: TOpenDialog;
    Carregarlinhas2Dx1y1x2y21: TMenuItem;
    actAbrirDesenho: TAction;
    actSalvarDesenhoData: TAction;
    SaveDialog1: TSaveDialog;
    actTabuaMare: TAction;
    actProcurarTabuaMare: TAction;
    StatusBar1: TStatusBar;
    DBGridPlataformas: TFilterDBGrid;
    ToolBar1: TToolBar;
    DBNavigatorCadastro: TDBNavigator;
    ToolButton2: TToolButton;
    BitBtn5: TBitBtn;
    DBEdit1: TDBEdit;
    btnClearFiltro: TToolButton;
    btnExcel: TToolButton;
    btnLayout: TToolButton;
    ColunasLayout: TStringGrid;
    procedure StatusBar1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
      const Rect: TRect);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure DBGridPlataformasDrawColumnCell(Sender: TObject;
      const Rect: TRect; DataCol: Integer; Column: TColumn;
      State: TGridDrawState);
    procedure actAtualizarTituoExecute(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
    procedure actCalcularXYExecute(Sender: TObject);

    procedure PanelTituloAjudaMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);

   
    procedure DBGridPlataformasCellClick(Column: TColumn);

  private
    //procedure DoMap;
    Capturing: Boolean;
    MouseDownSpot: TPoint;
    AlturaAgora: Double;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;

  public

  end;

var
  FrmPlataforma: TFrmPlataforma;

implementation

uses
  untPrincipal,untDataModule,untFrmPreview;

function GetStatusBarPanelXY(StatusBar : TStatusBar; X, Y: Integer) : Integer;
var
  i: Integer;
  R: TRect;
begin
  Result := -1;

  // Buscamos panel a panel hasta encontrar en cual está XY
  with StatusBar do
    for i := 0 to Panels.Count - 1 do
    begin
      // Obtenemos las dimensiones del panel
      SendMessage(Handle, WM_USER + 10, i, Integer(@R));
      if PtInRect(R, Point(x,y)) then
      begin
        Result := i;
        Break;
      end;
    end;
end;

{$R *.dfm}

procedure TFrmPlataforma.DBGridPlataformasCellClick(Column: TColumn);
begin
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT)) then
  begin
    try
      if (Self.DBGridPlataformas.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'booleanPlataforma') then
      begin
        DBGridPlataformas.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOQueryPlataforma.Edit;
        FrmDataModule.DataSourcePlataforma.DataSet.
        FieldByName('booleanPlataforma').AsBoolean:= not Self.DBGridPlataformas.SelectedField.AsBoolean;
        FrmDataModule.ADOQueryPlataforma.Post;
      end
      else if (Self.DBGridPlataformas.SelectedField.DataType = ftBoolean)AND
      (Column.Field.FieldName = 'booleanOrigem') then
      begin
        DBGridPlataformas.Options:=
        [dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
        FrmDataModule.ADOQueryPlataforma.Edit;
        FrmDataModule.DataSourcePlataforma.DataSet.
        FieldByName('booleanOrigem').AsBoolean:= not Self.DBGridPlataformas.SelectedField.AsBoolean;
        FrmDataModule.ADOQueryPlataforma.Post;
      end
      else
        DBGridPlataformas.Options:=
        [dgEditing,dgTitles,dgIndicator,dgColumnResize,dgColLines,dgRowLines,dgTabs,dgAlwaysShowSelection,dgTitleClick,dgTitleHotTrack];
    except
    end;
  end;
end;

procedure TFrmPlataforma.DBGridPlataformasDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn; State: TGridDrawState);
    var
      CheckBoxRectangle : TRect;
    Const
    CtrlState : array[Boolean] of Integer = (DFCS_BUTTONCHECK,
    DFCS_BUTTONCHECK or DFCS_CHECKED);
begin
  FrmPrincipal.GridZebrado(DBGridPlataformas,ColunasLayout,State,Rect,DataCol,Column);
  if ((Column.Field.FieldName = 'booleanPlataforma')OR
  (Column.Field.FieldName = 'booleanOrigem')) then
  begin
    Self.DBGridPlataformas.Canvas.FillRect(Rect);
    CheckBoxRectangle.Left := Rect.Left + 2;
    CheckBoxRectangle.Right := Rect.Right - 2;
    CheckBoxRectangle.Top := Rect.Top + 2;
    CheckBoxRectangle.Bottom := Rect.Bottom - 2;
    DrawFrameControl(Self.DBGridPlataformas.Canvas.Handle,
    CheckBoxRectangle, DFC_BUTTON,
    CtrlState[Column.Field.AsBoolean]);
  end;
end;

procedure TFrmPlataforma.actAtualizarTituoExecute(Sender: TObject);
  var
    plataforma: String;
begin
  //Nome da plataforma e guindaste no titulo de cada aba
  plataforma:=
  FrmDataModule.DataSourcePlataforma.
  DataSet.FieldByName('Plataforma').AsString+
  ' - Lat.: '+ FrmDataModule.DataSourcePlataforma.
  DataSet.FieldByName('Latitude').AsString+
  '; Lgn.: '+ FrmDataModule.DataSourcePlataforma.
  DataSet.FieldByName('Longitude').AsString;

  PanelTitulo.Caption:=
  'Cadastro de Instalações - '+plataforma;
end;

procedure TFrmPlataforma.actCalcularXYExecute(Sender: TObject);
  var
    Latitude,Longitude: Double;
    NoOrigem,NODestino: TPointFloat;
begin
  FrmDataModule.ADOQueryPlataforma.Active:= false;
  FrmDataModule.ADOQueryPlataforma.Active:= true;
  NoOrigem.X:= 0;
  NoOrigem.Y:= 0;
  FrmDataModule.ADOQueryPlataforma.Active:= false;
  FrmDataModule.ADOQueryPlataforma.Active:= true;
  //======================================================
  FrmPrincipal.ProgressBarIncializa(FrmDataModule.ADOQueryPlataforma.RecordCount,'Calculando coordenadas (X,Y)');
  FrmDataModule.DataSourcePlataforma.Enabled:= false;
  FrmDataModule.ADOQueryPlataforma.First;
  while not(FrmDataModule.ADOQueryPlataforma.Eof) do
  begin
    Latitude:= FrmDataModule.DataSourcePlataforma.
    DataSet.FieldByName('Latitude').AsFloat;
    Longitude:= FrmDataModule.DataSourcePlataforma.
    DataSet.FieldByName('Longitude').AsFloat;
    NODestino:= FrmPrincipal.LatLong_XY(Latitude,Longitude);
    FrmDataModule.ADOQueryPlataforma.Edit;
    FrmDataModule.DataSourcePlataforma.
    DataSet.FieldByName('CoordX').AsFloat:= RoundTo(NODestino.X,-4);
    FrmDataModule.DataSourcePlataforma.
    DataSet.FieldByName('CoordY').AsFloat:= RoundTo(NODestino.Y,-4);
    FrmDataModule.ADOQueryPlataforma.Post;
    FrmDataModule.ADOQueryPlataforma.Next;
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  FrmDataModule.ADOQueryPlataforma.First;
  FrmDataModule.DataSourcePlataforma.Enabled:= true;
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmPlataforma.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(ColunasLayout,true);
  SQLBase:= 'SELECT tblPlataforma.* FROM tblPlataforma '+
  SQLString+' ORDER BY Plataforma;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryPlataforma,StatusBar1);
end;

procedure TFrmPlataforma.PanelTituloAjudaMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  Capturing := true;
  MouseDownSpot.X := X;
  MouseDownSpot.Y := Y;
end;

procedure TFrmPlataforma.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
  FrmPlataforma:=nil;
end;

procedure TFrmPlataforma.FormCreate(Sender: TObject);
begin
  //===============================================================
  AlturaAgora:=0;
  //=====================================
  //Incicialização
  FrmDataModule.setFilterDBGrid(DBGridPlataformas);

  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);
  //Configuração de Perfil
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT)OR
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO)) then
  begin
    DBNavigatorCadastro.Enabled:= true;
    DBGridPlataformas.ReadOnly:= false;
    actCalcularXY.Enabled:= true;
  end
  else
  begin
    DBNavigatorCadastro.Enabled:= false;
    DBGridPlataformas.ReadOnly:= true;
    actCalcularXY.Enabled:= false;
  end;

  actProcurar.Execute;
end;

procedure TFrmPlataforma.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmPlataforma.StatusBar1DrawPanel(StatusBar: TStatusBar; Panel: TStatusPanel;
  const Rect: TRect);
begin
  StatusBar.Canvas.Font.Color := clBlue;
  Statusbar.Canvas.TextRect(Rect, Rect.Left, Rect.Top, Panel.Text);
end;

procedure TFrmPlataforma.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
