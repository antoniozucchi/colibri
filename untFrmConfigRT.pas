unit untFrmConfigRT;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ToolWin, Vcl.ComCtrls, Vcl.ExtCtrls,
  Vcl.DBCtrls, Vcl.StdCtrls, Vcl.Buttons, Vcl.Mask, Vcl.WinXPickers,
  DBDateTimePicker, Data.DB, Vcl.Grids, Vcl.DBGrids, untDBGridFilter,
  System.Actions, Vcl.ActnList, Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan;

type
  TFrmConfigRT = class(TForm)
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    GridPanel2: TGridPanel;
    Label3: TLabel;
    Label4: TLabel;
    DBEdit3: TDBEdit;
    DBEdit4: TDBEdit;
    Label5: TLabel;
    DBEdit5: TDBEdit;
    GridPanel1: TGridPanel;
    Label6: TLabel;
    DBEdit6: TDBEdit;
    Label8: TLabel;
    DBEdit10: TDBEdit;
    Label11: TLabel;
    DBEdit11: TDBEdit;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    DBDateTimePicker2: TDBDateTimePicker;
    DBEdit12: TDBEdit;
    DBEdit13: TDBEdit;
    DBEdit1: TDBEdit;
    Label1: TLabel;
    DBCheckBox1: TDBCheckBox;
    ToolBar1: TToolBar;
    DBNavigator1: TDBNavigator;
    PanelTitulo: TPanel;
    ToolBar3: TToolBar;
    btnClearFiltroRT: TToolButton;
    btnExcelRT: TToolButton;
    btnLaypoutRT: TToolButton;
    FilterDBGrid1: TFilterDBGrid;
    StatusBarRegrasRecolhimento: TStatusBar;
    RLLayout: TStringGrid;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    procedure FormCreate(Sender: TObject);
    procedure actProcurarExecute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmConfigRT: TFrmConfigRT;

implementation
  uses untDataModule, untPrincipal;

{$R *.dfm}

procedure TFrmConfigRT.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= frmPrincipal.SQLStringFiltroTabela(RLLayout,true);
  SQLBase:= 'SELECT tblRTRegraRecolhimento.* FROM tblRTRegraRecolhimento '+
  SQLString+' ORDER BY Origem;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryRegrasRecolhimento,
    StatusBarRegrasRecolhimento);
end;

procedure TFrmConfigRT.FormCreate(Sender: TObject);
begin
  FrmDataModule.ADOQueryConfigRT.Active:= false;
  FrmDataModule.ADOQueryConfigRT.Active:= true;
  FrmDataModule.setFilterDBGrid(FilterDBGrid1);
  actProcurar.Execute;
end;

end.
