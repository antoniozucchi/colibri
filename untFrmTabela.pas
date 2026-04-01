unit untFrmTabela;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ToolWin,
  Vcl.ComCtrls, Vcl.Grids, Vcl.ExtCtrls, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan, uZucchi;

type
  TFrmTabela = class(TForm)
    RLTabela: TStringGrid;
    ToolBar6: TToolBar;
    BitBtn13: TBitBtn;
    PanelTitulo: TPanel;
    ActionManager1: TActionManager;
    actExcel: TAction;
    procedure actExcelExecute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmTabela: TFrmTabela;

implementation
  uses untPrincipal;

{$R *.dfm}

procedure TFrmTabela.actExcelExecute(Sender: TObject);
begin
  ExcelStringGrid(RLTabela,FrmTabela.Caption,PanelTitulo.Caption,
    '',
    '',false,
    FrmPrincipal.ProgressBarPrincipal,FrmPrincipal.MemoPrincipal);
end;

end.
