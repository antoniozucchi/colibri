unit untTelaAbertura;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Imaging.pngimage,
  Vcl.StdCtrls, SDL_ProgBar;

type
  TFrmTelaAbertura = class(TForm)
    Image2: TImage;
    MemoMensagem: TMemo;
    ProgressBarTelaAbertura: TProgBar;
    lblColibri: TLabel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FrmTelaAbertura: TFrmTelaAbertura;

implementation

{$R *.dfm}

end.
