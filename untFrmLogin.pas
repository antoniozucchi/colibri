unit untFrmLogin;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, ActnList, PlatformDefaultStyleActnCtrls, ActnMan,
  ExtCtrls, pngimage, System.Actions, Vcl.Imaging.jpeg;
  {Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, ExtCtrls, pngimage, Buttons, ActnList,
  PlatformDefaultStyleActnCtrls, ActnMan, Mask;}

type
  TFrmLogin = class(TForm)
    Panel1: TPanel;
    BitBtnLogin: TBitBtn;
    BitBtnCancel: TBitBtn;
    edtUsuario: TEdit;
    edtSenha: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    ActionManager1: TActionManager;
    actLogin: TAction;
    procedure FormCreate(Sender: TObject);
    procedure BitBtnCancelClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure actLoginExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
  private
    { Private declarations }
    Function usuarioWindows: string;
    procedure WMMDIACTIVATE(var msg: TWMMDIACTIVATE);message WM_MDIACTIVATE;
  public
    { Public declarations }
    fecha : boolean; //DECLARAÇÃO DA VARIÁVEL FECHA
  end;

var
  FrmLogin: TFrmLogin;
  contaErro: Integer;

implementation
uses untPrincipal;

{$R *.dfm}

function ValidateUserLogonAPI(const UserName: string; const Domain: string;
  const PassWord: string) : boolean;
var
  Retvar: boolean;
  LHandle: THandle;
begin
  Retvar := LogonUser(PWideChar(UserName),
                                PWideChar(Domain),
                                PWideChar(PassWord),
                                LOGON32_LOGON_NETWORK,
                                LOGON32_PROVIDER_DEFAULT,
                                LHandle);

  if Retvar then
    CloseHandle(LHandle);

  Result := Retvar;
end;

procedure TFrmLogin.actLoginExecute(Sender: TObject);
var
  getUsuario, getDominio, getSenha: String;
begin
  try
    getUsuario:= UpperCase(edtUsuario.Text);
    getSenha:= edtSenha.Text;

    if(contaErro<3) then
    begin
      if ValidateUserLogonAPI(getUsuario,getDominio,getSenha)=true then
      begin
          Fecha := True;
          FrmLogin.Close;
          FrmPrincipal.Show;
          FrmPrincipal.carregarLoginUsuario(edtUsuario.Text);
      end
      else
      begin
        MessageBox(FrmLogin.Handle,'Senha ou nome de usuário invalidos!',
        'Login',MB_ICONERROR);
        contaErro:=contaErro+1;
      end;
    end
    else
    begin
      MessageBox(FrmLogin.Handle,'Você excedeu o limite máximo de tentativas de acesso',
      'Login',MB_ICONERROR);
      Fecha := True;
      FrmPrincipal.Close;//Isso vai terminar o programa
    end;
  except
    Application.Terminate;
  end;
end;

procedure TFrmLogin.BitBtnCancelClick(Sender: TObject);
begin
  Fecha := True;
  FrmLogin.Close;//Isso vai terminar o programa
end;

procedure TFrmLogin.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action:= caFree;
  FrmLogin:=nil;
end;

procedure TFrmLogin.FormCreate(Sender: TObject);
begin
  //======ADICIONAR TABSET DO FOMRMDI=======
  FrmPrincipal.MDIChildCreated(self.Handle);

  Fecha := false;
  contaErro := 0;
  edtUsuario.Text:= usuarioWindows;
end;

procedure TFrmLogin.FormDestroy(Sender: TObject);
begin
  FrmPrincipal.MDIChildDestroyed(self.Handle);
end;

procedure TFrmLogin.FormShow(Sender: TObject);
begin
  edtSenha.SetFocus;
end;

function TFrmLogin.usuarioWindows: string;
Var
  NetUserNameLength: DWord;
Begin
  NetUserNameLength:=50;
  SetLength(Result, NetUserNameLength);
  GetUserName(pChar(Result),NetUserNameLength);
  SetLength(Result, StrLen(pChar(Result)));
End;

procedure TFrmLogin.WMMDIACTIVATE(var msg: TWMMDIACTIVATE);
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
