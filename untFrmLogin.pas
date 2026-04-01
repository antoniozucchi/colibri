unit untFrmLogin;

interface

uses
  Winapi.Windows, Winapi.Messages,
  System.SysUtils, System.Variants, System.Classes,
  Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.StdCtrls, Vcl.Buttons, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan,
  Vcl.ExtCtrls, Vcl.Imaging.pngimage, System.Actions,
  Vcl.Imaging.jpeg, System.UITypes;

type
  TFrmLogin = class(TForm)
    ActionManager1: TActionManager;
    actLogin: TAction;
    GridPanel1: TGridPanel;
    Label1: TLabel;
    edtUsuario: TEdit;
    Label2: TLabel;
    edtSenha: TEdit;
    GridPanel2: TGridPanel;
    BitBtnLogin: TBitBtn;
    BitBtnCancel: TBitBtn;
    procedure FormCreate(Sender: TObject);
    procedure BitBtnCancelClick(Sender: TObject);
    procedure actLoginExecute(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FContaErro: Integer;
    function UsuarioWindows: string;
    procedure SepararUsuarioEDominio(const ALogin: string; out AUsuario, ADominio: string);
    procedure FinalizarAplicacao;
  public
    Fecha: Boolean;
  end;

var
  FrmLogin: TFrmLogin;

implementation

uses
  untPrincipal;

{$R *.dfm}

const
  MAX_TENTATIVAS_LOGIN = 3;

function ValidateUserLogonAPI(const UserName, Domain, Password: string): Boolean;
var
  LHandle: THandle;
  PDomain: PChar;
begin
  LHandle := 0;

  if Trim(Domain) <> '' then
    PDomain := PChar(Domain)
  else
    PDomain := nil;

  Result := LogonUser(
    PChar(UserName),
    PDomain,
    PChar(Password),
    LOGON32_LOGON_NETWORK,
    LOGON32_PROVIDER_DEFAULT,
    LHandle
  );

  if Result and (LHandle <> 0) then
    CloseHandle(LHandle);
end;

procedure TFrmLogin.SepararUsuarioEDominio(const ALogin: string; out AUsuario, ADominio: string);
var
  P: Integer;
  S: string;
begin
  S := Trim(ALogin);
  AUsuario := S;
  ADominio := '';

  // Formato: DOMINIO\usuario
  P := Pos('\', S);
  if P > 0 then
  begin
    ADominio := Copy(S, 1, P - 1);
    AUsuario := Copy(S, P + 1, MaxInt);
    Exit;
  end;

  // Formato UPN: usuario@dominio
  if Pos('@', S) > 0 then
  begin
    AUsuario := S;
    ADominio := '';
    Exit;
  end;

  // Usa o domínio atual do Windows
  ADominio := Trim(GetEnvironmentVariable('USERDOMAIN'));

  // Fallback para máquina local
  if ADominio = '' then
    ADominio := '.';
end;

procedure TFrmLogin.FinalizarAplicacao;
begin
  Fecha := True;
  Close;
  Application.Terminate;
end;

procedure TFrmLogin.actLoginExecute(Sender: TObject);
var
  LUsuario, LDominio, LSenha: string;
  LTentativasRestantes: Integer;
begin
  LUsuario := Trim(edtUsuario.Text);
  LSenha := edtSenha.Text;

  if LUsuario = '' then
  begin
    MessageDlg('Informe o usuário.', mtWarning, [mbOK], 0);
    edtUsuario.SetFocus;
    Exit;
  end;

  if LSenha = '' then
  begin
    MessageDlg('Informe a senha.', mtWarning, [mbOK], 0);
    edtSenha.SetFocus;
    Exit;
  end;

  SepararUsuarioEDominio(LUsuario, LUsuario, LDominio);

  try
    if ValidateUserLogonAPI(LUsuario, LDominio, LSenha) then
    begin
      Fecha := True;
      FrmPrincipal.carregarLoginUsuario(Trim(edtUsuario.Text));
      FrmPrincipal.Show;
      Close;
      Exit;
    end;

    Inc(FContaErro);

    if FContaErro >= MAX_TENTATIVAS_LOGIN then
    begin
      MessageDlg(
        'Você excedeu o limite máximo de tentativas de acesso.',
        mtError, [mbOK], 0
      );
      FinalizarAplicacao;
      Exit;
    end;

    LTentativasRestantes := MAX_TENTATIVAS_LOGIN - FContaErro;

    MessageDlg(
      Format('Usuário ou senha inválidos.%sTentativas restantes: %d',
        [sLineBreak, LTentativasRestantes]),
      mtError, [mbOK], 0
    );

    edtSenha.Clear;
    edtSenha.SetFocus;

  except
    on E: Exception do
    begin
      MessageDlg(
        'Erro ao validar login:' + sLineBreak + E.Message,
        mtError, [mbOK], 0
      );
      edtSenha.Clear;
      edtSenha.SetFocus;
    end;
  end;
end;

procedure TFrmLogin.BitBtnCancelClick(Sender: TObject);
begin
  FinalizarAplicacao;
end;

procedure TFrmLogin.FormCreate(Sender: TObject);
begin
  Fecha := False;
  FContaErro := 0;
  edtUsuario.Text := UsuarioWindows;
  edtSenha.PasswordChar := '*';
end;

procedure TFrmLogin.FormShow(Sender: TObject);
begin
  if Trim(edtUsuario.Text) = '' then
    edtUsuario.SetFocus
  else
    edtSenha.SetFocus;
end;

function TFrmLogin.UsuarioWindows: string;
var
  Buffer: array[0..255] of Char;
  BufferSize: DWORD;
begin
  BufferSize := Length(Buffer);
  if GetUserName(Buffer, BufferSize) then
    Result := Trim(Buffer)
  else
    Result := '';
end;

end.
