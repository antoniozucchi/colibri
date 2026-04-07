unit untParesObrigatorios;

{
  Form de cadastro de pares obrigatórios de distribuição.

  Um par obrigatório define que duas plataformas DEVEM ser atendidas pela
  mesma embarcação em uma operação de distribuição automática.
  Ex: PCM-02 + PCM-03 (M2 e M3, muito próximas — sempre no mesmo barco)

  O campo Ativo permite desativar temporariamente um par sem excluí-lo.
}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Buttons, DB, Data.Win.ADODB, DBCtrls, Grids, DBGrids,
  Vcl.ExtCtrls, Vcl.ComCtrls, Vcl.ToolWin, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan;

type
  TFrmParesObrigatorios = class(TForm)
    PanelTitulo: TPanel;
    PanelBotoes: TPanel;
    BtnFechar: TBitBtn;
    BtnAdicionar: TBitBtn;
    BtnExcluir: TBitBtn;
    BtnSalvar: TBitBtn;
    BtnCancelar: TBitBtn;
    DBGridPares: TDBGrid;
    LabelInfo: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BtnFecharClick(Sender: TObject);
    procedure BtnAdicionarClick(Sender: TObject);
    procedure BtnExcluirClick(Sender: TObject);
    procedure BtnSalvarClick(Sender: TObject);
    procedure BtnCancelarClick(Sender: TObject);
    procedure DBGridParesDrawColumnCell(Sender: TObject; const Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGridParesCellClick(Column: TColumn);
  private
    FQuery: TADOQuery;
    FDataSource: TDataSource;
    FConnConsulta: TADOConnection;
    procedure CarregarDados;
  public
    // Deve ser chamado antes de ShowModal — passa a conexão dbConsulta
    procedure Inicializar(AConnConsulta: TADOConnection);
  end;

implementation

{$R *.dfm}

procedure TFrmParesObrigatorios.FormCreate(Sender: TObject);
begin
  FQuery      := TADOQuery.Create(Self);
  FDataSource := TDataSource.Create(Self);
  FDataSource.DataSet := FQuery;
  DBGridPares.DataSource := FDataSource;
end;

procedure TFrmParesObrigatorios.Inicializar(AConnConsulta: TADOConnection);
begin
  FConnConsulta := AConnConsulta;
  CarregarDados;
end;

procedure TFrmParesObrigatorios.CarregarDados;
begin
  if not Assigned(FConnConsulta) then Exit;

  FQuery.Connection := FConnConsulta;
  FQuery.Close;
  FQuery.SQL.Text :=
    'SELECT idParObrigatorio, PlataformaA, PlataformaB, Ativo ' +
    'FROM tblParObrigatorio ' +
    'ORDER BY PlataformaA, PlataformaB';
  try
    FQuery.Open;
  except
    on E: Exception do
      ShowMessage('Erro ao carregar pares: ' + E.Message);
  end;
end;

procedure TFrmParesObrigatorios.FormDestroy(Sender: TObject);
begin
  if FQuery.State in [dsEdit, dsInsert] then
    FQuery.Cancel;
end;

procedure TFrmParesObrigatorios.BtnFecharClick(Sender: TObject);
begin
  if FQuery.State in [dsEdit, dsInsert] then
  begin
    if Application.MessageBox(
         'Há alterações não salvas. Deseja descartar?',
         'Confirmar', MB_YESNO + MB_ICONQUESTION) = IDNO then
      Exit;
    FQuery.Cancel;
  end;
  Close;
end;

procedure TFrmParesObrigatorios.BtnAdicionarClick(Sender: TObject);
begin
  FQuery.Append;
  FQuery.FieldByName('Ativo').AsBoolean := True;
  DBGridPares.SetFocus;
end;

procedure TFrmParesObrigatorios.BtnExcluirClick(Sender: TObject);
begin
  if FQuery.IsEmpty or (FQuery.State in [dsInsert]) then Exit;

  if Application.MessageBox(
       PChar('Excluir o par ' +
         FQuery.FieldByName('PlataformaA').AsString + ' + ' +
         FQuery.FieldByName('PlataformaB').AsString + '?'),
       'Confirmar exclusão',
       MB_YESNO + MB_ICONQUESTION) = IDYES then
  begin
    try
      FQuery.Delete;
    except
      on E: Exception do
        ShowMessage('Erro ao excluir: ' + E.Message);
    end;
  end;
end;

procedure TFrmParesObrigatorios.BtnSalvarClick(Sender: TObject);
begin
  if FQuery.State in [dsEdit, dsInsert] then
  begin
    // Validação mínima
    if Trim(FQuery.FieldByName('PlataformaA').AsString) = '' then
    begin
      ShowMessage('Informe a Plataforma A.');
      Exit;
    end;
    if Trim(FQuery.FieldByName('PlataformaB').AsString) = '' then
    begin
      ShowMessage('Informe a Plataforma B.');
      Exit;
    end;
    if FQuery.FieldByName('PlataformaA').AsString =
       FQuery.FieldByName('PlataformaB').AsString then
    begin
      ShowMessage('Plataforma A e Plataforma B devem ser diferentes.');
      Exit;
    end;
    try
      FQuery.Post;
    except
      on E: Exception do
        ShowMessage('Erro ao salvar: ' + E.Message);
    end;
  end;
end;

procedure TFrmParesObrigatorios.BtnCancelarClick(Sender: TObject);
begin
  if FQuery.State in [dsEdit, dsInsert] then
    FQuery.Cancel;
end;

procedure TFrmParesObrigatorios.DBGridParesDrawColumnCell(Sender: TObject;
  const Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
var
  CheckBoxRect: TRect;
const
  CtrlState: array[Boolean] of Integer =
    (DFCS_BUTTONCHECK, DFCS_BUTTONCHECK or DFCS_CHECKED);
begin
  if Column.Field.FieldName = 'Ativo' then
  begin
    DBGridPares.Canvas.FillRect(Rect);
    CheckBoxRect.Left   := Rect.Left   + 2;
    CheckBoxRect.Right  := Rect.Right  - 2;
    CheckBoxRect.Top    := Rect.Top    + 2;
    CheckBoxRect.Bottom := Rect.Bottom - 2;
    DrawFrameControl(DBGridPares.Canvas.Handle, CheckBoxRect,
      DFC_BUTTON, CtrlState[Column.Field.AsBoolean]);
  end;
end;

procedure TFrmParesObrigatorios.DBGridParesCellClick(Column: TColumn);
begin
  if (Column.Field.FieldName = 'Ativo') and
     (Column.Field.DataType = ftBoolean) then
  begin
    if not (FQuery.State in [dsEdit, dsInsert]) then
      FQuery.Edit;
    FQuery.FieldByName('Ativo').AsBoolean :=
      not FQuery.FieldByName('Ativo').AsBoolean;
    FQuery.Post;
  end;
end;

end.
