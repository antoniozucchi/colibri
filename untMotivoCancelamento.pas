unit untMotivoCancelamento;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ComCtrls, Vcl.StdCtrls,
  Vcl.Buttons, Vcl.DBCtrls, Vcl.ToolWin, Vcl.Grids, Vcl.DBGrids,
  untDBGridFilter, Vcl.ExtCtrls, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan;

type
  TFrmMotivoCancelamento = class(TForm)
    PanelCancelamentoMudanca: TPanel;
    PanelMotivoCancelamento: TPanel;
    DBGridPalavraChave: TFilterDBGrid;
    ToolBar9: TToolBar;
    DBNavigatorPalavraChave: TDBNavigator;
    BitBtn15: TBitBtn;
    BitBtn14: TBitBtn;
    StatusBarPalavraChave: TStatusBar;
    btnClearFiltro: TToolButton;
    btnExcel: TToolButton;
    ActionManager1: TActionManager;
    actProcurar: TAction;
    actInserir: TAction;
    actCancelar: TAction;
    RadioGroupFonteCancelamento: TRadioGroup;
    PanelTitulo: TPanel;
    procedure actProcurarExecute(Sender: TObject);
    procedure actInserirExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure actCancelarExecute(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
    procedure Iniciar(Fonte: Integer);
  end;

var
  FrmMotivoCancelamento: TFrmMotivoCancelamento;

implementation
  uses untPrincipal, untDataModule, untGerenciarSolicitacoes;

{$R *.dfm}

procedure TFrmMotivoCancelamento.actCancelarExecute(Sender: TObject);
begin
  FrmMotivoCancelamento.Close;
end;

procedure TFrmMotivoCancelamento.actInserirExecute(Sender: TObject);
  var
    PalavraChave: String;
    idProgramacaoExecutante,idProgramacaoDiaria,NumRegistros: Integer;
    passou: Boolean;
begin
  passou:= true;
  PalavraChave:= '';
  while not FrmDataModule.ADOQueryPalavraChave.Eof do
  begin
    if FrmDataModule.ADOQueryPalavraChave.FieldByName('booleanSelecao').AsBoolean then
    begin
      if PalavraChave <> '' then
        PalavraChave:= PalavraChave+'; '+FrmDataModule.ADOQueryPalavraChave.FieldByName('PalavraChave').AsString
      else
        PalavraChave:= FrmDataModule.ADOQueryPalavraChave.FieldByName('PalavraChave').AsString;
    end;
    FrmDataModule.ADOQueryPalavraChave.Next;
  end;
  //================================================================
  if PalavraChave = '' then
  begin
    passou:= false;
    MessageBox(0,'O campo "Palavra Chave" n'#227'o pode estar em branco!','Avalia'#231#227'o de Programa'#231#227'o',MB_ICONERROR);
  end;
  if passou then
  begin
    case RadioGroupFonteCancelamento.ItemIndex of
      0://Gerenciar Solicitacoes (Cancelamento)
      begin
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
        NumRegistros:= FrmDataModule.ADOQueryGerenciarExecutante.RecordCount;
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Atribuindo status para programa'#231#227'o...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('idProgramacaoExecutante').AsInteger;

            FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,0,
            'Cancelado',PalavraChave);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        //Calcular Cancelados
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Calculando: Aprovados, Cancelados e N'#176' de Executantes...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger;
            FrmPrincipal.GravarCanceladoAprovado(idProgramacaoDiaria);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
        FrmGerenciarSolicitacoes.actMatrizOrigemDestino.Execute;
        FrmGerenciarSolicitacoes.StatusLinhaSelecionada;
        FrmGerenciarSolicitacoes.actContadorSolicitacao.Execute;
      end;
      1: //Marcar os nao executados
      begin
        FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= false;
        NumRegistros:= FrmDataModule.ADOQueryConsultaExecutantesProgramados.RecordCount;
        FrmDataModule.ADOQueryConsultaExecutantesProgramados.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Atribuindo Status de Execu'#231#227'o...');
        while not FrmDataModule.ADOQueryConsultaExecutantesProgramados.Eof do
        begin
          if FrmDataModule.DataSourceConsultaExecutantesProgramados.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoExecutante:= FrmDataModule.DataSourceConsultaExecutantesProgramados.
            DataSet.FieldByName('idProgramacaoExecutante').AsInteger;
            FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,1,
            '',PalavraChave);
          end;
          FrmDataModule.ADOQueryConsultaExecutantesProgramados.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        //==========================================================
        FrmPrincipal.ProgressBarAtualizar;
        FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= false;
        FrmDataModule.ADOQueryConsultaExecutantesProgramados.Active:= true;
        FrmDataModule.DataSourceConsultaExecutantesProgramados.Enabled:= true;
      end;
      2: //Gerenciamento de Solicitacoes (Mudanca)
      begin
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= false;
        NumRegistros:= FrmDataModule.ADOQueryGerenciarExecutante.RecordCount;
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Atribuindo status para programa'#231#227'o...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoExecutante:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('idProgramacaoExecutante').AsInteger;

            FrmPrincipal.AvaliarProgramacaoExecutante(idProgramacaoExecutante,0,
            'Mudan'#231'a',PalavraChave);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        //Calcular Cancelados
        FrmDataModule.ADOQueryGerenciarExecutante.First;
        FrmPrincipal.ProgressBarIncializa(NumRegistros,
        'Calculando: Aprovados, Cancelados e N'#176' de Executantes...');
        while not FrmDataModule.ADOQueryGerenciarExecutante.Eof do
        begin
          if FrmDataModule.DataSourceGerenciarExecutante.DataSet.
          FieldByName('booleanSelecao').AsBoolean then
          begin
            idProgramacaoDiaria:= FrmDataModule.DataSourceGerenciarExecutante.
            DataSet.FieldByName('CodigoProgramacaoDiaria').AsInteger;
            FrmPrincipal.GravarCanceladoAprovado(idProgramacaoDiaria);
          end;
          FrmDataModule.ADOQueryGerenciarExecutante.Next;
          FrmPrincipal.ProgressBarIncremento(1);
        end;
        FrmPrincipal.ProgressBarAtualizar;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= false;
        FrmDataModule.ADOQueryGerenciarExecutante.Active:= true;
        FrmDataModule.DataSourceGerenciarExecutante.Enabled:= true;
        FrmGerenciarSolicitacoes.actMatrizOrigemDestino.Execute;
        FrmGerenciarSolicitacoes.StatusLinhaSelecionada;
        FrmGerenciarSolicitacoes.actContadorSolicitacao.Execute;
      end;
    end;
  end;
  FrmMotivoCancelamento.Close;
end;

procedure TFrmMotivoCancelamento.actProcurarExecute(Sender: TObject);
  var
    SQLString,SQLBase: String;
begin
  SQLString:= BuildFilterSQL(DBGridPalavraChave,true);
  SQLBase:= 'SELECT tblPalavraChave.* FROM tblPalavraChave'+
  SQLString+' ORDER BY PalavraChave;';
  FrmPrincipal.ProcuraQuery(SQLBase,FrmDataModule.ADOQueryPalavraChave,StatusBarPalavraChave);
end;

procedure TFrmMotivoCancelamento.FormCreate(Sender: TObject);
begin
  actProcurar.Execute;

  if ((FrmPrincipal.logPerfil = 'Administrador')OR
  (FrmPrincipal.logPerfil = 'Programa'#231#227'o')OR
  (FrmPrincipal.logPerfil = 'Supervisor')) then
  begin
    DBNavigatorPalavraChave.Enabled:= true;
    DBGridPalavraChave.ReadOnly:= false;
    actInserir.Enabled:= true;
    actCancelar.Enabled:= true;
  end
  else
  begin
    DBNavigatorPalavraChave.Enabled:= false;
    DBGridPalavraChave.ReadOnly:= true;
    actInserir.Enabled:= false;
    actCancelar.Enabled:= false;
  end;
end;

procedure TFrmMotivoCancelamento.Iniciar(Fonte: Integer);
begin
  RadioGroupFonteCancelamento.ItemIndex:= Fonte;
  case Fonte of
    0: PanelTitulo.Caption:= 'Motivo de Cancelamento e Palavras Chave';
    1: PanelTitulo.Caption:= 'Motivo da N'#227'o Execu'#231#227'o';
    2: PanelTitulo.Caption:= 'Motivo da Mudan'#231'a e Palavras Chave';
  end;
end;

end.


