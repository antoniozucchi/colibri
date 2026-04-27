unit untConexaoLOCAL;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Buttons, Vcl.ComCtrls,
  Vcl.ToolWin, Vcl.ExtCtrls, System.Actions, Vcl.ActnList,
  Vcl.PlatformDefaultStyleActnCtrls, Vcl.ActnMan;

type
  TFrmConexaoLOCAL = class(TForm)
    ActionManager1: TActionManager;
    actLOCAL: TAction;
    actREDE: TAction;
    actUpload: TAction;
    actDownload: TAction;
    actSelUpload: TAction;
    actSelDownload: TAction;
    PanelTitulo: TPanel;
    GridPanel1: TGridPanel;
    Label1: TLabel;
    Label4: TLabel;
    chkUPColibri: TCheckBox;
    Label5: TLabel;
    Label7: TLabel;
    chkDownColibri: TCheckBox;
    chkUPConsulta: TCheckBox;
    chkDownConsulta: TCheckBox;
    chkUPRT: TCheckBox;
    chkDownRT: TCheckBox;
    ToolBar7: TToolBar;
    BitBtn8: TBitBtn;
    BitBtn10: TBitBtn;
    BitBtn9: TBitBtn;
    BitBtn7: TBitBtn;
    GridPanel2: TGridPanel;
    Label8: TLabel;
    edtRede: TEdit;
    Label9: TLabel;
    edtLocal: TEdit;
    BitBtn1: TBitBtn;
    BitBtn2: TBitBtn;
    procedure actLOCALExecute(Sender: TObject);
    procedure actUploadExecute(Sender: TObject);
    procedure actREDEExecute(Sender: TObject);
    procedure actSelUploadExecute(Sender: TObject);
    procedure actSelDownloadExecute(Sender: TObject);
    procedure actDownloadExecute(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    procedure HabilitaChkBox(chk: TCheckBox; sel: Boolean);
    { Private declarations }
  public
    { Public declarations }
    enderecoREDE,enderecoLOCAL: String;
  end;

var
  FrmConexaoLOCAL: TFrmConexaoLOCAL;

implementation
  uses untPrincipal, untDataModule;

{$R *.dfm}

procedure TFrmConexaoLOCAL.actDownloadExecute(Sender: TObject);
  var
    FileName,FilePath,enderecoREDE,
    enderecoREDEdbConsulta,enderecoREDEdbRT,
    enderecoLOCALdbColibri,enderecoLOCALdbConsulta,enderecoLOCALdbRT: String;
begin
  enderecoREDE:= FrmPrincipal.registroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  //LOCAL
  enderecoLOCALdbColibri:= ExtractFilePath(Application.ExeName)+'LOCAL\'+FileName;
  enderecoLOCALdbConsulta:= ExtractFilePath(Application.ExeName)+'LOCAL\dbConsulta.mdb';
  enderecoLOCALdbRT:= ExtractFilePath(Application.ExeName)+'LOCAL\dbRT.mdb';
  //REDE
  enderecoREDEdbConsulta:= FilePath+'dbConsulta.mdb';
  enderecoREDEdbRT:= FilePath+'dbRT.mdb';
  //DOWNLOAD
  FrmPrincipal.ProgressBarIncializa(3,'Download banco de dados...');
  //=============================================================================
  if chkDownColibri.Checked then
  begin
    CopyFile(PChar(enderecoREDE), PChar(enderecoLOCALdbColibri), false);
    FrmPrincipal.conectarBDDireto(enderecoLOCALdbColibri,FrmDataModule.ADOConnectionColibri);
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  if chkDownConsulta.Checked then
  begin
    CopyFile(PChar(enderecoREDEdbConsulta), PChar(enderecoLOCALdbConsulta), false);
    FrmPrincipal.conectarBDDireto(enderecoLOCALdbConsulta,FrmDataModule.ADOConnectionConsulta);
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  if chkDownRT.Checked then
  begin
    CopyFile(PChar(enderecoREDEdbRT), PChar(enderecoLOCALdbRT), false);
    FrmPrincipal.conectarBDDireto(enderecoLOCALdbRT,FrmDataModule.ADOConnectionRT);
  end;
  FrmPrincipal.ProgressBarIncremento(1);
  //=============================================================================
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmConexaoLOCAL.actLOCALExecute(Sender: TObject);
  var
    enderecoLOCAL,
    enderecoREDE,FileName,FilePath,dbConsulta,dbRT,Substituido: String;
begin
  enderecoREDE:= FrmPrincipal.RegistroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  enderecoLOCAL:= ExtractFilePath(Application.ExeName);
  if Application.MessageBox(PChar(
  'Deseja substituir os bancos de dados LOCAIS atuais pelos bancos de dados da REDE selecionados na coluna DOWNLOAD?'),
  '.::ATENCAO::.',36) = 6 then
  begin
    //Copiar e Substituir Arquivos
    dbConsulta:= FilePath+'\dbConsulta.mdb';
    dbRT:= FilePath+'\dbRT.mdb';
    FrmPrincipal.ProgressBarIncializa(6,'Conexao LOCAL...');
    Substituido:= '';
    //----------------------------------------------------
    if chkDownColibri.Checked then
    begin
      CopyFile(PChar(enderecoREDE), PChar(enderecoLOCAL+'\LOCAL\'+FileName), false);
      Substituido:= ' - dbColibri';
    end;
    FrmPrincipal.ProgressBarIncremento(1);
    //----------------------------------------------------
    if chkDownConsulta.Checked then
    begin
      CopyFile(PChar(dbConsulta), PChar(enderecoLOCAL+'\LOCAL\dbConsulta.mdb'), false);
      Substituido:= Substituido + sLineBreak + ' - dbConsulta';
    end;
    FrmPrincipal.ProgressBarIncremento(1);
    //----------------------------------------------------
    if chkDownRT.Checked then
    begin
      CopyFile(PChar(dbRT), PChar(enderecoLOCAL+'\LOCAL\dbRT.mdb'), false);
      Substituido:= Substituido + sLineBreak + ' - dbRT';
    end;
    FrmPrincipal.ProgressBarIncremento(1);
    //Conexao
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\'+FileName,FrmDataModule.ADOConnectionColibri);
    FrmPrincipal.ProgressBarIncremento(1);
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\dbConsulta.mdb',FrmDataModule.ADOConnectionConsulta);
    FrmPrincipal.ProgressBarIncremento(1);
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\dbRT.mdb',FrmDataModule.ADOConnectionRT);
    FrmPrincipal.ProgressBarIncremento(1);
    if Substituido <> '' then
      MessageBox(0,PChar('Bancos de dados substituidos: '+ sLineBreak+ Substituido),
        'Colibri',MB_ICONINFORMATION);
  end
  else
  begin
    FrmPrincipal.ProgressBarIncializa(3,'Conexao LOCAL...');
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\'+FileName,FrmDataModule.ADOConnectionColibri);
    FrmPrincipal.ProgressBarIncremento(1);
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\dbConsulta.mdb',FrmDataModule.ADOConnectionConsulta);
    FrmPrincipal.ProgressBarIncremento(1);
    FrmPrincipal.conectarBDDireto(enderecoLOCAL+'\LOCAL\dbRT.mdb',FrmDataModule.ADOConnectionRT);
    FrmPrincipal.ProgressBarIncremento(1);
  end;
  FrmPrincipal.ProgressBarAtualizar;
  FrmPrincipal.Caption:= 'Colibri: LOCAL->REDE';
  FrmPrincipal.booLOCAL:= true;
  actLOCAL.Enabled:= false;
  actREDE.Enabled:= true;
  PanelTitulo.Caption:= 'Conectado LOCAL';
  PanelTitulo.Color:= clRed;
  if  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM) OR
      (FrmPrincipal.logPerfil = FrmPrincipal.PERFILSUPERVISAO) OR
      (FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO) OR
      (FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT) then
    actUpload.Enabled:= true
  else
    actUpload.Enabled:= false;
  actDownload.Enabled:= true;
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
end;

procedure TFrmConexaoLOCAL.actREDEExecute(Sender: TObject);
  var
    enderecoREDE,FilePath,dbConsulta,dbRT: String;
begin
  enderecoREDE:= FrmPrincipal.RegistroEndereco('Banco de dados');
  FilePath:= ExtractFilePath(enderecoREDE);
  dbConsulta:= FilePath+'\dbConsulta.mdb';
  dbRT:= FilePath+'\dbRT.mdb';
  FrmPrincipal.ProgressBarIncializa(3,'Conexao REDE...');
  FrmPrincipal.conectarBD(enderecoREDE,FrmDataModule.ADOConnectionColibri);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.conectarBDDireto(dbConsulta,FrmDataModule.ADOConnectionConsulta);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.conectarBDDireto(dbRT,FrmDataModule.ADOConnectionRT);
  FrmPrincipal.ProgressBarIncremento(1);
  FrmPrincipal.ProgressBarAtualizar;
  FrmPrincipal.booLOCAL:= false;
  actLOCAL.Enabled:= true;
  actREDE.Enabled:= false;
  PanelTitulo.Caption:= 'Conectado REDE';
  PanelTitulo.Color:= clGreen;
  actUpload.Enabled:= false;
  actDownload.Enabled:= false;
  FrmDataModule.ADOQueryColibri.Active:= false;
  FrmDataModule.ADOQueryColibri.Active:= true;
end;

procedure TFrmConexaoLOCAL.actSelDownloadExecute(Sender: TObject);
begin
  if chkDownColibri.Checked and chkDownColibri.Enabled then
  begin
    chkDownColibri.Checked:= false;
    actSelDownload.ImageIndex:=  232;
  end
  else if not chkDownColibri.Checked and chkDownColibri.Enabled then
  begin
    chkDownColibri.Checked:= true;
    actSelDownload.ImageIndex:=  231;
  end;
  //----------------------------------------------------------
  if chkDownConsulta.Checked and chkDownConsulta.Enabled then
    chkDownConsulta.Checked:= false
  else if not chkDownConsulta.Checked and chkDownConsulta.Enabled then
    chkDownConsulta.Checked:= true;
  //----------------------------------------------------------
  if chkDownRT.Checked and chkDownRT.Enabled then
    chkDownRT.Checked:= false
  else if not chkDownRT.Checked and chkDownRT.Enabled then
    chkDownRT.Checked:= true;
end;

procedure TFrmConexaoLOCAL.actSelUploadExecute(Sender: TObject);
begin
  if chkUpColibri.Checked and chkUpColibri.Enabled then
  begin
    chkUpColibri.Checked:= false;
    actSelUpload.ImageIndex:=  232;
  end
  else if not chkUpColibri.Checked and chkUpColibri.Enabled then
  begin
    chkUpColibri.Checked:= true;
    actSelUpload.ImageIndex:=  231;
  end;
  //----------------------------------------------------------
  if chkUPConsulta.Checked and chkUPConsulta.Enabled then
    chkUPConsulta.Checked:= false
  else if not chkUPConsulta.Checked and chkUPConsulta.Enabled then
    chkUPConsulta.Checked:= true;
  //----------------------------------------------------------
  if chkUPRT.Checked and chkUPRT.Enabled then
    chkUPRT.Checked:= false
  else if not chkUPRT.Checked and chkUPRT.Enabled then
    chkUPRT.Checked:= true;
end;

procedure TFrmConexaoLOCAL.actUploadExecute(Sender: TObject);
  var
    FileName,FilePath,enderecoREDE,
    enderecoREDEdbConsulta,enderecoREDEdbRT,
    enderecoLOCALdbColibri,enderecoLOCALdbConsulta,enderecoLOCALdbRT: String;
begin
  enderecoREDE:= FrmPrincipal.registroEndereco('Banco de dados');
  FileName:= ExtractFileName(enderecoREDE);
  FilePath:= ExtractFilePath(enderecoREDE);
  //LOCAL
  enderecoLOCALdbColibri:= ExtractFilePath(Application.ExeName)+'LOCAL\'+FileName;
  enderecoLOCALdbConsulta:= ExtractFilePath(Application.ExeName)+'LOCAL\dbConsulta.mdb';
  enderecoLOCALdbRT:= ExtractFilePath(Application.ExeName)+'LOCAL\dbRT.mdb';
  //REDE
  enderecoREDEdbConsulta:= FilePath+'dbConsulta.mdb';
  enderecoREDEdbRT:= FilePath+'dbRT.mdb';
  //UPLOAD
  FrmPrincipal.ProgressBarIncializa(3,'Upload banco de dados...');
  //============================================================================
  if chkUPColibri.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbColibri da REDE?'),'.::ATENCAO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbColibri), PChar(enderecoREDE), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  if chkUPConsulta.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbConsulta da REDE?'),'.::ATENCAO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbConsulta), PChar(enderecoREDEdbConsulta), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  if chkUPRT.Checked then
    if Application.MessageBox(PChar('Deseja realmente substituir o banco de dados dbRT da REDE?'),'.::ATENCAO::.',36) = 6 then
      CopyFile(PChar(enderecoLOCALdbRT), PChar(enderecoREDEdbRT), false);
  FrmPrincipal.ProgressBarIncremento(1);
  //============================================================================
  FrmPrincipal.ProgressBarAtualizar;
end;

procedure TFrmConexaoLOCAL.FormCreate(Sender: TObject);
begin
  enderecoREDE:= ExtractFilePath(FrmPrincipal.registroEndereco('Banco de dados'));
  enderecoLOCAL:= ExtractFilePath(Application.ExeName)+'LOCAL';
  edtRede.Text:= enderecoREDE;
  edtLocal.Text:= enderecoLOCAL;
  if FrmPrincipal.booLOCAL then
  begin
    actLOCAL.Enabled:= false;
    actREDE.Enabled:= true;
    PanelTitulo.Caption:= 'Conectado LOCAL';
    PanelTitulo.Color:= clRed;
  end
  else
  begin
    actLOCAL.Enabled:= true;
    actREDE.Enabled:= false;
    PanelTitulo.Caption:= 'Conectado REDE';
    PanelTitulo.Color:= clGreen;
  end;
  if ((FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)or
  (FrmPrincipal.logPerfil = FrmPrincipal.PERFILADM)) then //Administrador
  begin
    actUpload.Enabled:= true;
    actDownload.Enabled:= true;
    HabilitaChkBox(chkUPColibri,true);
    HabilitaChkBox(chkUPConsulta,true);
    HabilitaChkBox(chkUPRT,true);
    HabilitaChkBox(chkDownColibri,true);
    HabilitaChkBox(chkDownConsulta,true);
    HabilitaChkBox(chkDownRT,true);
  end
  else if FrmPrincipal.logPerfil = FrmPrincipal.PERFILPROGRAMACAO then
  begin
    actUpload.Enabled:= true;
    actDownload.Enabled:= true;
    HabilitaChkBox(chkUPColibri,true);
    HabilitaChkBox(chkUPConsulta,true);
    HabilitaChkBox(chkUPRT,true);
    HabilitaChkBox(chkDownColibri,true);
    HabilitaChkBox(chkDownConsulta,true);
    HabilitaChkBox(chkDownRT,true);
  end
  else if FrmPrincipal.logPerfil = FrmPrincipal.PERFILRT then
  begin
    actUpload.Enabled:= true;
    actDownload.Enabled:= true;
    HabilitaChkBox(chkUPColibri,false);
    HabilitaChkBox(chkUPConsulta,true);
    HabilitaChkBox(chkUPRT,true);
    HabilitaChkBox(chkDownColibri,true);
    HabilitaChkBox(chkDownConsulta,true);
    HabilitaChkBox(chkDownRT,true);
  end
  else //if logPerfil = 'Consulta' then
  begin
    actUpload.Enabled:= false;
    actDownload.Enabled:= true;
    HabilitaChkBox(chkUPColibri,false);
    HabilitaChkBox(chkUPConsulta,false);
    HabilitaChkBox(chkUPRT,false);
    HabilitaChkBox(chkDownColibri,true);
    HabilitaChkBox(chkDownConsulta,true);
    HabilitaChkBox(chkDownRT,true);
  end;
end;

procedure TFrmConexaoLOCAL.HabilitaChkBox(chk: TCheckBox; sel: Boolean);
begin
  chk.Checked:= sel;
  chk.Enabled:= sel;
end;

end.
