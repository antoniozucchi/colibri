object FrmDataModule: TFrmDataModule
  Height = 781
  Width = 2260
  PixelsPerInch = 120
  object ADOQueryProgramacaoServico_Cadastro: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Servico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (CodigoProgramacaoDiaria=@Servico);')
    Left = 60
    Top = 10
  end
  object ADOQueryProgramacaoExecutante_Cadastro: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (CodigoProgramacaoDiaria=@CodigoProgramacaoDiaria)'
      'ORDER BY NomeExecutante;')
    Left = 10
    Top = 10
  end
  object ADOQueryProgramacaoDiaria_Consulta: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria;')
    Left = 100
    Top = 10
  end
  object DataSourceProgramacaoDiaria_Consulta: TDataSource
    DataSet = ADOQueryProgramacaoDiaria_Consulta
    Left = 10
    Top = 560
  end
  object DataSourceProgramacaoExecutante_Cadastro: TDataSource
    DataSet = ADOQueryProgramacaoExecutante_Cadastro
    Left = 50
    Top = 560
  end
  object DataSourceProgramacaoServico_Cadastro: TDataSource
    DataSet = ADOQueryProgramacaoServico_Cadastro
    Left = 250
    Top = 610
  end
  object ADOConnectionColibri: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\De' +
      'v\Projetos\Colibri\Banco de dados\UO-SEAL-ATP-OP-SM\bdColibri.co' +
      'libri;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB' +
      ':System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Databas' +
      'e Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking' +
      ' Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bul' +
      'k Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Cr' +
      'eate System Database=False;Jet OLEDB:Encrypt Database=False;Jet ' +
      'OLEDB:Don'#39't Copy Locale on Compact=False;Jet OLEDB:Compact Witho' +
      'ut Replica Repair=False;Jet OLEDB:SFP=False;'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 282
    Top = 316
  end
  object ADOQueryPlataforma: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryPlataformaBeforePost
    AfterRefresh = ADOQueryPlataformaAfterRefresh
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPlataforma.*'
      'FROM tblPlataforma'
      'ORDER BY Plataforma;')
    Left = 1300
    Top = 100
  end
  object DataSourcePlataforma: TDataSource
    DataSet = ADOQueryPlataforma
    Left = 290
    Top = 610
  end
  object ADOQueryExecutante: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryExecutanteBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblExecutante.*'
      'FROM tblExecutante'
      'ORDER BY NomeExecutante;')
    Left = 1420
    Top = 100
  end
  object DataSourceExecutante: TDataSource
    DataSet = ADOQueryExecutante
    Left = 90
    Top = 560
  end
  object ADOQueryUsuario: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblUsuario.*'
      'FROM tblUsuario'
      'ORDER BY Chave;')
    Left = 140
    Top = 10
  end
  object DataSourceUsuario: TDataSource
    DataSet = ADOQueryUsuario
    Left = 210
    Top = 610
  end
  object ADOQueryCarregaExecutanteAPLAT: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2016'
      end>
    SQL.Strings = (
      'SELECT tblExecutanteAPLAT.*'
      'FROM tblExecutanteAPLAT'
      
        'WHERE ((DataEmbarque <= @DataProcura) AND (DataDesembarque >=@Da' +
        'taProcura));')
    Left = 950
  end
  object DataSourceCarregaExecutanteAPLAT: TDataSource
    DataSet = ADOQueryCarregaExecutanteAPLAT
    Left = 940
    Top = 440
  end
  object ADOQueryCarregaExecutanteSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2016'
      end>
    SQL.Strings = (
      'SELECT tblExecutanteSAP.*'
      'FROM tblExecutanteSAP'
      'WHERE (DataEmbarque LIKE @DataProcura);')
    Left = 980
  end
  object DataSourceCarregaExecutanteSAP: TDataSource
    DataSet = ADOQueryCarregaExecutanteSAP
    Left = 960
    Top = 610
  end
  object DataSourceProgramacaoDiaria_Cadastro: TDataSource
    DataSet = ADOQueryProgramacaoDiaria_Cadastro
    Left = 420
    Top = 500
  end
  object ADOQueryProgramacaoDiaria_Cadastro: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria'
      'WHERE (txtTipoEtapaServico = @TipoEtapaServico)'
      'ORDER BY txtDestino;')
    Left = 260
    Top = 10
  end
  object ADOQueryImportarExecutanteAPLAT: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryImportarExecutanteAPLATBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblExecutanteAPLAT.*'
      'FROM tblExecutanteAPLAT'
      'ORDER BY NomeExecutante;')
    Left = 920
    Top = 50
  end
  object DataSourceImportarExecutanteAPLAT: TDataSource
    DataSet = ADOQueryImportarExecutanteAPLAT
    Left = 910
    Top = 490
  end
  object ADOQueryImportarExecutanteSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryImportarExecutanteSAPBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblExecutanteSAP.*'
      'FROM tblExecutanteSAP'
      'ORDER BY NomeExecutante;')
    Left = 950
    Top = 50
  end
  object DataSourceImportarExecutanteSAP: TDataSource
    DataSet = ADOQueryImportarExecutanteSAP
    Left = 1060
    Top = 390
  end
  object ADOQueryImportarCarteiraTrabalho: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho'
      'ORDER BY Plataforma;')
    Left = 890
    Top = 50
  end
  object DataSourceImportarCarteiraTrabalho: TDataSource
    DataSet = ADOQueryImportarCarteiraTrabalho
    Left = 880
    Top = 490
  end
  object ADOQueryConsultaExecutante_DataCodigoSAP: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoSAP'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2019'
      end>
    SQL.Strings = (
      
        'SELECT tblProgramacaoExecutante.*, tblProgramacaoDiaria.DataProg' +
        'ramacao'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria'
      
        'WHERE ((tblProgramacaoExecutante.CodigoSAP= @CodigoSAP)AND(tblPr' +
        'ogramacaoDiaria.DataProgramacao=@DataProgramacao));')
    Left = 10
    Top = 70
  end
  object DataSourceConsultaExecutante_DataCodigoSAP: TDataSource
    DataSet = ADOQueryConsultaExecutante_DataCodigoSAP
    Left = 10
    Top = 500
  end
  object DataSourceInserirExecutante: TDataSource
    DataSet = ADOQueryInserirExecutante1
    Left = 90
    Top = 440
  end
  object ADOQueryInserirExecutante1: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (CodigoProgramacaoDiaria= 0);')
    Left = 410
    Top = 120
  end
  object DataSourceInserirServico: TDataSource
    DataSet = ADOQueryInserirServico
    Left = 130
    Top = 440
  end
  object ADOQueryInserirServico: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (CodigoProgramacaoDiaria= 0);')
    Left = 460
    Top = 10
  end
  object DataSourceInserirProgramacao: TDataSource
    DataSet = ADOQueryInserirProgramacao
    Left = 170
    Top = 440
  end
  object ADOQueryInserirProgramacao: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria;')
    Left = 500
    Top = 10
  end
  object ADOQueryConsultaProgramacao_ID: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@ProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria'
      'WHERE (idProgramacaoDiaria = @ProgramacaoDiaria);')
    Left = 220
    Top = 10
  end
  object DataSourceConsultaProgramacao_ID: TDataSource
    DataSet = ADOQueryConsultaProgramacao_ID
    Left = 380
    Top = 500
  end
  object DataSourceConsultaPlataforma: TDataSource
    DataSet = ADOQueryConsultaPlataforma
    Left = 250
    Top = 500
  end
  object ADOQueryConsultaPlataforma: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPlataforma.*'
      'FROM tblPlataforma'
      'ORDER BY Plataforma;')
    Left = 1380
    Top = 100
  end
  object DataSourceInserirExecutanteCadastro: TDataSource
    DataSet = ADOQueryInserirExecutanteCadastro
    Left = 210
    Top = 500
  end
  object ADOQueryInserirExecutanteCadastro: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblExecutante.*'
      'FROM tblExecutante;')
    Left = 1340
    Top = 100
  end
  object DataSourceConsultaExecutante_TipoEtapaServico: TDataSource
    DataSet = ADOQueryConsultaExecutante_TipoEtapaServico
    Left = 50
    Top = 440
  end
  object ADOQueryConsultaExecutante_TipoEtapaServico: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblExecutante.*'
      'FROM tblExecutante'
      'WHERE (codigoTipoEtapaServico=@TipoEtapaServico)'
      'ORDER BY NomeExecutante;')
    Left = 1370
    Top = 50
  end
  object ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblTipoEtapaServico.*'
      'FROM tblTipoEtapaServico'
      'WHERE (TipoEtapaServico= @TipoEtapaServico);')
    Left = 1330
    Top = 50
  end
  object DataSourceConsultaCodigoTipoEtapaServico_TipoEtapaServico: TDataSource
    DataSet = ADOQueryConsultaCodigoTipoEtapaServico_TipoEtapaServico
    Left = 210
    Top = 440
  end
  object ADOQueryTipoEtapaServico: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryTipoEtapaServicoBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblTipoEtapaServico.*'
      'FROM tblTipoEtapaServico'
      'ORDER BY TipoEtapaServico;')
    Left = 1460
    Top = 100
  end
  object DataSourceTipoEtapaServico: TDataSource
    DataSet = ADOQueryTipoEtapaServico
    Left = 170
    Top = 610
  end
  object ADOQueryConsultaTipoEtapaServico_ID: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblTipoEtapaServico.*'
      'FROM tblTipoEtapaServico'
      'WHERE (idTipoEtapaServico= @TipoEtapaServico);')
    Left = 1400
    Top = 50
  end
  object DataSourceConsultaTipoEtapaServico_ID: TDataSource
    DataSet = ADOQueryConsultaTipoEtapaServico_ID
    Left = 10
    Top = 440
  end
  object ADOQueryCarteiraOM: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho;')
    Left = 1010
  end
  object DataSourceCarteiraOM: TDataSource
    DataSet = ADOQueryCarteiraOM
    Left = 990
    Top = 610
  end
  object ADOQueryGerenciarSolicitacoes: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/09/2016'
      end>
    SQL.Strings = (
      
        'SELECT tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.Dat' +
        'aProgramacao'
      'FROM tblProgramacaoDiaria'
      
        'GROUP BY tblProgramacaoDiaria.txtDestino, tblProgramacaoDiaria.D' +
        'ataProgramacao'
      'HAVING (DataProgramacao like @DataProgramacao);')
    Left = 540
    Top = 10
  end
  object DataSourceGerenciarSolicitacoes: TDataSource
    DataSet = ADOQueryGerenciarSolicitacoes
    Left = 250
    Top = 440
  end
  object ADOQueryConsultaPlataforma_Nome: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Plataforma'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblPlataforma.*'
      'FROM tblPlataforma'
      'WHERE (Plataforma LIKE @Plataforma);')
    Left = 1300
    Top = 50
  end
  object DataSourceConsultaPlataforma_Nome: TDataSource
    DataSet = ADOQueryConsultaPlataforma_Nome
    Left = 290
    Top = 440
  end
  object DataSourceGerenciarExecutante: TDataSource
    DataSet = ADOQueryGerenciarExecutante
    Left = 330
    Top = 440
  end
  object ADOQueryGerenciarExecutante: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' '
      
        'tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecuta' +
        'nte.CodigoProgramacaoDiaria '
      
        'WHERE (((tblProgramacaoDiaria.DataProgramacao LIKE '#39'01/01/2016'#39')' +
        ') '
      
        'AND (tblProgramacaoDiaria.txtDestino LIKE '#39'PCM-09'#39')) ORDER BY Or' +
        'igem;')
    Left = 50
    Top = 70
  end
  object DataSourceGerenciarServico: TDataSource
    DataSet = ADOQueryGerenciarServico
    Left = 370
    Top = 440
  end
  object ADOQueryGerenciarServico: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Servico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (CodigoProgramacaoDiaria=@Servico);')
    Left = 90
    Top = 120
  end
  object DataSourceGerenciarProgramacao: TDataSource
    DataSet = ADOQueryGerenciarProgramacao
    Left = 410
    Top = 440
  end
  object ADOQueryGerenciarProgramacao: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Destino'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria'
      
        'WHERE ((DataProgramacao like @DataProcura) AND (txtDestino like ' +
        '@Destino))'
      'ORDER BY txtTipoEtapaServico;')
    Left = 130
    Top = 120
  end
  object DataSourceConsultaProgramacao_Destino_DataProgramacao: TDataSource
    DataSet = ADOQueryConsultaProgramacao_Destino_DataProgramacao
    Left = 450
    Top = 440
  end
  object ADOQueryConsultaProgramacao_Destino_DataProgramacao: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Destino'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria'
      
        'WHERE ((DataProgramacao like @DataProcura) AND (txtDestino like ' +
        '@Destino))'
      'ORDER BY txtTipoEtapaServico;')
    Left = 170
    Top = 120
  end
  object DataSourceConsultaProgramacaoExecutante_ID: TDataSource
    DataSet = ADOQueryConsultaProgramacaoExecutante_ID
    Left = 490
    Top = 440
  end
  object ADOQueryConsultaProgramacaoExecutante_ID: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@idProgramacaoExecutante'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (idProgramacaoExecutante = @idProgramacaoExecutante );')
    Left = 210
    Top = 120
  end
  object DataSourceEmbarcacoes: TDataSource
    DataSet = ADOQueryEmbarcacoes
    Left = 50
    Top = 500
  end
  object ADOQueryEmbarcacoes: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryEmbarcacoesBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblEmbarcacao.*'
      'FROM tblEmbarcacao'
      'ORDER BY NomeEmbarcacao;')
    Left = 1300
    Top = 150
  end
  object ADOQueryTM_Embarcacao: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblEmbarcacao.*'
      'FROM tblEmbarcacao'
      'WHERE ((TipoEmbarcacao like '#39'Lancha de Passageiros'#39') AND'
      '(StatusEmbarcacao like '#39'Operando'#39'))'
      'ORDER BY NomeEmbarcacao;')
    Left = 1390
    Top = 150
  end
  object DataSourceTM_Embarcacao: TDataSource
    DataSet = ADOQueryTM_Embarcacao
    Left = 90
    Top = 500
  end
  object ADOQueryRoteamento: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryRoteamentoBeforePost
    Parameters = <
      item
        Name = '@DataRoteamento'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2017'
      end>
    SQL.Strings = (
      'SELECT tblRoteamento.*'
      'FROM tblRoteamento'
      'WHERE (DataRoteamento LIKE @DataRoteamento)'
      'ORDER BY NomeRota,HoraRoteamento;')
    Left = 250
    Top = 70
  end
  object DataSourceRoteamento: TDataSource
    DataSet = ADOQueryRoteamento
    Left = 130
    Top = 500
  end
  object ADOQueryTM_Plataformas: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPlataforma.*'
      'FROM tblPlataforma'
      'ORDER BY Plataforma;')
    Left = 1460
  end
  object DataSourceTM_Plataformas: TDataSource
    DataSet = ADOQueryTM_Plataformas
    Left = 500
    Top = 500
  end
  object ADOQueryTM_RotaExecutantes: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoRoteamento'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      
        'SELECT tblRoteamento.*, tblAux_Rota_Distribuicao.*, tblProgramac' +
        'aoExecutante.*, tblProgramacaoDiaria.*'
      
        'FROM tblProgramacaoDiaria INNER JOIN (tblRoteamento INNER JOIN (' +
        'tblProgramacaoExecutante INNER JOIN tblAux_Rota_Distribuicao ON ' +
        'tblProgramacaoExecutante.idProgramacaoExecutante = tblAux_Rota_D' +
        'istribuicao.CodigoProgramacaoExecutante) ON tblRoteamento.idRote' +
        'amento = tblAux_Rota_Distribuicao.CodigoRota) ON tblProgramacaoD' +
        'iaria.idProgramacaoDiaria = tblProgramacaoExecutante.CodigoProgr' +
        'amacaoDiaria'
      'WHERE (tblAux_Rota_Distribuicao.CodigoRota=@CodigoRoteamento)'
      'ORDER BY NomeExecutante;')
    Left = 250
    Top = 120
  end
  object DataSourceTM_RotaExecutantes: TDataSource
    DataSet = ADOQueryTM_RotaExecutantes
    Left = 170
    Top = 500
  end
  object DataSourceExecutantesDisponiveis: TDataSource
    DataSet = ADOQueryExecutantesDisponiveis
    Left = 530
    Top = 440
  end
  object ADOQueryExecutantesDisponiveis: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria '
      
        'WHERE (tblProgramacaoExecutante.InseridoProgramacaoTransporte LI' +
        'KE FALSE) '
      
        'AND (tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante' +
        '.Origem) '
      
        'AND (tblProgramacaoExecutante.StatusProgramacao LIKE '#39'Aprovado'#39')' +
        ' '
      'AND (tblProgramacaoDiaria.DataProgramacao LIKE @DataProcura) '
      'ORDER BY tblProgramacaoDiaria.txtDestino;')
    Left = 290
    Top = 120
  end
  object ADOQueryConsultaExecutantes_DataProgramacao: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*, tblProgramacaoDiaria.*'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria'
      
        'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE @DataProgramaca' +
        'o);')
    Left = 90
    Top = 70
  end
  object DataSourceConsultaExecutantes_DataProgramacao: TDataSource
    DataSet = ADOQueryConsultaExecutantes_DataProgramacao
    Left = 340
    Top = 500
  end
  object ADOQueryAux_Rota_Distribuicao: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblAux_Rota_Distribuicao.*'
      'FROM tblAux_Rota_Distribuicao;')
    Left = 130
    Top = 70
  end
  object DataSourceAux_Rota_Distribuicao: TDataSource
    DataSet = ADOQueryAux_Rota_Distribuicao
    Left = 520
    Top = 620
  end
  object ADOQueryExcluirVinculoExecutante: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@idAux_Rota_Distribuicao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblAux_Rota_Distribuicao.*'
      'FROM tblAux_Rota_Distribuicao'
      'WHERE (idAux_Rota_Distribuicao= @idAux_Rota_Distribuicao);')
    Left = 170
    Top = 70
  end
  object DataSourceExcluirVinculoExecutante: TDataSource
    DataSet = ADOQueryExcluirVinculoExecutante
    Left = 460
    Top = 500
  end
  object ADOQueryTemporarioDBColibri: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoSAP'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '0'
      end
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2016'
      end>
    SQL.Strings = (
      
        'SELECT tblProgramacaoExecutante.*, tblProgramacaoDiaria.DataProg' +
        'ramacao'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria'
      
        'WHERE ((tblProgramacaoExecutante.CodigoSAP= @CodigoSAP) AND (tbl' +
        'ProgramacaoDiaria.DataProgramacao=@DataProgramacao));')
    Left = 210
    Top = 70
  end
  object DataSourceTemporarioDBColibri: TDataSource
    DataSet = ADOQueryTemporarioDBColibri
    Left = 140
    Top = 560
  end
  object ADOQueryConsultaServicosProgramados: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/09/2016'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoServico.*'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoServico ON tb' +
        'lProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoServico.C' +
        'odigoProgramacaoDiaria'
      'WHERE (tblProgramacaoDiaria.DataProgramacao = @DataProgramacao)'
      'ORDER BY tblProgramacaoDiaria.txtTipoEtapaServico;')
    Left = 290
    Top = 70
  end
  object DataSourceConsultaServicosProgramados: TDataSource
    DataSet = ADOQueryConsultaServicosProgramados
    Left = 180
    Top = 560
  end
  object DataSourceExecutantesSAP: TDataSource
    DataSet = ADOQueryExecutantesSAP
    Left = 1050
    Top = 610
  end
  object ADOQueryExecutantesSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblExecutanteSAP.*'
      'FROM tblExecutanteSAP'
      'WHERE (DataEmbarque LIKE @DataProcura)'
      'ORDER BY NomeExecutante;')
    Left = 1070
  end
  object ADOQueryConsultaEmbarcacao: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblEmbarcacao.*'
      'FROM tblEmbarcacao'
      'ORDER BY NomeEmbarcacao;')
    Left = 1360
    Top = 150
  end
  object DataSourceConsultaEmbarcacao: TDataSource
    DataSet = ADOQueryConsultaEmbarcacao
    Left = 250
    Top = 560
  end
  object ADOQueryConsultaExecutantesProgramados: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*, tblProgramacaoDiaria.*'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria'
      
        'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE @DataProgramaca' +
        'o);')
    Left = 320
    Top = 70
  end
  object DataSourceConsultaExecutantesProgramados: TDataSource
    DataSet = ADOQueryConsultaExecutantesProgramados
    Left = 290
    Top = 560
  end
  object DataSourceForaOperacao: TDataSource
    DataSet = ADOQueryForaOperacao
    Left = 850
    Top = 490
  end
  object DataSourceColibri: TDataSource
    DataSet = ADOQueryColibri
    Left = 330
    Top = 560
  end
  object DataSourceGeradores: TDataSource
    DataSet = ADOQueryGeradores
    Left = 330
    Top = 610
  end
  object ADOQueryForaOperacao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryForaOperacaoBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.* '
      'FROM tblNotaManutencao'
      'ORDER BY Plataforma;')
    Left = 860
    Top = 50
  end
  object ADOQueryColibri: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblColibri.* '
      'FROM tblColibri;')
    Left = 360
    Top = 70
  end
  object ADOQueryGeradores: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryGeradoresBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblGerador.* '
      'FROM tblGerador'
      'ORDER BY NomeGerador;')
    Left = 1260
    Top = 150
  end
  object ADOQueryConsultaExecutante_CodigoSAP: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryGeradoresBeforePost
    Parameters = <
      item
        Name = '@CodigoSAP'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblExecutante.*'
      'FROM tblExecutante'
      'WHERE (CodigoSAP = @CodigoSAP);')
    Left = 1440
    Top = 50
  end
  object DataSourceConsultaExecutante_CodigoSAP: TDataSource
    DataSet = ADOQueryConsultaExecutante_CodigoSAP
    Left = 540
    Top = 500
  end
  object DataSourceImportarNotaManutencao: TDataSource
    DataSet = ADOQueryImportarNotaManutencao
    Left = 940
    Top = 490
  end
  object ADOQueryImportarNotaManutencao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.*'
      'FROM tblNotaManutencao;')
    Left = 980
    Top = 50
  end
  object DataSourceConsultaProgramacaoServico_ID: TDataSource
    DataSet = ADOQueryConsultaProgramacaoServico_ID
    Left = 410
    Top = 560
  end
  object ADOQueryConsultaProgramacaoServico_ID: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = 'idProgramacaoServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (idProgramacaoServico = @idProgramacaoServico);')
    Left = 400
    Top = 70
  end
  object DataSourceContarExecCancelados_ID: TDataSource
    DataSet = ADOQueryNumAprovado
    Left = 450
    Top = 560
  end
  object ADOQueryNumAprovado: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE ((CodigoProgramacaoDiaria LIKE @CodigoProgramacaoDiaria)'
      'AND(StatusProgramacao LIKE '#39'Aprovado'#39'));')
    Left = 440
    Top = 70
  end
  object ADOQueryNumCancelado: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE ((CodigoProgramacaoDiaria= @CodigoProgramacaoDiaria)AND'
      '(StatusProgramacao = '#39'Cancelado'#39'));')
    Left = 480
    Top = 70
  end
  object DataSourceContarExecTotal_ID: TDataSource
    DataSet = ADOQueryNumCancelado
    Left = 490
    Top = 560
  end
  object ADOQueryProgramacaoExecutante_Consulta: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (CodigoProgramacaoDiaria=@CodigoProgramacaoDiaria)'
      'ORDER BY NomeExecutante;')
    Left = 520
    Top = 70
  end
  object DataSourceProgramacaoExecutante_Consulta: TDataSource
    DataSet = ADOQueryProgramacaoExecutante_Consulta
    Left = 530
    Top = 560
  end
  object DataSourceServicosSAP_TipoEtapaServico_Destino: TDataSource
    DataSet = ADOQueryServicosSAP_TipoEtapaServico_Destino
    Left = 820
    Top = 440
  end
  object ADOQueryServicosSAP_TipoEtapaServico_Destino: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Destino'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.* FROM tblCarteiraTrabalho '
      'WHERE ((txtTipoEtapaServicoOP LIKE  @TipoEtapaServico)AND'
      '(Plataforma LIKE @Destino)AND(StatusUsuarioOM LIKE "%GIOP%"));')
    Left = 830
  end
  object DataSourceProgramacaoServico_Consulta: TDataSource
    DataSet = ADOQueryProgramacaoServico_Consulta
    Left = 570
    Top = 560
  end
  object ADOQueryProgramacaoServico_Consulta: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Servico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (CodigoProgramacaoDiaria=@Servico);')
    Left = 180
    Top = 10
  end
  object ADOQueryNumExecutantes: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (CodigoProgramacaoDiaria=@CodigoProgramacaoDiaria);')
    Left = 450
    Top = 120
  end
  object DataSourceNumExecutantes: TDataSource
    DataSet = ADOQueryNumExecutantes
    Left = 450
    Top = 610
  end
  object ADOQueryImportarTextoLongoCarteira: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblTextoLongoCarteira.*'
      'FROM tblTextoLongoCarteira;')
    Left = 1010
    Top = 50
  end
  object DataSourceImportarTextoLongoCarteira: TDataSource
    DataSet = ADOQueryImportarTextoLongoCarteira
    Left = 970
    Top = 390
  end
  object ADOQueryProgramados: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      
        'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.*, tblRo' +
        'teamento.* '
      
        'FROM tblRoteamento INNER JOIN ((tblProgramacaoDiaria INNER JOIN ' +
        'tblProgramacaoExecutante ON tblProgramacaoDiaria.idProgramacaoDi' +
        'aria = '
      
        'tblProgramacaoExecutante.CodigoProgramacaoDiaria) INNER JOIN tbl' +
        'Aux_Rota_Distribuicao ON tblProgramacaoExecutante.idProgramacaoE' +
        'xecutante = '
      
        'tblAux_Rota_Distribuicao.CodigoProgramacaoExecutante) ON tblRote' +
        'amento.idRoteamento = tblAux_Rota_Distribuicao.CodigoRota '
      'WHERE (tblProgramacaoDiaria.DataProgramacao LIKE '#39'05/06/2017'#39') '
      
        'AND (tblProgramacaoDiaria.txtDestino <> tblProgramacaoExecutante' +
        '.Origem) '
      
        'AND (((tblProgramacaoExecutante.StatusProgramacao like '#39'Aprovado' +
        #39'))) '
      'ORDER BY NomeRota,txtDestino,txtTipoEtapaServico;')
    Left = 10
    Top = 120
  end
  object DataSourceProgramados: TDataSource
    DataSet = ADOQueryProgramados
    Left = 370
    Top = 610
  end
  object ADOQueryCarteiraOP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho;')
    Left = 1040
  end
  object DataSourceCarteiraOP: TDataSource
    DataSet = ADOQueryCarteiraOP
    Left = 1020
    Top = 610
  end
  object DataSourceCentroTrabalho_Descricao: TDataSource
    DataSet = ADOQueryCentroTrabalho_Descricao
    Left = 300
    Top = 500
  end
  object ADOQueryCentroTrabalho_Descricao: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CentroTrabalho'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@txtTipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblCentroTrabalho.*'
      'FROM tblCentroTrabalho'
      'WHERE ((CentroTrabalho = @CentroTrabalho)'
      'AND(txtTipoEtapaServico = @txtTipoEtapaServico));')
    Left = 1260
    Top = 100
  end
  object DataSourceTipoEtapaServico_Carteira: TDataSource
    DataSet = ADOQueryTipoEtapaServico_Carteira
    Left = 480
    Top = 610
  end
  object ADOQueryTipoEtapaServico_Carteira: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblTipoEtapaServico.*'
      'FROM tblTipoEtapaServico'
      'WHERE (BooleanCarteira = true)'
      'ORDER BY TipoEtapaServico;')
    Left = 1470
    Top = 150
  end
  object ADOQueryPrioridadeForcada: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryPrioridadeForcadaBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPrioridadeForcada.*'
      'FROM tblPrioridadeForcada'
      'ORDER BY Prioridade;')
    Left = 920
  end
  object DataSourcePrioridadeForcada: TDataSource
    DataSet = ADOQueryPrioridadeForcada
    Left = 910
    Top = 440
  end
  object ADOQueryLocalInstalacaoSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblLocalInstalacao.* '
      'FROM tblLocalInstalacao '
      'ORDER BY LocalInstalacao;')
    Left = 790
    Top = 100
  end
  object DataSourceLocalInstalacaoSAP: TDataSource
    DataSet = ADOQueryLocalInstalacaoSAP
    Left = 50
    Top = 610
  end
  object ADOConnectionMemoria: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\Us' +
      'ers\augus\Documents\Projetos\Colibri\Banco de dados\UO-SEAL-ATP-' +
      'OP-SM\dbMemoria.mdb;Mode=Share Deny None;Persist Security Info=F' +
      'alse;Jet OLEDB:System database="";Jet OLEDB:Registry Path="";Jet' +
      ' OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Da' +
      'tabase Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OL' +
      'EDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="' +
      '";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Datab' +
      'ase=False;Jet OLEDB:Don'#39't Copy Locale on Compact=False;Jet OLEDB' +
      ':Compact Without Replica Repair=False;Jet OLEDB:SFP=False;'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 50
    Top = 220
  end
  object DataSourceIW37N: TDataSource
    DataSet = ADOQueryIW37N
    Left = 1030
    Top = 390
  end
  object ADOQueryIW37N: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho;')
    Left = 1070
    Top = 50
  end
  object DataSourceCadastroUsuario: TDataSource
    DataSet = ADOQueryCadastroUsuario
    Left = 130
    Top = 610
  end
  object ADOQueryCadastroUsuario: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryCadastroUsuarioBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblUsuario.*'
      'FROM tblUsuario'
      'ORDER BY Chave;')
    Left = 50
    Top = 120
  end
  object DataSourceTemporarioDBConsulta1: TDataSource
    DataSet = ADOQueryTemporarioDBConsulta1
    Left = 850
    Top = 440
  end
  object ADOQueryTemporarioDBConsulta1: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCentroTrabalho.*'
      'FROM tblCentroTrabalho;')
    Left = 1510
    Top = 50
  end
  object ADOQueryRestricaoOperacional: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho'
      'WHERE (TextoBreveOM LIKE "RO:%")'
      'ORDER BY Plataforma, txtTipoEtapaServicoOM;')
    Left = 820
    Top = 50
  end
  object DataSourceRestricaoOperacional: TDataSource
    DataSet = ADOQueryRestricaoOperacional
    Left = 810
    Top = 490
  end
  object ADOQueryTotalCampo: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Dataprocura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Campo'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Status'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria '
      'WHERE ((tblProgramacaoDiaria.DataProgramacao  LIKE @Dataprocura)'
      'AND(tblProgramacaoDiaria.txtDestino LIKE @Campo)'
      'AND(tblProgramacaoExecutante.StatusProgramacao LIKE @Status)'
      
        'AND((tblProgramacaoExecutante.Origem LIKE "TMIB")OR(tblProgramac' +
        'aoExecutante.Origem LIKE "PCM-09")));')
    Left = 420
    Top = 10
  end
  object DataSourceTotalCampo: TDataSource
    DataSet = ADOQueryTotalCampo
    Left = 820
    Top = 550
  end
  object ADOQueryTotalOrigem: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@DataProcura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@Origem'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.* '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria '
      'WHERE ((tblProgramacaoDiaria.DataProgramacao  LIKE @DataProcura)'
      'AND(tblProgramacaoExecutante.Origem LIKE @Origem)'
      
        'AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado"))' +
        ';')
    Left = 380
    Top = 10
  end
  object DataSourceTotalOrigem: TDataSource
    DataSet = ADOQueryTotalOrigem
    Left = 860
    Top = 550
  end
  object ADOQueryPassageirosTM_SIM: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Dataprocura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = '@Origem'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = '@Destino'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.*  '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria '
      'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE @Dataprocura)'
      'AND(tblProgramacaoExecutante.Origem LIKE @Origem)'
      'AND(tblProgramacaoDiaria.txtDestino LIKE @Destino)'
      
        'AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado"))' +
        ';')
    Left = 340
    Top = 10
  end
  object DataSourcePassageirosTM_SIM: TDataSource
    DataSet = ADOQueryPassageirosTM_SIM
    Left = 900
    Top = 550
  end
  object ADOQueryPassageirosTM_NAO: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Dataprocura'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = '@Origem'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end
      item
        Name = '@Destino'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*, tblProgramacaoExecutante.*  '
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria '
      'WHERE ((tblProgramacaoDiaria.DataProgramacao LIKE @Dataprocura)'
      'AND(tblProgramacaoExecutante.Origem LIKE @Origem)'
      'AND(tblProgramacaoDiaria.txtDestino LIKE @Destino)'
      'AND(tblProgramacaoExecutante.StatusProgramacao LIKE "Aprovado")'
      
        'AND(tblProgramacaoExecutante.InseridoProgramacaoTransporte LIKE ' +
        'False));')
    Left = 300
    Top = 10
  end
  object DataSourcePassageirosTM_NAO: TDataSource
    DataSet = ADOQueryPassageirosTM_NAO
    Left = 940
    Top = 550
  end
  object DataSourceImportarLocalInstalacao: TDataSource
    DataSet = ADOQueryImportarLocalInstalacao
    Left = 780
    Top = 440
  end
  object ADOQueryImportarLocalInstalacao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryImportarLocalInstalacaoBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblLocalInstalacao.* '
      'FROM tblLocalInstalacao'
      'ORDER BY LocalInstalacao;')
    Left = 790
  end
  object DataSourceRECLocalSAP: TDataSource
    DataSet = ADOQueryRECLocalSAP
    Left = 740
    Top = 440
  end
  object ADOQueryRECLocalSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Plataforma'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@TAG'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblLocalInstalacao.*'
      'FROM tblLocalInstalacao'
      'WHERE ((Plataforma LIKE @Plataforma) AND'
      ' (DenominacaoLocalInstalacao LIKE @TAG));')
    Left = 750
  end
  object DataSourceForaFSMC: TDataSource
    DataSet = ADOQueryForaFSMC
    Left = 700
    Top = 440
  end
  object ADOQueryForaFSMC: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = 'DenominacaoLocalInstalacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblNotaManutencao.* FROM tblNotaManutencao'
      'WHERE (((TextoBreveNOTA LIKE "FO:%"))'
      'AND((FimAvaria IS NULL))'
      'AND((LocalInstalacao LIKE "%FSMC%"))'
      'AND((DenominacaoLocalInstalacao NOT LIKE "%Enquip%"))'
      'AND((DenominacaoLocalInstalacao NOT LIKE "%Apoio%")));')
    Left = 710
  end
  object DataSourceForaFSCI: TDataSource
    DataSet = ADOQueryForaFSCI
    Left = 770
    Top = 490
  end
  object ADOQueryForaFSCI: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.* FROM tblNotaManutencao'
      'WHERE (((TextoBreveNOTA LIKE "FO:%"))AND((FimAvaria IS NULL))AND'
      '((LocalInstalacao LIKE "%FSCI%")));')
    Left = 710
    Top = 50
  end
  object DataSourcePendenciaSAP: TDataSource
    DataSet = ADOQueryPendenciaSAP
    Left = 700
    Top = 550
  end
  object ADOQueryPendenciaSAP: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho;')
    Left = 750
    Top = 50
  end
  object ADOQueryTextoLongoOM: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@OrdemManutencao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblTextoLongoCarteira.* '
      'FROM tblTextoLongoCarteira '
      'WHERE (OrdemManutencao LIKE @OrdemManutencao);')
    Left = 710
    Top = 100
  end
  object DataSourceTextoLongoOM: TDataSource
    DataSet = ADOQueryTextoLongoOM
    Left = 660
    Top = 550
  end
  object DataSourceAnalisarTipoEtapaServico: TDataSource
    DataSet = ADOQueryAnalisarTipoEtapaServico
    Left = 620
    Top = 500
  end
  object ADOQueryAnalisarTipoEtapaServico: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho;')
    Left = 750
    Top = 100
  end
  object DataSourcePreventivaAtrasada: TDataSource
    DataSet = ADOQueryPreventivaAtrasada
    Left = 660
    Top = 440
  end
  object ADOQueryPreventivaAtrasada: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.*'
      'FROM tblCarteiraTrabalho'
      'WHERE (TextoBreveOM LIKE "RO:%")'
      'ORDER BY Plataforma, txtTipoEtapaServicoOM;')
    Left = 780
    Top = 50
  end
  object DataSourceConsultaExecutante_Documento_Data: TDataSource
    DataSet = ADOQueryConsultaExecutante_Documento_Data
    Left = 610
    Top = 550
  end
  object ADOQueryConsultaExecutante_Documento_Data: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CPF'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2019'
      end>
    SQL.Strings = (
      
        'SELECT tblProgramacaoExecutante.*, tblProgramacaoDiaria.DataProg' +
        'ramacao'
      
        'FROM tblProgramacaoDiaria INNER JOIN tblProgramacaoExecutante ON' +
        ' tblProgramacaoDiaria.idProgramacaoDiaria = tblProgramacaoExecut' +
        'ante.CodigoProgramacaoDiaria'
      
        'WHERE ((tblProgramacaoExecutante.Documento= @CPF)AND(tblPrograma' +
        'caoDiaria.DataProgramacao=@DataProgramacao));')
    Left = 370
    Top = 120
  end
  object ADOConnectionImportar: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\Us' +
      'ers\augus\Documents\Projetos\Colibri\Banco de dados\UO-SEAL-ATP-' +
      'OP-SM\bdColibri.colibri;Mode=Share Deny None;Persist Security In' +
      'fo=False;Jet OLEDB:System database="";Jet OLEDB:Registry Path=""' +
      ';Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLED' +
      'B:Database Locking Mode=1;Jet OLEDB:Global Partial Bulk Ops=2;Je' +
      't OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Passwo' +
      'rd="";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt D' +
      'atabase=False;Jet OLEDB:Don'#39't Copy Locale on Compact=False;Jet O' +
      'LEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 516
    Top = 254
  end
  object DataSourceProgramacaoDiaria_Importar: TDataSource
    DataSet = ADOQueryProgramacaoDiaria_Importar
    Left = 560
    Top = 620
  end
  object ADOQueryProgramacaoDiaria_Importar: TADOQuery
    Connection = ADOConnectionImportar
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria;')
    Left = 660
    Top = 220
  end
  object DataSourceProgramacaoExecutante_Importar: TDataSource
    DataSet = ADOQueryProgramacaoExecutante_Importar
    Left = 580
    Top = 500
  end
  object ADOQueryProgramacaoExecutante_Importar: TADOQuery
    Connection = ADOConnectionImportar
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@CodigoProgramacaoDiaria'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoExecutante.*'
      'FROM tblProgramacaoExecutante'
      'WHERE (CodigoProgramacaoDiaria=@CodigoProgramacaoDiaria)'
      'ORDER BY NomeExecutante;')
    Left = 1020
    Top = 220
  end
  object DataSourceProgramacaoServico_Importar: TDataSource
    DataSet = ADOQueryProgramacaoServico_Importar
    Left = 610
    Top = 440
  end
  object ADOQueryProgramacaoServico_Importar: TADOQuery
    Connection = ADOConnectionImportar
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@Servico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoServico.*'
      'FROM tblProgramacaoServico'
      'WHERE (CodigoProgramacaoDiaria=@Servico);')
    Left = 980
    Top = 220
  end
  object DataSourceTemporarioDBImportar: TDataSource
    DataSet = ADOQueryTemporarioDBImportar
    Left = 660
    Top = 500
  end
  object ADOQueryTemporarioDBImportar: TADOQuery
    Connection = ADOConnectionImportar
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoDiaria.*'
      'FROM tblProgramacaoDiaria;')
    Left = 1350
    Top = 220
  end
  object ADOConnectionConsulta: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\De' +
      'v\Projetos\Colibri\Banco de dados\UO-SEAL-ATP-OP-SM\dbConsulta.m' +
      'db;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB:Sy' +
      'stem database="";Jet OLEDB:Registry Path="";Jet OLEDB:Database P' +
      'assword="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mo' +
      'de=1;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk T' +
      'ransactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Creat' +
      'e System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLE' +
      'DB:Don'#39't Copy Locale on Compact=False;Jet OLEDB:Compact Without ' +
      'Replica Repair=False;Jet OLEDB:SFP=False;'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 508
    Top = 194
  end
  object DataSourceTemporarioDBMemoria: TDataSource
    DataSet = ADOQueryTemporarioDBMemoria
    Left = 220
    Top = 560
  end
  object ADOQueryTemporarioDBMemoria: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblLocalInstalacao.* '
      'FROM tblLocalInstalacao;')
    Left = 860
  end
  object ADOConnectionGantt: TADOConnection
    ConnectionString = 
      'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=C:\Us' +
      'ers\ZUCHI\Documents\Projetos\Colibri\Win32\Debug\Arquivos\dbGant' +
      't.mdb;Mode=Share Deny None;Persist Security Info=False;Jet OLEDB' +
      ':System database="";Jet OLEDB:Registry Path="";Jet OLEDB:Databas' +
      'e Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking' +
      ' Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bul' +
      'k Transactions=1;Jet OLEDB:New Database Password="";Jet OLEDB:Cr' +
      'eate System Database=False;Jet OLEDB:Encrypt Database=False;Jet ' +
      'OLEDB:Don'#39't Copy Locale on Compact=False;Jet OLEDB:Compact Witho' +
      'ut Replica Repair=False;Jet OLEDB:SFP=False;'
    KeepConnection = False
    LoginPrompt = False
    Provider = 'Microsoft.Jet.OLEDB.4.0'
    Left = 514
    Top = 322
  end
  object ADOQueryGantt: TADOQuery
    Connection = ADOConnectionGantt
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblGanttCarteira.*'
      'FROM tblGanttCarteira;')
    Left = 1290
    Top = 280
  end
  object DataSourceGantt: TDataSource
    DataSet = ADOQueryGantt
    Left = 510
    Top = 370
  end
  object DataSourceNotaManutencao: TDataSource
    DataSet = ADOQueryNotaManutencao
    Left = 1100
    Top = 390
  end
  object ADOQueryNotaManutencao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.*'
      'FROM tblNotaManutencao'
      'ORDER BY Plataforma;')
    Left = 830
    Top = 100
  end
  object DataSourceExecutante_TipoEtapaServico: TDataSource
    DataSet = ADOQueryExecutante_TipoEtapaServico
    Left = 50
    Top = 400
  end
  object ADOQueryExecutante_TipoEtapaServico: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = '@TipoEtapaServico'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '1'
      end>
    SQL.Strings = (
      'SELECT tblExecutante.*'
      'FROM tblExecutante'
      'WHERE ((txtTipoEtapaServico=@TipoEtapaServico)AND(Ativo = True))'
      'ORDER BY NomeExecutante;')
    Left = 1540
    Top = 100
  end
  object DataSourceImportarTextoLongoOperacao: TDataSource
    DataSet = ADOQueryImportarTextoLongoOperacao
    Left = 1090
    Top = 610
  end
  object ADOQueryImportarTextoLongoOperacao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblTextoLongoOperacao.*'
      'FROM tblTextoLongoOperacao;')
    Left = 1110
    Top = 50
  end
  object DataSourceAnaliseGM: TDataSource
    DataSet = ADOQueryAnaliseGM
    Left = 790
    Top = 620
  end
  object ADOQueryAnaliseGM: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.* '
      'FROM tblNotaManutencao'
      'ORDER BY Plataforma;')
    Left = 950
    Top = 100
  end
  object DataSourceRestricaoFSMC: TDataSource
    DataSet = ADOQueryRestricaoFSMC
    Left = 910
    Top = 620
  end
  object ADOQueryRestricaoFSMC: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCarteiraTrabalho.* FROM tblCarteiraTrabalho'
      'WHERE ((TextoBreveOM LIKE "RO:%")'
      'AND(LocalInstalacao LIKE "%FSMC%")'
      'AND(StatusUsuarioOM NOT LIKE "%AENT%")'
      'AND(DenominacaoLocalInstalacao NOT LIKE "%Enquip%")'
      'AND(DenominacaoLocalInstalacao NOT LIKE "%Apoio%"));')
    Left = 680
  end
  object DataSourceMovimentacaoCarga: TDataSource
    DataSet = ADOQueryMovimentacaoCarga
    Left = 1170
    Top = 610
  end
  object ADOQueryMovimentacaoCarga: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryMovimentacaoCargaBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblMovimentacaoCarga.*'
      'FROM tblMovimentacaoCarga'
      'ORDER BY DataNecessidade;')
    Left = 1580
    Top = 150
  end
  object DataSourceCondicaoMar: TDataSource
    DataSet = ADOQueryCondicaoMar
    Left = 1040
    Top = 520
  end
  object ADOQueryCondicaoMar: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryCondicaoMarBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCondicaoMar.*'
      'FROM tblCondicaoMar'
      'ORDER BY DataCondicaoMar;')
    Left = 860
    Top = 280
  end
  object DataSourceCondicaoEmbarcacao: TDataSource
    DataSet = ADOQueryCondicaoEmbarcacao
    Left = 1000
    Top = 520
  end
  object ADOQueryCondicaoEmbarcacao: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryCondicaoEmbarcacaoBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblCondicaoEmbarcacao.*'
      'FROM tblCondicaoEmbarcacao'
      'ORDER BY DataCondicaoEmbarcacao;')
    Left = 820
    Top = 280
  end
  object DataSourceProgramacaoNotas: TDataSource
    DataSet = ADOQueryProgramacaoNotas
    Left = 1220
    Top = 460
  end
  object ADOQueryProgramacaoNotas: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryProgramacaoNotasBeforePost
    Parameters = <
      item
        Name = '@DataProgramacao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = '01/01/2019'
      end>
    SQL.Strings = (
      'SELECT tblProgramacaoNotas.* '
      'FROM tblProgramacaoNotas'
      'WHERE (DataProgramacao LIKE @DataProgramacao);')
    Left = 900
    Top = 280
  end
  object DataSourceImportarItemManutencao: TDataSource
    DataSet = ADOQueryImportarItemManutencao
    Left = 1100
    Top = 620
  end
  object ADOQueryImportarItemManutencao: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblItemManutencao.* '
      'FROM tblItemManutencao'
      'ORDER BY DescricaoItem;')
    Left = 1110
  end
  object ADOQueryContadorSolicitacao: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = 'Descricao'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT tblContadorSolicitacao.*'
      'FROM tblContadorSolicitacao'
      'ORDER BY Descricao;')
    Left = 860
    Top = 220
  end
  object DataSourceContadorSolicitacao: TDataSource
    DataSet = ADOQueryContadorSolicitacao
    Left = 870
    Top = 370
  end
  object DataSourceProgramacaoCalendario: TDataSource
    DataSet = ADOQueryProgramacaoCalendario1
    Left = 1220
    Top = 350
  end
  object ADOQueryProgramacaoCalendario1: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoCalendario.*'
      'FROM tblProgramacaoCalendario'
      'ORDER BY DataProgramacao;')
    Left = 1270
    Top = 220
  end
  object ADOQueryPlataformaControle: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    BeforePost = ADOQueryPlataformaControleBeforePost
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPlataforma.*'
      'FROM tblPlataforma'
      'WHERE (BooleanPlataforma = True) '
      'ORDER BY Plataforma;')
    Left = 1500
    Top = 100
  end
  object DataSourcePlataformaControle: TDataSource
    DataSet = ADOQueryPlataformaControle
    Left = 1260
    Top = 610
  end
  object DataSourceImportarMemoria: TDataSource
    DataSet = ADOQueryImportarMemoria
    Left = 1140
    Top = 450
  end
  object ADOQueryImportarMemoria: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblNotaManutencao.*'
      'FROM tblNotaManutencao;')
    Left = 1150
    Top = 50
  end
  object DataSourceImportarICPM: TDataSource
    DataSet = ADOQueryImportarICPM
    Left = 1300
    Top = 610
  end
  object ADOQueryImportarICPM: TADOQuery
    Connection = ADOConnectionMemoria
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblICPM.* '
      'FROM tblICPM;')
    Left = 1150
  end
  object ADOQueryPalavraChave: TADOQuery
    Connection = ADOConnectionConsulta
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblPalavraChave.*'
      'FROM tblPalavraChave;')
    Left = 1170
    Top = 280
  end
  object DataSourcePalavraChave: TDataSource
    DataSet = ADOQueryPalavraChave
    Left = 1180
    Top = 460
  end
  object ADOQueryAuxTabelaRT: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = 'idPProgramacaoRT'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT  tblProgramacaoRT.* FROM  tblProgramacaoRT'
      'WHERE idPProgramacaoRT = 1;')
    Left = 218
    Top = 216
  end
  object DataSourceAuxTabelaRT: TDataSource
    DataSet = ADOQueryAuxTabelaRT
    Left = 210
    Top = 288
  end
  object ADOQueryCancelarRT: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <
      item
        Name = 'idPProgramacaoRT'
        Attributes = [paNullable]
        DataType = ftWideString
        NumericScale = 255
        Precision = 255
        Size = 510
        Value = Null
      end>
    SQL.Strings = (
      'SELECT  tblProgramacaoRT.* FROM  tblProgramacaoRT'
      'WHERE idPProgramacaoRT = 1;')
    Left = 266
    Top = 216
  end
  object DataSourceCancelarRT: TDataSource
    DataSet = ADOQueryCancelarRT
    Left = 258
    Top = 288
  end
  object ADOQueryGestaoRT: TADOQuery
    Connection = ADOConnectionColibri
    CursorType = ctStatic
    LockType = ltPessimistic
    Parameters = <>
    SQL.Strings = (
      'SELECT tblProgramacaoRT.* FROM tblProgramacaoRT;')
    Left = 314
    Top = 216
  end
  object DataSourceGestaoRT: TDataSource
    DataSet = ADOQueryGestaoRT
    Left = 306
    Top = 288
  end
end
