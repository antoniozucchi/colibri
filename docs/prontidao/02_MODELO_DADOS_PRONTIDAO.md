# Modelo de Dados - Prontidão

As tabelas já foram criadas em tempo de execução no banco dbProntidao.mdb.

Tabelas principais:

- tblProntidaoCampanha
- tblProntidaoPremissaCampanha
- tblWAImportacao
- tblWAMensagem
- tblProntidaoProdutividadeTurno
- tblProntidaoEquipeTurno
- tblProntidaoDeslocamentoEfetivo
- tblProntidaoDeslocamentoEquipe
- tblProntidaoDesvioTurno
- tblTipoDesvioProdutividade
- tblFuncaoProdutividade
- tblAliasFuncaoProdutividade
- tblDeParaFuncaoColibriProdutividade
- tblProntidaoConfrontoEquipe

Padrão de tipos usado no Access:

- VARCHAR(x): textos curtos
- MEMO: textos longos, mensagens do WhatsApp, observações e textos fonte
- LONG: IDs e quantidades inteiras
- DOUBLE: horas decimais e indicadores
- DATETIME: datas e horários
- YESNO: campos booleanos

Cada linha da produtividade representa:

1 campanha + 1 data de referência + 1 turno + 1 origem + 1 plataforma de serviço

Se o SOV trabalha em dois turnos no mesmo dia, devem existir duas linhas:
- uma para DIURNO;
- outra para NOTURNO.

PlataformaBase é a plataforma da campanha/grupo.
PlataformaServico é a plataforma onde o serviço realmente ocorreu.

Exemplo:
- PlataformaBase = PCM-01
- PlataformaServico = PCM-04

Isso significa que o efetivo saiu da campanha da PCM-01 para executar atividade em outra plataforma.