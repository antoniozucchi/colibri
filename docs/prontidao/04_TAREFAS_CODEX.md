# Tarefas para Codex - Módulo Prontidão

## Objetivo inicial

Implementar no Colibri a importação de arquivo TXT exportado do WhatsApp para alimentar o módulo de prontidão no banco dbProntidao.mdb.

## Componentes importantes

Banco:
- dbProntidao.mdb

Conexão:
- FrmDataModule.ADOConnectionProntidao

Tabelas já criadas:
- tblWAImportacao
- tblWAMensagem
- tblProntidaoProdutividadeTurno
- tblProntidaoEquipeTurno
- tblProntidaoDeslocamentoEfetivo
- tblProntidaoDeslocamentoEquipe
- tblProntidaoDesvioTurno
- tblFuncaoProdutividade
- tblAliasFuncaoProdutividade
- tblTipoDesvioProdutividade
- tblProntidaoPremissaCampanha

## Primeira entrega desejada

Criar uma unit Delphi responsável por importar TXT do WhatsApp.

Sugestão de unit:
uProntidaoWhatsAppImporter.pas

Funções esperadas:
- Ler arquivo TXT.
- Identificar mensagens no formato WhatsApp.
- Agrupar mensagens com múltiplas linhas.
- Extrair DataHoraMensagem.
- Extrair Remetente.
- Extrair TextoOriginal.
- Gerar TextoNormalizado.
- Classificar TipoMensagem:
  - Sistema
  - Operacional
  - Imagem
  - Documento
  - Apagada
- Gerar HashMensagem para evitar duplicidade.
- Gravar em tblWAMensagem.
- Criar registro em tblWAImportacao.
- Atualizar quantidade de mensagens importadas.

## Formatos de mensagem a aceitar

Formato 1:
[10/03/2026, 10:29:55] Daniel Camarão Petrobras: Bom dia a todos

Formato 2:
10/03/2026, 10:29 - Daniel Camarão Petrobras: Bom dia a todos

Formato 3:
10/03/2026 10:29 - Daniel Camarão Petrobras: Bom dia a todos

Uma mensagem pode ter múltiplas linhas. Linhas que não começam com data/hora devem ser anexadas ao texto da mensagem anterior.

## Não processar como operacional

Ignorar ou classificar como Sistema:
- mensagens e ligações são protegidas
- criou este grupo
- adicionou
- removeu
- saiu
- imagem ocultada
- documento omitido
- mensagem apagada

## Segunda entrega desejada

Criar parser operacional para identificar:
- campanha de prontidão;
- data de referência;
- turno DIURNO/NOTURNO;
- plataforma base;
- plataforma serviço;
- origem;
- embarcação;
- horário de transbordo;
- hora de chegada;
- hora de início das atividades;
- hora de parada;
- hora de saída;
- distribuição das equipes;
- deslocamentos;
- desvios.

## Terceira entrega desejada

Criar tela de revisão manual no Colibri para validar os dados extraídos antes de consolidar.

## Cuidados

Não alterar tabelas existentes do Colibri sem aprovação.
Não remover campos existentes.
Não usar Nz em SQL/Access.
Preferir funções Delphi para tratar nulos.
Usar ADOConnectionProntidao para o banco dbProntidao.
Manter texto original do WhatsApp para rastreabilidade.