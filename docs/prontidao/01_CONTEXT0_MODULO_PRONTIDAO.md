# Contexto do Módulo de Prontidão do Colibri

O Colibri é uma aplicação Delphi 12 VCL com banco de dados Microsoft Access usando ADO.

Foi criado um banco exclusivo para o módulo de prontidão chamado dbProntidao.mdb, acessado pelo componente:

FrmDataModule.ADOConnectionProntidao

O objetivo do módulo é substituir a alimentação manual da planilha "ANEXO G - HORAS EFETIVAS REALIZADAS A BORDO.xlsx", utilizando dados extraídos de arquivos TXT exportados de grupos do WhatsApp.

O fluxo desejado é:

1. Importar arquivo TXT do WhatsApp.
2. Separar mensagens individuais.
3. Ignorar mensagens administrativas do grupo, imagens ocultadas, documentos omitidos e mensagens apagadas.
4. Identificar mensagens operacionais de prontidão.
5. Extrair dados de campanha, data, turno, plataforma, origem, embarcação, horários, equipe, desvios e deslocamentos.
6. Gravar o dado bruto para rastreabilidade.
7. Permitir revisão manual no Colibri.
8. Consolidar os dados oficiais de produtividade.
9. Confrontar o efetivo informado no WhatsApp com a programação de executantes do Colibri.
10. Exportar a base tratada para Excel/Python ou gerar o Anexo G como relatório.

O módulo deve preservar sempre o texto original do WhatsApp, para auditoria.