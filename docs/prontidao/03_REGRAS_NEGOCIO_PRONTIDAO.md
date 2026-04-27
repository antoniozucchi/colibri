# Regras de Negócio - Prontidão

## Fonte dos dados

Colibri = planejado.
WhatsApp = realizado informado pela fiscalização.
Revisão manual = dado oficial consolidado.

## Turnos

Cada turno deve gerar uma linha de produtividade.

Turnos principais:
- DIURNO
- NOTURNO

O turno noturno pode iniciar em uma data e terminar no dia seguinte. O cálculo de horas deve tratar corretamente passagem pela meia-noite.

## Origem e embarcação

Origem é diferente de embarcação.

Exemplo:
- SOV encostado na PCM-01.
- Parte do efetivo sai da PCM-01 para PCM-04 usando SURFER.
- Origem pode ser PCM-01 ou SOV.
- Embarcação de deslocamento pode ser SURFER.

Guardar:
- OrigemTipo
- OrigemCodigo
- EmbarcacaoTipo
- EmbarcacaoNome

## Deslocamento de efetivo

Quando parte do efetivo da campanha vai trabalhar em outra plataforma, esse efetivo não deve compor a produtividade da plataforma base da prontidão.

Deve ser registrado:
- destino;
- embarcação usada;
- quantidade deslocada por função;
- motivo;
- texto fonte do WhatsApp.

## Premissas de campanha

A campanha possui premissas por período de vigência.

Cenários possíveis:
- 1 turno com máximo de 15 PAX por dia. Parâmetro direto = 13.
- 1 turno com máximo de 25 PAX por dia. Parâmetro direto = 23.
- 2 turnos com máximo de 15 PAX por turno. Parâmetro direto = 13.
- 2 turnos com máximo de 25 PAX por turno. Parâmetro direto = 23.

As premissas podem mudar ao longo da campanha. Por isso a tabela tblProntidaoPremissaCampanha possui DataInicioVigencia e DataFimVigencia.

HEP por turno vem do arquivo "ANEXO B - MODELAGEM DE HORAS EFETIVAS A BORDO.xlsx" e deve ser cadastrado como premissa da campanha.

## POB direto e indireto

Para cálculo de déficit de POB, considerar como indiretos somente:
- Supervisão
- Fiscal Caldeiraria/Pintura/Hotelaria

Os demais devem contar como POB direto, salvo ajuste posterior no cadastro tblFuncaoProdutividade.

DeficitPOBDireto = POBDiretoReferencia - QtdPOBDiretoInformado

Se a diferença for negativa, o déficit deve ser zero, mas a diferença real deve ser preservada em DiferencaPOBDireto.

## Tempos mortos

Tempos Mortos:
reduzem a janela de horas efetivas.

Tempos Mortos Intra-Efetiva:
a equipe estava dentro da janela efetiva, mas sem produzir plenamente.

## Confronto com Colibri

O confronto inicial deve ser quantitativo por função/etapa, não nominal.

Chave provável:
- DataReferencia
- PlataformaServico
- Turno
- TipoEtapaServico
- FuncaoNormalizada

A programação do Colibri será comparada futuramente com base na tblProgramacaoExecutante, usando um de/para entre txtFuncao e TipoEtapaServico.