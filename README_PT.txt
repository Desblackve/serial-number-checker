Criador: Desblack "Dess"
Contato: desblackve@gmail.com


SERIAL NUMBER CHECKER
=====================

Versão 2.1A


1. O QUE O APLICATIVO FAZ
-------------------------

O Serial Number Checker tem dois modos (botões no canto superior
esquerdo):

- CHECKER: carrega uma lista de números de série (S/N) de um arquivo CSV
  ou Excel e verifica, escaneando ou digitando, quais estão presentes,
  quais faltam e quais sobram. No final, exporta um relatório Excel
  colorido.
- CRIAR LISTA S/N: monta uma lista nova com o que você escaneia,
  organizada por pallets (em colunas ou planilhas) ou, com o Modo Rack,
  pelas posições de um rack. Exporta para Excel ou CSV.

A interface está em espanhol, inglês e português. O botão README (no
canto superior direito) abre este guia dentro do aplicativo.


2. EXEMPLOS RÁPIDOS
-------------------

EXEMPLO 1 - Conferir um inventário com um arquivo

  1. Clique em "Carregar CSV / Excel" e escolha o arquivo (por exemplo,
     inventario_outubro.xlsx).
  2. Escaneie os equipamentos um por um. Cada leitura mostra:
       Verde      ENCONTRADO        (Planilha Lote1 · linha 7)
       Amarelo    JÁ ESCANEADO      (o contador não sobe)
       Vermelho   NÃO ENCONTRADO    (salvo em Sobras)
  3. Veja o contador: "128 / 340" e "38% concluído · 212 pendentes".
  4. Clique em "Exportar relatório": é criado
     inventario_outubro_resultado.xlsx com os S/N em verde (escaneado),
     vermelho (pendente) e azul (duplicado), mais as planilhas Resumo,
     Leituras e Sobras.

EXEMPLO 2 - Escaneei o código errado

  Se escanear o código errado, pressione Ctrl+Z (ou o botão "Desfazer"
  de Últimas leituras): a última leitura é anulada e o contador volta ao
  que era.

EXEMPLO 3 - Tenho os S/N em um e-mail ou em outra planilha do Excel

  Copie a lista (um por linha) e cole com Ctrl+V no campo de leitura. O
  aplicativo processa todos de uma vez e mostra um resumo:
  "Colados 25: 22 encontrados · 2 não encontrados · 1 repetidos".

EXEMPLO 4 - Montar uma lista de pallets

  1. Mude para o modo "Criar lista S/N".
  2. Escaneie os S/N do primeiro pallet: eles são salvos na coluna A
     (Pallet 1).
  3. Clique em "Próximo pallet" (ou F2): os seguintes vão para a coluna
     B (Pallet 2), e assim por diante.
  4. Se quiser um pallet em uma planilha separada, clique em "Nova
     planilha" ao lado das abas de planilhas.
  5. Clique em "Exportar lista" (ou Ctrl+E) para salvá-la em Excel ou
     CSV.

EXEMPLO 5 - Preencher um rack posição por posição (Modo Rack)

  1. Em "Criar lista S/N", marque "Modo Rack".
  2. Digite a localização (por exemplo, DC05), as linhas e colunas do
     rack (por exemplo, 6 e 8) e a ordem: Horizontal ou Vertical.
  3. Clique em "Adicionar novo rack": é criado o DC05-1.
  4. Escaneie: o contador mostra a próxima posição, por exemplo (0,0), e
     cada S/N é salvo ali. Depois (1,0), (2,0)...
  5. Quando o rack ficar cheio, o aplicativo avisa "O rack DC05-1 foi
     completado". Clique em "Adicionar novo rack" para seguir com o
     DC05-2.

EXEMPLO 6 - Conferir se cada S/N vai no pallet certo

  1. Em "Criar lista S/N", marque "Validar lista" e clique em "Carregar
     lista" para escolher o CSV ou Excel com os S/N que devem ir nos
     pallets.
  2. Escaneie normalmente. Se um S/N não estiver nessa lista, ele não é
     registrado e aparece em vermelho "S/N NÃO VÁLIDO PARA O PALLET".
  3. O contador mostra quantos já foram e quantos faltam, por exemplo
     "Validados 45 de 60 · faltam 15".

EXEMPLO 7 - Ver o avanço de longe

  Pressione F11 (ou "Visão ampliada"): a tela mostra apenas o contador e
  o resultado em tamanho grande. F11 ou Esc para voltar.


3. NOVIDADES
------------

VERSÃO 2.1A - Validar lista e nomes de planilha

- Validar lista (Criar lista): só são registrados os S/N de uma lista
  carregada; os outros recebem um aviso em vermelho e não entram.
- Com Validar lista também é possível usar "Alternativas A/B/U".
- A primeira planilha se chama "Pallet 1" e qualquer planilha pode ser
  renomeada com um clique duplo na aba.
- Uma descrição curta aparece ao passar o mouse sobre "Modo Rack",
  "Validar lista" e as abas de planilhas.
- Ao exportar para Excel, as colunas saem com a largura exata para o S/N
  mais longo.
- Checker: copiar o S/N pendente com Ctrl+C ou com o botão direito →
  "Copiar S/N".

VERSÃO 2.0 - Design novo e Modo Rack

- Design novo: cartões escuros sobre um fundo de linhas, botões com
  ícones e barra de título escura no Windows.
- Executável para Windows (SerialNumberChecker.exe): não é mais preciso
  instalar o Python.
- Botão README: este guia pode ser lido dentro do aplicativo.
- Modo Rack: lista guiada pelas posições (x,y) de cada rack.
- Sobras: os S/N que não estão no arquivo ficam em uma aba própria e são
  exportados em uma planilha separada.
- Desfazer a última leitura (Ctrl+Z ou botão "Desfazer").
- Planilha e linha: ao encontrar um S/N, o aplicativo indica onde ele
  está no arquivo.
- Filtro de formato: se for escaneado outro código da caixa (por
  exemplo, um de 13 caracteres), avisa "NÃO PARECE UM S/N" e não o
  conta.
- Colar vários S/N de uma vez com Ctrl+V.
- Contador com barra de progresso, ritmo (S/N por minuto) e tempo
  estimado para terminar.
- Abas Pendentes / Sobras com busca.
- Visão ampliada (F11).
- O relatório inclui as planilhas Resumo, Leituras (hora de cada
  leitura) e Sobras.
- Imprimir (Ctrl+P): resumo, pendentes e sobras (Checker) ou a planilha
  atual / todas as planilhas (Criar lista).
- Criar lista: botão "Nova planilha", colunas que podem ser esticadas e
  tecla Backspace para excluir o registro selecionado.
- Mais cabeçalhos de S/N são reconhecidos (ver seção 15).
- A sessão é salva sozinha: ao abrir o aplicativo de novo, você continua
  de onde parou.


4. MODO CHECKER, PASSO A PASSO
------------------------------

  1. Abra o Serial Number Checker.
  2. Se precisar, mude o idioma no seletor do canto superior direito
     (Español, English ou Português). A mudança é imediata.
  3. Clique em "Carregar CSV / Excel" e selecione o arquivo com os S/N.
  4. São mostrados o nome do arquivo, os valores lidos e os S/N únicos.
  5. Se o arquivo for um relatório exportado antes, o aplicativo pergunta
     se você quer continuar de onde parou (seção 9).
  6. Escaneie ou digite um S/N no campo de leitura.
  7. Ao chegar a 17 caracteres, a verificação é automática. Também é
     possível pressionar ENTER.
  8. Repita com todos os S/N que quiser verificar.
  9. No final, clique em "Exportar relatório".
  10. Para começar do zero, use o botão "Reiniciar" (seção 16).

A caixa "Alternativas A/B/U" da barra superior liga ou desliga a busca
alternativa da seção 5.


5. REGRAS DE CORRESPONDÊNCIA
----------------------------

- Primeiro é procurada uma correspondência exata com um S/N carregado.
- Se não houver e "Alternativas A/B/U" estiver marcada, o sétimo
  caractere é substituído por A, B e U. Se uma alternativa corresponder,
  é marcado como encontrado o S/N que realmente aparece na lista (a
  mensagem diz "variante" e o que foi lido).
- Um S/N já encontrado não soma de novo no contador.
- Se o código lido não tiver o tamanho de nenhum S/N do arquivo, o
  aplicativo avisa "NÃO PARECE UM S/N" (provavelmente foi escaneado
  outro código da caixa) e ele não é salvo como sobra.

IMPORTANTE: a substituição do sétimo caractere é uma busca alternativa.
Não garante que o número de série seja autêntico nem é uma validação
completa do formato.


6. CORES, SONS E MENSAGENS
--------------------------

No aplicativo:

- VERDE: S/N encontrado. Som de confirmação. A mensagem indica se foi
  exato ou uma variante, e em qual planilha e linha do arquivo ele está.
- AMARELO: S/N já escaneado. Outro som. O contador não sobe e o
  aplicativo fica amarelo até a próxima leitura ou até pressionar Esc.
- VERMELHO: S/N não encontrado (salvo em Sobras) ou código que não
  parece um S/N. Som de erro.
- A borda do campo de leitura pisca na cor do resultado.
- Abaixo do campo de leitura há uma legenda com as cores do relatório.

No relatório Excel:

- VERDE: S/N encontrado durante a conferência.
- VERMELHO: S/N ainda não encontrado.
- AZUL: S/N duplicado no arquivo. A primeira ocorrência mantém o verde
  ou vermelho e as repetições (em qualquer planilha) ficam em azul.

Se foi encontrado por uma alternativa A/B/U, é colorido o S/N real da
lista.


7. CONTADOR, PROGRESSO E RITMO
------------------------------

O contador mostra os S/N encontrados em relação ao total de S/N únicos:

    128 / 340
    38% concluído · 212 pendentes
    12/min · faltam ~18 min

- Um S/N repetido no arquivo conta uma única vez.
- Ao carregar um arquivo, o aplicativo mostra quantos valores foram
  lidos e quantos são únicos. Exemplo: "12 valores lidos · 10 S/N
  únicos" significa que há 2 valores repetidos no arquivo.
- O ritmo e o tempo estimado aparecem após algumas leituras e são
  calculados com os últimos 10 minutos.


8. PENDENTES, SOBRAS E HISTÓRICO
--------------------------------

- ÚLTIMAS LEITURAS: os 5 últimos S/N lidos e seu resultado. O botão
  "Desfazer" (ou Ctrl+Z) anula a última leitura; pode ser repetido.
- PENDENTES: os S/N do arquivo que ainda faltam escanear.
- SOBRAS: os S/N escaneados que não estão no arquivo, com quantas vezes
  foram lidos e a hora da primeira leitura.
- COPIAR: selecione um S/N em Pendentes ou Sobras e pressione Ctrl+C, ou
  clique nele com o botão direito e escolha "Copiar S/N".
- BUSCAR: filtra a aba aberta enquanto você digita. Enter ou Esc
  devolvem o cursor ao campo de leitura.


9. CONTINUAR UM RELATÓRIO EXPORTADO
-----------------------------------

Se você carregar um relatório XLSX exportado antes pelo aplicativo, pode
continuar de onde parou.

- O aplicativo detecta os S/N em verde do relatório e pergunta no aviso
  amarelo se você quer continuar.
- "Continuar": esses S/N são marcados como encontrados. "Começar do
  zero": o arquivo é carregado sem progresso.
- As planilhas Resumo, Leituras e Sobras do relatório são ignoradas na
  leitura e substituídas ao exportar de novo.
- Arquivos CSV não podem ser continuados porque não guardam cores.

NOTA: é reconhecido o verde exato do relatório. Se um arquivo tiver
células com esse mesmo verde por outro motivo (por exemplo, o estilo
"Bom" do Excel), também vai perguntar; nesse caso escolha "Começar do
zero".


10. EXPORTAR RELATÓRIO
----------------------

"Exportar relatório" (ou Ctrl+E) salva um Excel (.xlsx) com o
resultado.

- Se o arquivo carregado for XLSX, o relatório mantém suas planilhas.
- Se for CSV, é criado um XLSX para poder colorir as células.
- As colunas de S/N são coloridas de verde, vermelho ou azul (seção 6).
- As colunas de S/N saem com a largura exata para o S/N mais longo.
- São adicionadas três planilhas:
    Resumo       arquivo, totais, encontrados, pendentes, sobras e as
                 horas da primeira e da última leitura.
    Leituras     cada leitura com sua hora, o que foi lido e o resultado.
    Sobras       os S/N que não estão no arquivo (se houver).
- O arquivo original não é modificado: o relatório é salvo à parte e não
  é permitido salvá-lo por cima do original.
- Exceção: se você estiver continuando um relatório (seção 9), pode
  salvar o progresso sobre esse mesmo relatório. O Windows pedirá para
  confirmar que deseja substituí-lo.
- O nome sugerido é nome_do_arquivo_resultado.xlsx.


11. MODO CRIAR LISTA S/N
------------------------

Serve para montar uma lista nova com o que você escaneia. Aceita
qualquer texto (sem mudar maiúsculas nem limitar o tamanho).

- DESTINO do próximo pallet:
    Nova coluna        cada pallet na coluna seguinte (padrão).
    Nova planilha      cada pallet em sua própria planilha.
    Da célula          começa na célula selecionada no visor.
- NOME (opcional): título do pallet ou da planilha. Se ficar vazio, é
  usado "Pallet 1", "Pallet 2"...
- "Próximo pallet" (F2): começa o pallet seguinte.
- "Nova planilha" (ao lado das abas de planilhas): cria uma planilha
  nova e vai para ela.
- NOME DAS PLANILHAS: a primeira se chama "Pallet 1". Clique duas vezes
  na aba de qualquer planilha para renomeá-la (Enter para salvar, Esc
  para cancelar). O nome novo é usado ao exportar e imprimir.
- VALIDAR LISTA: confere se cada S/N está em uma lista carregada (seção
  12).
- VISOR: mostra a planilha exatamente como será exportada. Clique para
  selecionar uma célula, clique duas vezes para editá-la. As colunas são
  esticadas arrastando sua borda na linha de letras; clique duas vezes na
  borda para ajustá-las ao texto.
- DUPLICADOS: entram em azul e o aplicativo fica amarelo com um aviso:
  "Excluir duplicados" ou "Ocultar aviso".
- EXCLUIR: Backspace ou Del com o campo vazio, ou o botão "Excluir",
  apaga o registro selecionado. "Desfazer" (Ctrl+Z) o recupera.
- EXPORTAR LISTA (Ctrl+E): para Excel (com as mesmas planilhas e cores
  do visor) ou para CSV (S/N, pallet, planilha, célula e se é
  duplicado). No Excel as colunas saem com a largura exata para o S/N
  mais longo; um CSV não guarda larguras de coluna.
- Um valor que começa com "=" é salvo como texto, não como fórmula.


12. VALIDAR LISTA
-----------------

Para garantir que nos pallets só entrem os S/N corretos. É ativado com a
caixa "Validar lista" na barra superior do modo Criar lista.

- CARREGAR LISTA: escolha um CSV ou Excel com os S/N esperados (são
  reconhecidos os mesmos cabeçalhos do Checker, seção 15).
- SEM LISTA: enquanto não carregar uma, o aplicativo mostra um aviso
  amarelo e não registra nada do que for escaneado.
- S/N VÁLIDO: entra no pallet normalmente.
- S/N DUPLICADO: entra em azul, como sempre.
- S/N QUE NÃO ESTÁ NA LISTA: não é registrado e aparece em vermelho "S/N
  NÃO VÁLIDO PARA O PALLET", com o som de erro.
- FALTAM: o contador mostra "Validados X de Y · faltam Z".
- ALTERNATIVAS A/B/U: com Validar lista ativo, esta caixa também aparece,
  com a mesma regra do Checker (seção 5).
- MODO RACK: o Modo Rack só coloca máquinas e não valida, então uma
  opção bloqueia a outra (cinza): com o Modo Rack ativo não é possível
  marcar "Validar lista", e com Validar lista ativo não é possível marcar
  "Modo Rack".


13. MODO RACK
-------------

Para guardar equipamentos em racks, posição por posição. É ativado com a
caixa "Modo Rack" na barra superior do modo Criar lista.

- LOCALIZAÇÃO: o lugar físico, por exemplo DC05. Os racks se chamam
  DC05-1, DC05-2, DC05-3...
- LINHAS e COLUNAS: tamanho do rack (6 linhas e 8 colunas por padrão).
- COORDENADAS (x,y): x é a coluna e y a linha, começando em (0,0) no
  canto superior esquerdo. O visor numera a partir de 0.
- ORDEM: Horizontal preenche (0,0), (1,0), (2,0)... linha por linha.
  Vertical preenche (0,0), (0,1), (0,2)... coluna por coluna.
- "ADICIONAR NOVO RACK" (ou F2, ou ao lado das abas): cria o rack
  seguinte da localização e começa em (0,0).
- GUIA: o contador mostra a próxima posição em tamanho grande e o visor
  a marca com um quadro verde. Cada leitura é salva ali e avança.
- PULAR: clique em uma posição vazia do visor e a próxima leitura irá
  para lá.
- RACK COMPLETO: ao preencher a última posição, o aplicativo avisa "O
  rack DC05-1 foi completado" e não registra mais nada nesse rack até
  você adicionar um novo.
- EXPORTAR: uma planilha por rack desenhada como a estante (coordenadas
  x em cima e y à esquerda) e uma planilha "Lista" com Localização,
  Rack, Posição (x,y), X, Y, S/N e se é duplicado.


14. IMPRIMIR E VISÃO AMPLIADA
-----------------------------

IMPRIMIR (botão "Imprimir" ou Ctrl+P): o aplicativo pergunta no aviso
amarelo o que imprimir.

- Checker: marque Resumo, Pendentes e/ou Sobras e clique em "Imprimir".
- Criar lista: "Planilha atual" ou "Todas as planilhas".
- A impressão sai na impressora padrão do Windows. Se o computador tiver
  Excel, imprime com o Excel; se não, com o Bloco de Notas.

VISÃO AMPLIADA (F11): apenas o contador e o último resultado em tamanho
grande, para vê-los de longe. É possível continuar escaneando. F11 ou
Esc para voltar.


15. ARQUIVOS COMPATÍVEIS E CABEÇALHOS
-------------------------------------

Formatos:

- XLSX: arquivos do Excel. Todas as planilhas da pasta de trabalho são
  lidas.
- CSV: uma única tabela. O separador (vírgula, ponto e vírgula,
  tabulação ou |) é detectado sozinho.

Cabeçalhos reconhecidos como coluna de S/N (sem importar maiúsculas,
acentos, espaços ou sinais):

    SN, S/N, NS, N/S
    SERIAL, SERIAL NO, SERIAL NR, SERIAL NBR, SERIAL NUM, SERIAL NUMBER
    SERIE, N SERIE, Nº DE SERIE, NO SERIE, NO. DE SERIE
    NUM SERIE, NUM. DE SERIE, NÚMERO SERIE, NÚMERO DE SERIE
    N SERIAL, NÚMERO SERIAL

Regras de leitura:

- Se houver vários cabeçalhos de S/N (em colunas ou planilhas
  diferentes), todos são lidos.
- As colunas que não são de S/N não são usadas.
- Células vazias e valores NONE ou NAN são ignorados.
- Se uma célula tiver uma fórmula, é usado o valor calculado salvo pelo
  Excel.
- Os números são lidos sem decimais (123456.0 é lido como 123456).


16. ATALHOS, REINÍCIO E SESSÃO
------------------------------

Atalhos de teclado:

    Enter        verificar / registrar
    Ctrl+Z       desfazer
    Ctrl+E       exportar
    Ctrl+P       imprimir
    Ctrl+V       colar vários S/N
    Ctrl+C       copiar o S/N selecionado em Pendentes ou Sobras
    F11          visão ampliada
    Esc          ocultar o aviso / sair da visão ampliada
    F2           próximo pallet ou novo rack (Criar lista)
    Backspace    excluir o registro selecionado (Criar lista, com o campo
                 vazio). Del faz o mesmo.

REINICIAR: o botão "Reiniciar" (canto inferior direito) apaga a sessão
do modo ativo. Se houver dados, o primeiro clique mostra "Certeza?" e é
preciso clicar de novo em até 3 segundos. O arquivo original não é
apagado do seu computador.

SESSÃO: o progresso é salvo sozinho. Se você fechar o aplicativo, ao
abri-lo recupera o arquivo carregado, o que foi escaneado, as sobras e a
lista ou os racks em criação. Mesmo assim, exporte o relatório antes de
reiniciar.


17. REQUISITOS E INSTALAÇÃO
---------------------------

EXECUTÁVEL (recomendado):

- Descompacte a pasta inteira e abra o SerialNumberChecker.exe. Não o
  separe da pasta _internal.
- Não precisa de Python nem de outras instalações.
- Se o Windows mostrar "O Windows protegeu o computador", clique em
  "Mais informações" e "Executar assim mesmo". Isso acontece porque o
  programa não tem assinatura digital.

A PARTIR DO CÓDIGO (opcional):

- Python 3 para Windows (inclui Tkinter) e a biblioteca openpyxl:

      pip install openpyxl

- Depois execute serial_number_checker.py.
- Para criar o executável, use build_exe.bat.

Os sons funcionam no Windows. O ícone do aplicativo é arthur.ico.


18. RECOMENDAÇÕES
-----------------

- Verifique se o arquivo carregado é o correto.
- Confira se as colunas de S/N têm um cabeçalho reconhecido (seção 15).
- Lembre-se de que todas as planilhas dos arquivos XLSX são lidas.
- Se um S/N não aparecer como encontrado, verifique se está escrito
  corretamente e se existe em uma coluna de S/N.
- Revise a aba Sobras antes de exportar.
- Exporte o relatório antes de reiniciar.
- Guarde uma cópia do arquivo original sem modificações.


19. GitHub
----------

Projeto, novidades e downloads:

    https://github.com/Desblackve/serial-number-checker

------------------------------------------------------------
