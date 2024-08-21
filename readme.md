# Painel de Gestão de Chamados EPM
Correção das datas ao realizar a importação dos chamados para o sharepoint.

#### Problema:
❌ Ao gerar a lista de chamados, as datas contidas no arquivo Excel possui datas em finais de semana e o sharepoint acusa erro nisso.

#### Solução:
✅ Este script analisa as datas do arquivo original, verifica se possui algum no final de semana e se possuir, ele joga para a sexta-feira anterior.
