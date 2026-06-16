# gererp

Gerador de Planilha de MTR para importar no ERP.

## Nova base local SQLite

A aplicação agora usa um banco local `gererp.db` para manter os cadastros que antes ficavam apenas em planilhas ou no código:

- CNPJs e código **Gerador CNPJCPF**;
- resíduos;
- placas e motoristas;
- vínculos de prefixo do documento com resíduos;
- estrutura inicial para histórico de processamentos e avisos.

Na primeira execução, se o banco ainda não tiver dados, o sistema importa automaticamente as planilhas auxiliares existentes no diretório do script:

- `cnpj_cetesb.xlsx`;
- `Listagem_Residuos_Geral_02_04_26.xlsx`.

Os arquivos de entrada recebidos de usuários e os arquivos de saída gerados continuam sendo planilhas `.xlsx` transitórias do processamento.

## Manutenção dos cadastros

A seção **Cadastros do banco** abre uma janela de manutenção com abas para:

- inserir, alterar e apagar/desativar CNPJs e o código Gerador CNPJCPF;
- exportar a base ativa de CNPJs com seus respectivos códigos CETESB/Gerador CNPJCPF para `.xlsx`;
- inserir, alterar e apagar/desativar placas e motoristas;
- pesquisar, inserir, corrigir e apagar/desativar resíduos individualmente, incluindo IBAMA, classe, estado físico, acondicionamento, tratamento, dados ONU/risco/embarque e campos auxiliares.

As alterações feitas nessa janela são salvas no `gererp.db` e passam a ser usadas na próxima geração da planilha de saída.

### Importar ou reimportar as bases auxiliares

Se algum cadastro não aparecer no banco depois da primeira execução, use os botões de importação na seção **Cadastros do banco**. Cada botão pede para você escolher o local exato da planilha antes de importar:

- **Importar CNPJs…**: selecione a planilha de CNPJs e código Gerador CNPJCPF, normalmente `cnpj_cetesb.xlsx`;
- **Importar resíduos…**: selecione a planilha de resíduos, normalmente `Listagem_Residuos_Geral_02_04_26.xlsx`.

Os botões podem ser usados novamente quando qualquer uma das planilhas auxiliares for atualizada. Durante a importação, o app mostra uma barra de avanço com percentual, linhas lidas e registros importados para indicar que a operação está em andamento.

## Validações e busca na geração

Antes de gerar a planilha de saída, os campos **Código-Destinador** e **Código-Transportador** são obrigatórios. Na seção **Grupos prefixo → resíduos**, o campo de seleção de resíduos permite digitar números ou letras para filtrar por código interno, código IBAMA ou descrição antes de adicionar o resíduo ao prefixo.

## CNPJs sem Código CETESB

Ao gerar a planilha de saída, registros cujo CNPJ não possua Código CETESB cadastrado são ignorados. Se houver pendências, o sistema cria automaticamente uma planilha separada ao lado da saída, com o sufixo `_CNPJs_sem_codigo.xlsx`, contendo CNPJ, quantidade de ocorrências, documentos, placas e motivo da exclusão.
