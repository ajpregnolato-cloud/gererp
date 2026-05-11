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
