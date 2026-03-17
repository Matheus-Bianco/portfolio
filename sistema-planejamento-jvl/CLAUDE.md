# Memoria do Projeto - Sistema PEI SED Joinville

## CLASP
- CLASP esta configurado e funcionando neste projeto.
- Script ID: `18Mf5vgN50qIG_9Nl8XTDnf2AKfDlJChx4Io72xSPuq-ZfqCck5G2t5D0`
- `.clasp.json` ja existe na raiz do projeto.
- `clasp pull` baixa os arquivos do Apps Script (vem como .js, equivalente a .gs).
- `clasp push` envia os arquivos locais para o Apps Script.
- `clasp open` abre o editor no navegador.
- `clasp deploy --description "descricao"` cria novo deploy.

## API REST (ApiService.gs)
- `doPost()` habilitado com 19 endpoints.
- Autenticacao por API Key configurada em `API_CONFIG.API_KEY`.
- Funcao `testarAPI()` disponivel para validar endpoints no editor do Apps Script.
