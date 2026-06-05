# Marketplace Sales Dashboard

Dashboard React/TypeScript para analise de vendas de e-commerce e marketplaces.

## Fonte de dados principal

O dashboard agora carrega a base diretamente de uma planilha publica do Google Sheets:

```text
https://docs.google.com/spreadsheets/d/1JldFrcw8oaVAWXXhFyJCMVvm_Be9IXju90UAqpzSLXM/edit?gid=27856229#gid=27856229
```

Configuracao usada no codigo:

```text
Spreadsheet ID: 1JldFrcw8oaVAWXXhFyJCMVvm_Be9IXju90UAqpzSLXM
GID da aba: 27856229
```

O app busca o CSV publico do Sheets ao abrir e atualiza automaticamente a cada 5 minutos. O botao `Atualizar Sheets` faz uma nova leitura imediata sem recarregar a pagina.

## Passo a passo

1. No Google Sheets, clique em `Compartilhar`.
2. Em `Acesso geral`, deixe como `Qualquer pessoa com o link`.
3. Mantenha permissao de `Leitor`.
4. Execute o dashboard localmente:

```powershell
node scripts/server.mjs
```

5. Abra:

```text
http://localhost:4173
```

Ao usar o servidor local, a rota `/api/google-sheets` funciona como proxy para evitar problemas de CORS. Se a planilha do Google ficar indisponivel, o app usa `public/base-dashboard.csv/xlsx` como fallback local.

## Trocar a planilha futuramente

Atualize os valores abaixo em:

- `src/app.tsx`
- `scripts/server.mjs`

```text
GOOGLE_SHEETS_SPREADSHEET_ID
GOOGLE_SHEETS_GID
```

Depois recompile:

```powershell
node scripts/build.mjs
```

O upload manual de Excel ou CSV continua disponivel no botao `Carregar planilha`.
