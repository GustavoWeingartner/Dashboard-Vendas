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

O app busca o CSV publico do Sheets ao abrir e atualiza automaticamente a cada 5 minutos. O botao `Atualizar` faz uma nova leitura imediata sem recarregar a pagina.

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

O upload manual foi removido da interface; os arquivos locais permanecem apenas como fallback caso o Google Sheets fique indisponivel.

## Deploy na Vercel

Use estas configuracoes:

```text
Build Command: npm run build
Output Directory: public
```

O comando `npm run build` gera dentro de `public` tudo que a Vercel precisa publicar:

- `index.html`
- `styles.css`
- `app.compiled.js`
- `vendor/`
- `assets/`
- `base-dashboard-data.js`
- fallbacks locais `base-dashboard.csv` e `base-dashboard.xlsx`
