# Publicar o relatório como site (Streamlit Community Cloud — grátis)

## 1) Repositório PRIVADO no GitHub
1. Crie uma conta no GitHub e um repositório **Private** (ex.: `relatorio-qualidade`).
2. Envie TODOS os arquivos desta pasta (app.py, pipeline_web.py, requirements.txt, scripts/, clientes/, .streamlit/config.toml).
   Pelo site: "Add file → Upload files" e arraste a pasta.

## 2) Publicar
1. Entre em https://share.streamlit.io com a conta do GitHub e clique em **Create app**.
2. Escolha o repositório, branch `main` e o arquivo principal `app.py`.
3. Em **Advanced settings → Secrets**, cole:
   ```
   APP_PASSWORD = "uma-senha-forte"
   GEMINI_API_KEY = "sua-chave-do-gemini"   # análise de IA (opcional; sem ela o PDF sai sem as páginas de IA)
   TELEGRAM_TOKEN = "TOKEN_DO_BOT"
   ```
4. Deploy. Você recebe um link; passe o link + a senha aos usuários.
   (Em Settings → Sharing do app, dá para liberar só e-mails específicos.)

## 3) Cadastrar clientes
Cada cliente é uma pasta em `clientes/` (copie `clientes/exemplo`):
- `cadastro.xlsx` — abas Talhoes (cd_fazenda, cd_zona, cd_talhao, nome_talhao), Equipamentos (cd_equipamento, frota),
  Processos (cd_operacao, descricao_processo). Qualquer .xlsx da pasta é lido.
- `*.geojson` — arquivo de talhões (propriedades FAZENDA, ZONA, TALHAO, DESC_TALHA + contorno). A chave do talhão é fazenda+zona+talhão e DESC_TALHA vira o nome no relatório. (`*.kml` também funciona.) Os arquivos `.mbtiles` não são necessários.
- `telegram.env` — só `TELEGRAM_CHAT_ID=<id do grupo do cliente>`.
Depois de alterar, faça commit no GitHub; o site se atualiza sozinho.

## Notas
- O token do bot NUNCA vai no repositório, só nos Secrets.
- O bot precisa ser administrador do grupo (permissão de enviar arquivos).
- Plano gratuito: memória limitada (~1 GB). CSV de ~470 mil linhas usou ~870 MB; até ~150 mil linhas usa ~300 MB.
- O app "dorme" sem uso; a primeira abertura pode levar alguns segundos.
