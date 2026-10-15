# Painel de Receitas JB

Painel online com as receitas da aba **Lançamento de Receitas** da planilha do cliente. Ele se atualiza sozinho **todo dia às 6h (horário de Brasília)**, sem depender de nenhum computador ligado.

- **Link do painel:** https://solverjobs-design.github.io/painel-receitas-jb/
- **Repositório:** https://github.com/solverjobs-design/painel-receitas-jb
- **Acesso:** o link pede uma senha. Solver e cliente usam a mesma. Marcando "Lembrar neste aparelho", a senha não é pedida de novo naquele navegador. O botão **Sair** apaga a senha guardada.

## O que o cliente precisa fazer

Só continuar preenchendo a planilha como sempre. O painel lê a planilha todo dia às 6h e mostra **"Dados atualizados em DD/MM/AAAA HH:MM"** no topo.

## Como funciona

1. Todo dia às 6h, o **GitHub Actions** roda o `gerar.py`.
2. O script entra na planilha com uma **conta de serviço do Google**, que tem acesso só de leitura, e lê a aba configurada em `config.json`.
3. Ele confere os dados: se a aba sumiu, se falta alguma coluna ou se há poucas linhas. **Se algo estiver errado, ele para e o painel do dia anterior continua no ar.** O GitHub manda um e-mail de falha para o dono do repositório.
4. Com os dados ok, ele calcula os indicadores, monta o painel a partir do `modelo.html`, **criptografa o conteúdo com a senha** (AES-256) e publica no **GitHub Pages**.

O repositório guarda só o código. Os dados da planilha não ficam no repositório nem nos logs, e o painel publicado só é legível com a senha.

## Tarefas do dia a dia

| Quero... | Como fazer |
|---|---|
| Atualizar agora | GitHub → aba **Actions** → **Atualizar painel** → **Run workflow** |
| Ver se rodou | GitHub → aba **Actions** (verde = ok, vermelho = falhou) |
| Pausar a atualização | Actions → Atualizar painel → menu **⋯** → **Disable workflow** |
| Retomar | Actions → Atualizar painel → **Enable workflow** |
| Trocar a senha | Settings → Secrets and variables → Actions → `PAINEL_SENHA` → **Update**. Depois rode **Run workflow**. Quem marcou "lembrar" terá que digitar a senha nova. |
| Mudar o horário | Em `.github/workflows/atualizar.yml`, altere `cron: "0 9 * * *"`. O horário é UTC: Brasília + 3 h (6h → 9, 7h → 10). |
| Ler outra aba | Em `config.json`, troque `"aba"`. Se as colunas mudarem de nome, ajuste `"colunas"`. |
| Agrupar um tratamento novo | Em `config.json`, acrescente o início do nome (sem acento, minúsculo) no grupo certo de `"grupos"`. Tratamentos fora dos grupos aparecem com o próprio nome. |

## Configuração (feita uma vez)

**Secrets do repositório** (Settings → Secrets and variables → Actions):

- `GOOGLE_SA_JSON`: conteúdo inteiro do arquivo `.json` da chave da conta de serviço.
- `PAINEL_SENHA`: senha do painel (mínimo de 8 caracteres; de preferência 4 palavras aleatórias).

**Planilha:** compartilhada como **Leitor** com o e-mail da conta de serviço (termina em `.iam.gserviceaccount.com`). A planilha deve continuar **restrita** (não "qualquer pessoa com o link").

**GitHub Pages:** Settings → Pages → Source: **GitHub Actions**.

## Mensagens de erro comuns (aba Actions)

- `sem acesso à planilha` → a planilha não está compartilhada com o e-mail da conta de serviço.
- `aba "..." não encontrada` → a aba foi renomeada; ajuste `config.json`.
- `a coluna "..." sumiu da aba` → um cabeçalho da linha 1 foi renomeado ou apagado.
- `PAINEL_SENHA não está cadastrado` / `GOOGLE_SA_JSON não está cadastrado` → falta o secret.

## Teste local (opcional)

```bash
pip install -r requirements.txt
PAINEL_SENHA=uma-senha-de-teste python gerar.py --xlsx copia-da-planilha.xlsx
```

O resultado fica em `site/index.html`. Abra-o por um servidor local (por exemplo, `python -m http.server --directory site`), porque a descriptografia exige `https` ou `localhost`.

## Arquivos

| Arquivo | Para quê |
|---|---|
| `config.json` | planilha, aba, colunas e agrupamento de tratamentos |
| `gerar.py` | lê a planilha, valida, calcula e gera o painel criptografado |
| `modelo.html` | layout e gráficos do painel |
| `.github/workflows/atualizar.yml` | agendamento diário e publicação |
