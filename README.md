
# get_fve_token

Biblioteca interna para **coletar tokens enviados por e-mail** usando **Microsoft Graph (Application)**.

> ⚠️ Requer permissões **Mail.Read** (ler) e **Mail.ReadWrite** (marcar como lido) no modo **Application** + **Admin consent**.

## Instalação (via git tag)

```bash
pip install "git+https://github.com/JoaoRBA/get_fve_token.git@v0.1.0#egg=get_fve_token"
```

## Variáveis de ambiente
- GRAPH_TENANT_ID
- GRAPH_CLIENT_ID
- GRAPH_CLIENT_SECRET

## Uso
```python
from get_fve_token import GraphMailClient, fetch_token_simple

client = GraphMailClient()
token = client.fetch_token(
    mailbox="lm.barros@ramosbenedetti.adv.br",
    subject_keyword="Santander, sua segurança em primeiro lugar",
    timeout_seconds=300,
    mark_read=True,
)
print("TOKEN:", token)
```
