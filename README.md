# de-para-dre

Aplicativo Streamlit para conciliação entre contábil e cliente/DRE, com parametrização por empresa via Supabase.

## Dependências

O projeto usa:
- `streamlit`
- `pandas`
- `supabase`
- `openpyxl`

## Execução local

```bash
streamlit run app.py
```

## Deploy no Streamlit Cloud

O Streamlit Cloud precisa do arquivo `requirements.txt` na raiz do repositório para instalar as dependências.

Também é necessário configurar os secrets do app com:
- `SUPABASE_URL`
- `SUPABASE_KEY`
- `SUPABASE_SERVICE_ROLE_KEY` opcional, se você quiser usar a service key no backend
