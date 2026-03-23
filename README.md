# Relaxa Vagabundo(a) esse Assistente de Relatório de Estágio Vai te salvar ☺️

Gerador de relatório acadêmico com IA (Google Gemini) para o curso de Licenciatura em Ciências: Biologia e Química.

## Como instalar

```bash
pip install -r requirements.txt
```

## Como configurar a chave de API

1. Acesse https://aistudio.google.com/apikey e crie uma chave
2. Abra o arquivo `assistente.py`
3. Na linha 12, substitua `SUA CHAVE AQUI` pela sua chave:
```python
minha_chave = os.environ.get("GEMINI_API_KEY", "cole_sua_chave_aqui")
```

## Como executar

```bash
python assistente.py
```

## ⚠️ Atenção

- Nunca compartilhe sua chave de API publicamente
- O arquivo `projeto_estagio_vangles.json` é criado automaticamente e salva seu progresso
- Os arquivos `.json` e `.docx` são ignorados pelo Git (não sobem para o GitHub)
