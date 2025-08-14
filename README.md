# 📊 Sistema de Relatórios de Vendas

Um sistema completo de relatórios de vendas em Python com interface web interativa. Permite gerar:

- Arquivos Excel individuais por vendedor  
- Planilha resumo com gráficos  
- Relatório PDF profissional  
- Download em ZIP de todos os arquivos  

---

## 🔹 Funcionalidades

- Upload de Excel com colunas obrigatórias: **Vendedor, Produto, Vendas**  
- Arquivos separados por vendedor para fácil distribuição  
- **Planilha resumo**:  
  - Aba Resumo → total de vendas por vendedor  
  - Aba Gráficos → gráfico de barras e pizza  
  - Abas individuais para cada vendedor  
- Relatório PDF completo com tabelas e gráficos  
- Visualização interativa de gráficos na interface **Streamlit**  
- ZIP para download contendo todos os arquivos  

---

## 🔹 Tecnologias

- **Python 3.12+**  
- **Streamlit** – Interface web interativa  
- **Pandas** – Manipulação e agregação de dados  
- **Matplotlib** – Criação de gráficos  
- **OpenPyXL** – Criação e edição de planilhas Excel  
- **ReportLab** – Geração de PDFs com gráficos e tabelas  
- **ZipFile / io / tempfile** – Gerenciamento de arquivos temporários e compactação  

---

## 🔹 Estrutura da Planilha

A planilha precisa conter exatamente estas colunas:

| Vendedor | Produto | Vendas |
|----------|--------|--------|
| João     | Caneta | 150.50 |
| Maria    | Caderno| 320.00 |

> Você pode baixar o modelo direto da aplicação para preencher.

## 🔹 Como Executar

1. **Clone o repositório:**

```bash
git clone https://github.com/seu-usuario/relatorios-vendas.git
cd relatorio-vendas
```

2. **Instale as dependências:**
```bash
pip install streamlit pandas matplotlib openpyxl reportlab
```

3. **Execute a aplicação:**

```bash
streamlit run app.py
```

4. **Na interface:**
- Faça upload da planilha Excel.
- Visualize gráficos e resumo de vendas.
