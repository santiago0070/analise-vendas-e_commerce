
# 📊 Projeto de Carga de Dados - E-commerce de Cursos Online

Este projeto tem como objetivo **importar, validar, limpar e inserir dados de vendas de cursos online** a partir de uma planilha Excel para uma base de dados **SQL Server**.

---

## 🚀 Funcionalidades

- 📥 Leitura de dados de uma planilha Excel (`.xlsx`)
- 🧹 Validação e limpeza da coluna de datas
- 🧮 Conversão de dados para os tipos corretos (inclusive monetários)
- 🗃️ Inserção dos dados na tabela `Vendas` do banco de dados SQL Server

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.x**
- **Pandas** (manipulação de dados)
- **PyODBC** (integração com SQL Server)
- **SQL Server** (armazenamento de dados)
- **Excel** (fonte dos dados)

---

## 📁 Estrutura dos Dados

**Planilha Excel (`Base Tratada`) deve conter as colunas:**

- `Data` (formato Excel serial date ou datetime)
- `Ano` (int)
- `Mes` (string ou int)
- `Vendedor` (string)
- `Cliente` (string)
- `Regiao` (string)
- `Produto` (string)
- `Valor` (float ou string com valores monetários)

---

## ⚙️ Como Executar

1. Instale as bibliotecas necessárias:
   ```bash
   pip install pandas pyodbc openpyxl
   ```

2. Configure sua string de conexão com o SQL Server:
   ```python
   conn_str = (
       r'DRIVER={{SQL Server}};'
       r'SERVER=SEU_SERVIDOR;'
       r'DATABASE=VendasCursosOnline;'
       r'Trusted_Connection=yes;'
   )
   ```

3. Altere o caminho do arquivo Excel:
   ```python
   df = pd.read_excel(r'C:\seu_caminho\arquivo.xlsx', sheet_name='Base Tratada')
   ```

4. Execute o script Python:
   ```bash
   python seu_script.py
   ```

---

## 🧪 Validação de Datas

O script valida se a coluna `Data`:

- Está no formato datetime
- Ou pode ser convertida a partir de valores numéricos (serial do Excel)
- Não contém valores nulos ou fora do intervalo aceitável

---

## 🗃️ Tabela no SQL Server

A tabela `Vendas` precisa existir no banco `VendasCursosOnline` com a seguinte estrutura:

```sql
CREATE TABLE Vendas (
    Data DATE,
    Ano INT,
    Mes INT,
    Vendedor VARCHAR(100),
    Cliente VARCHAR(100),
    Regiao VARCHAR(100),
    Produto VARCHAR(100),
    Valor FLOAT
);
```

---

## 📝 Observações

- A coluna `Valor` é convertida com `errors='coerce'`, o que substitui valores inválidos por `NaN`.
- Registros duplicados são removidos antes da inserção.
- Certifique-se de que os dados da planilha estejam limpos para evitar erros na carga.

---

## 👨‍💻 Autor

**Rodrigo Santiago**  
Analista de Dados | Python • SQL • Power BI  
📧 rodrigosantiago@email.com

