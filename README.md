# Dashboard de Comparação de Nomes entre Planilhas

Sistema Flask para comparação de nomes entre duas planilhas usando algoritmos de fuzzy matching otimizados para português brasileiro.

## 📋 Características

- **Upload de Planilhas**: Suporte para Excel (.xlsx, .xls) e CSV
- **Fuzzy Matching**: Algoritmos avançados de comparação de strings
- **Interface Moderna**: Dashboard responsivo com Bootstrap 5
- **Otimizado para PT-BR**: Tratamento de acentos e caracteres especiais
- **Exportação**: Resultados em formato Excel
- **Configurável**: Ajuste de limiar de similaridade e algoritmos

## 🚀 Instalação

1. **Clone ou baixe o projeto**
```bash
cd COMPARAR
```

2. **Instale as dependências**
```bash
pip install -r requirements.txt
```

3. **Execute a aplicação**
```bash
python app.py
```

4. **Acesse o dashboard**
```
http://localhost:5000
```

## 📊 Como Usar

### 1. Upload das Planilhas
- Selecione duas planilhas (Excel ou CSV)
- Clique em "Fazer Upload"

### 2. Configuração
- Escolha as colunas para comparação em cada planilha
- Ajuste o limiar de similaridade (50-100%)
- Selecione o algoritmo de comparação

### 3. Comparação
- Clique em "Iniciar Comparação"
- Aguarde o processamento

### 4. Resultados
- Visualize estatísticas e correspondências
- Exporte os resultados em Excel

## 🔧 Algoritmos Disponíveis

### Token Sort Ratio (Recomendado)
- Melhor para nomes com palavras em ordens diferentes
- Exemplo: "João Silva" vs "Silva, João"

### Ratio Simples
- Comparação direta caractere por caractere
- Mais rápido, menos flexível

### Partial Ratio
- Encontra a melhor correspondência parcial
- Útil para nomes com prefixos/sufixos

### Token Set Ratio
- Ignora palavras duplicadas
- Bom para nomes com títulos repetidos

## 📁 Estrutura do Projeto

```
COMPARAR/
├── app.py                 # Aplicação Flask principal
├── requirements.txt       # Dependências Python
├── README.md             # Documentação
├── templates/
│   └── index.html        # Interface do dashboard
├── static/
│   ├── css/
│   │   └── style.css     # Estilos personalizados
│   └── js/
│       └── app.js        # JavaScript do frontend
└── uploads/              # Pasta para arquivos temporários
```

## 🛠️ Tecnologias Utilizadas

- **Backend**: Flask, Pandas, RapidFuzz
- **Frontend**: Bootstrap 5, JavaScript ES6
- **Processamento**: OpenPyXL, XlRd
- **Algoritmos**: Levenshtein Distance, Token-based matching

## ⚙️ Configurações Avançadas

### Limiar de Similaridade
- **50-69%**: Correspondências muito flexíveis
- **70-84%**: Correspondências moderadas
- **85-100%**: Correspondências rigorosas

### Tratamento de Nomes PT-BR
O sistema automaticamente:
- Remove acentos (á, é, í, ó, ú, ç)
- Normaliza espaços em branco
- Converte para minúsculas
- Trata caracteres especiais

## 📈 Casos de Uso

### Deduplicação de Dados
- Identificar registros duplicados
- Limpeza de bases de dados
- Consolidação de informações

### Matching de Clientes
- Relacionar bases diferentes
- Identificar clientes em comum
- Análise de sobreposição

### Validação de Dados
- Verificar consistência
- Identificar erros de digitação
- Padronização de nomes

## 🔍 Exemplos de Correspondências

| Nome Original | Correspondência | Score |
|---------------|-----------------|-------|
| João da Silva | Joao Silva | 95% |
| Maria Santos | M. Santos | 87% |
| José Oliveira | Jose de Oliveira | 92% |

## 📋 Requisitos do Sistema

- Python 3.7+
- 4GB RAM (recomendado)
- Navegador moderno
- Planilhas com até 100.000 linhas

## 🐛 Solução de Problemas

### Erro de Upload
- Verifique o formato do arquivo
- Tamanho máximo: 16MB
- Formatos suportados: .xlsx, .xls, .csv

### Comparação Lenta
- Reduza o número de linhas
- Aumente o limiar de similaridade
- Use algoritmo "Ratio Simples"

### Resultados Inesperados
- Ajuste o limiar de similaridade
- Teste diferentes algoritmos
- Verifique a qualidade dos dados

## 📞 Suporte

Para dúvidas ou problemas:
1. Verifique a documentação
2. Consulte os logs da aplicação
3. Teste com dados menores

## 🔄 Atualizações Futuras

- [ ] Suporte a mais formatos de arquivo
- [ ] API REST para integração
- [ ] Processamento em lote
- [ ] Machine Learning para melhor matching
- [ ] Interface multilíngue

---

**Desenvolvido com ❤️ para facilitar a comparação de dados em português brasileiro**