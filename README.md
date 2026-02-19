# 🎲 Simulação de Monte Carlo — Análise Preventiva de Risco de Ruptura

**Plano de Sortimentos | Supply Chain & Ciência de Dados**

---

## 📋 Índice

1. [O que este projeto faz](#o-que-este-projeto-faz)
2. [Por que usar Monte Carlo](#por-que-usar-monte-carlo)
3. [Estrutura de estoque](#estrutura-de-estoque)
4. [Instalação](#instalação)
5. [Como usar](#como-usar)
6. [Entendendo as colunas do output](#entendendo-as-colunas-do-output)
7. [Interpretando os resultados](#interpretando-os-resultados)
8. [Parâmetros ajustáveis](#parâmetros-ajustáveis)
9. [Limitações do modelo](#limitações-do-modelo)
10. [FAQ](#faq)

---

## 🎯 O que este projeto faz

Este script Python simula **10.000 cenários futuros** para cada SKU do seu Plano de Sortimentos, calculando:

✅ **Probabilidade de ruptura** nos próximos 30 dias  
✅ **Dias de cobertura** antes de zerar o estoque  
✅ **Estoque de Segurança dinâmico** (recomendação baseada em risco real, não fórmula estática)  
✅ **Risco consolidado** (ALTO / MÉDIO / BAIXO)

---

## 🤔 Por que usar Monte Carlo?

### Método tradicional (determinístico):
```
Cobertura = Estoque Total ÷ Venda Média
```
**Problema:** Ignora completamente a **variabilidade** da demanda, atrasos no lead time e problemas no selamento.

### Método Monte Carlo (probabilístico):
```
Simula 10.000 meses diferentes considerando:
  - Demanda variável (Normal/Poisson + picos aleatórios)
  - Lead time incerto (Log-Normal)
  - Selamento com variabilidade operacional
  
Resultado: Distribuição de probabilidades
```
**Vantagem:** Você sabe a **probabilidade real** de ruptura, não apenas um número médio.

---

## 📦 Estrutura de Estoque

O modelo respeita a estrutura **multi-nós** do seu supply chain:

```
ESTOQUE TOTAL = Extrema + Itapeva[Total] + Linhares[Bruto]
```

Onde:
- **Extrema:** Estoque já selado (NFC), disponível para venda imediata
- **Itapeva[Total]:** Itapeva[Bruto] + Itapeva[NFC] (mix de bruto e selado)
- **Linhares[Bruto]:** Sem selo, precisa passar por industrialização

### Processo de industrialização (NFC):
O estoque **bruto** precisa passar pelo processo de **selamento** antes de estar disponível para venda no site. A taxa de selamento varia por dia e por curva ABC:

| Curva | Taxa de selamento/dia |
|---|---|
| AA | 150% da venda média |
| A  | 130% da venda média |
| B  | 100% da venda média |
| C  | 80% da venda média |

---

## 🔧 Instalação

### Pré-requisitos:
- Python 3.8 ou superior
- Arquivo `Plano_de_Sortimentos_4_0.xlsx` na mesma pasta do script

### Instalar dependências:
```bash
pip install pandas numpy openpyxl
```

---

## 🚀 Como usar

### 1. Colocar arquivos na mesma pasta:
```
📁 Minha_Pasta/
  ├── monte_carlo_supply_FINAL.py
  ├── Plano_de_Sortimentos_4_0.xlsx
  └── README.md
```

### 2. Executar no terminal:
```bash
python monte_carlo_supply_FINAL.py
```

### 3. Aguardar ~2 minutos:
```
[1/4] Lendo Plano de Sortimentos...
   ✓ 199 SKUs encontrados

[2/4] Preparando estrutura de estoque...
   ✓ Com estoque disponível: 47 SKUs

[3/4] Rodando 10,000 simulações por SKU...
   Progresso: 47/47 (100%)
   ✓ Simulações concluídas!

[4/4] Criando Excel formatado...
   ✓ Excel salvo: Analise_Preventiva_Monte_Carlo.xlsx
```

### 4. Abrir o arquivo gerado:
`Analise_Preventiva_Monte_Carlo.xlsx`

---

## 📊 Entendendo as Colunas do Output

| Coluna | O que significa | Exemplo |
|--------|----------------|---------|
| **Item** | Nome do SKU | Mini Tote Preta |
| **Curva ABC** | Classificação de importância | AA (mais crítico) → C (menos crítico) |
| **NFC_Dispon** | Estoque disponível para venda HOJE | 9.081 unidades |
| **Bruto** | Estoque sem selo (precisa industrialização) | 15.350 unidades |
| **Total** | NFC + Bruto | 24.431 unidades |
| **Venda/dia** | Média de vendas (últimos 14 dias) | 381 unidades/dia |
| **Dias_Cobertura** | Dias até ruptura (mediana dos cenários) | 30 dias |
| **Prob_Ruptura_%** | Probabilidade de romper nos próximos 30 dias | 0% (não vai romper) |
| **Prob_NFC_Zero_%** | Probabilidade do NFC zerar (mesmo sem ruptura total) | 5% |
| **ES_P95** | Estoque de Segurança recomendado (nível 95%) | 1.547 unidades |
| **ES_Atual** | Estoque de Segurança configurado no plano | 200 unidades |
| **Gap_ES** | Diferença (ES_Atual - ES_P95) | **-1.347** ⚠️ INSUFICIENTE |
| **Risco** | Classificação final | ALTO / MÉDIO / BAIXO |

---

## 📈 Interpretando os Resultados

### 🔴 Risco ALTO (Prob. Ruptura > 30%)
**O que fazer:**
1. Priorizar Follow Up desse SKU
2. Verificar se há estoque em outro nó para transferência
3. Acelerar processo de selamento (se houver bruto)
4. Considerar pausar promoções até reposição

**Exemplo:**
```
Lancheira BTS Preta | Dias: 2 | NFC: 0 | Bruto: 10 | Prob: 100%
```
→ Vai romper em 2 dias com certeza absoluta. **AÇÃO IMEDIATA.**

---

### 🟡 Risco MÉDIO (Prob. Ruptura 10-30%)
**O que fazer:**
1. Monitorar **semanalmente**
2. Preparar plano de contingência
3. Não reduzir estoque de segurança

**Exemplo:**
```
Produto X | Dias: 25 | Prob: 15%
```
→ 15% de chance de romper. Não é crítico mas merece atenção.

---

### 🟢 Risco BAIXO (Prob. Ruptura < 10%)
**O que fazer:**
1. Situação confortável
2. Avaliar se não está sobre-estocado (capital de giro parado)

---

### ⚪ Estoque PARADO
**O que fazer:**
1. Verificar se é lançamento futuro ou descontinuado
2. Se for descontinuado, considerar liquidação
3. Se for problema de exposição no site, corrigir

---

## 🔩 Parâmetros Ajustáveis

### No topo do script:
```python
N_SIM = 10_000      # Número de simulações (mais = mais preciso, mais lento)
HORIZON = 30        # Horizonte de análise em dias
np.random.seed(42)  # Seed aleatória (para reproducibilidade)
```

### Coeficiente de variação da demanda:
```python
sigma_demanda = venda_media * 0.35  # 35% de variabilidade
```
Ajuste se seus SKUs têm variação maior ou menor que 35%.

### Probabilidade de pico de vendas:
```python
if np.random.random() < 0.05:  # 5% de chance
    demanda *= 2.0              # Dobra a demanda
```
Ajuste se picos são mais ou menos frequentes.

---

## ⚠️ Limitações do Modelo

O que o modelo **NÃO** captura:

❌ **Canibalização entre SKUs** — Se a Mini Tote Preta rompe, clientes compram a Cinza  
❌ **Sazonalidade explícita** — Usa média flat para os 30 dias  
❌ **Correlação de demanda** — Promoção de um produto afeta vendas de outros  
❌ **Capacidade de selamento compartilhada** — Trata cada SKU isoladamente  
❌ **Decisão de descontinuação** — Não detecta se produto está em fase de saída  

### Como mitigar:
- Rodar a simulação **semanalmente** para capturar mudanças de tendência
- Ajustar manualmente SKUs em promoção (aumentar venda média)
- Combinar com análise qualitativa da operação

---

## ❓ FAQ

### **P: Por que a Mini Tote Preta mostra 30 dias de cobertura se tem 24.431 unidades e vende 381/dia?**
**R:** Porque 24.431 ÷ 381 = **64 dias** teoricamente. Mas o modelo limita a exibição a 30 dias (horizonte da análise). Se quiser ver coberturas maiores, aumente `HORIZON = 60`.

---

### **P: O que significa "Gap_ES" negativo?**
**R:** Seu Estoque de Segurança atual é **menor** do que o recomendado pela simulação.

**Exemplo:**
```
ES_Atual: 200
ES_P95:   1.547
Gap:      -1.347 ⚠️
```
→ Você está **1.347 unidades abaixo** do buffer necessário para ter 95% de certeza de não romper.

---

### **P: Posso rodar no Google Colab?**
**R:** Sim! Use a versão `Monte_Carlo_Google_Colab_199_SKUs.py` (arquivo separado).

---

### **P: Como interpretar "Prob_Ruptura = 56.1%"?**
**R:** De 10.000 cenários simulados, em **5.610 cenários** (56.1%) houve pelo menos 1 dia de ruptura nos próximos 30 dias.

É quase uma **moeda cara ou coroa** — há chance real de romper mas também de não romper.

---

### **P: Por que SKUs com muito estoque têm risco ALTO?**
**R:** Porque o **bruto é insuficiente**. 

**Exemplo:**
```
Mini Tote Off White
NFC:   3.486 (ok para ~22 dias)
Bruto: 60    (só cobre 0,4 dias de selamento)
```
→ Mesmo com 3.546 unidades totais, vai romper porque o bruto acaba rápido e não há reposição.

---

### **P: Como explicar Monte Carlo para o time sem conhecimento técnico?**
**R:** Use essa analogia:

> *"Imagine que você joga 10.000 versões do mês de março. Em algumas versões o navio atrasa, em outras a Mini Tote viraliza no Instagram, em outras o selamento trava. No final, a gente conta: em quantos desses 10.000 meses a gente rompeu?"*

---

## 📞 Suporte

Para dúvidas ou ajustes no modelo, consulte:
- Documentação NumPy: https://numpy.org/doc/
- Documentação Pandas: https://pandas.pydata.org/docs/
- Openpyxl (Excel): https://openpyxl.readthedocs.io/

---

## 📄 Licença

Este projeto é de uso interno para análise de supply chain.

---

## 🎓 Referências

- **Simulação de Monte Carlo:** Metropolis, N.; Ulam, S. (1949). "The Monte Carlo Method"
- **Gestão de Estoques:** Silver, E. A.; Pyke, D. F.; Peterson, R. (1998). "Inventory Management and Production Planning and Scheduling"
- **Distribuições Probabilísticas:** Gentle, J. E. (2003). "Random Number Generation and Monte Carlo Methods"

---

**Última atualização:** 2025-02-19  
**Versão:** 1.0 FINAL
