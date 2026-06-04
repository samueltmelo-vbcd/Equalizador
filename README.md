# Equalização de Estoques — App Web

Aplicação para análise de equalização de estoques consignados por hospital.

## O que o app faz

- Upload do arquivo `.xlsb` com a BASE_ESTUDOS
- Seleção de linhas de produto (Hemodinâmica, Ortopedia, etc.) e regionais
- Configuração da política de cobertura alternativa (120 / 150 / 180 dias)
- Análise de excesso por hospital, com distinção entre excesso real (consumo) e excesso de grade
- Análise de risco de vencimento por lote (usa CMV_2 do próprio arquivo ou base de lotes externa)
- Download do Excel com 6 abas: Resumo, Transferências, Retorno ao CD, Vencimento, Por Hospital, Base Completa

---

## Opção 1 — Streamlit Community Cloud (mais simples, gratuito)

### Passo a passo

1. Crie uma conta em https://streamlit.io/cloud (usa login do Google ou GitHub)

2. Crie um repositório no GitHub com estes 3 arquivos:
   - `app.py`
   - `requirements.txt`
   - `README.md` (este arquivo)

3. No Streamlit Community Cloud, clique em **New app**

4. Conecte o repositório GitHub, selecione o branch `main` e o arquivo `app.py`

5. Clique em **Deploy** — em ~2 minutos a URL estará disponível

6. Compartilhe a URL com a equipe — qualquer pessoa acessa pelo navegador sem instalar nada

**URL gerada:** `https://[seu-usuario]-equalizacao-[hash].streamlit.app`

> Limite gratuito: 1 app público, sem limite de usuários, recursos moderados (suficiente para arquivos de até ~50 MB)

---

## Opção 2 — Google Cloud Run (mais robusto)

### Pré-requisitos

- Conta Google Cloud com faturamento ativado (free tier cobre uso leve)
- [Google Cloud CLI](https://cloud.google.com/sdk/docs/install) instalado

### Deploy em 3 comandos

```bash
# 1. Autenticar
gcloud auth login

# 2. Configurar projeto (crie um no console.cloud.google.com se necessário)
gcloud config set project SEU_PROJECT_ID

# 3. Deploy direto (sem precisar buildar localmente)
gcloud run deploy equalizacao-estoques \
  --source . \
  --region us-central1 \
  --platform managed \
  --allow-unauthenticated \
  --memory 1Gi \
  --timeout 300
```

**URL gerada:** `https://equalizacao-estoques-[hash]-uc.a.run.app`

> Custo estimado: ~US$ 0 para uso interno leve (Cloud Run cobra por requisição, free tier de 2M req/mês)

---

## Estrutura de arquivos esperada no upload

### Base principal (obrigatória)
Arquivo `.xlsb` com aba **BASE_ESTUDOS** contendo:

| Coluna | Descrição |
|--------|-----------|
| HOSPITAL / HOSPITAL AJUSTADO | Nome do hospital |
| SIGLA | Sigla do hospital |
| REGIONAL | Regional (SP, RJ, DF, PB, PE, BA, PA...) |
| LINHA | Linha de produto (HEMODINAMICA, ORTOPEDIA...) |
| SKU / COD VALORES | Código do produto |
| DESCRICAO | Descrição do produto |
| FÍSICO | Quantidade em estoque físico |
| CMM | Consumo médio mensal |
| GRADE | Grade contratual (0 = sem grade) |
| VALOR UNIT | Valor unitário |
| VALOR TOTAL | Valor total do estoque |
| CONSIDERAR NO ESTUDO? | SIM / NÃO |
| STATUS GRADE | ITEM COM GRADE / ITEM SEM GRADE |
| STATUS | ITEM COM CONSUMO / ITEM SEM CONSUMO |

### Base de lotes (opcional)
Arquivo `.xlsx` ou `.csv` com:

| Coluna | Alternativas aceitas |
|--------|---------------------|
| COD_PRODUTO | COD VALORES, Codigo do item |
| HOSPITAL | Filial, SIGLA |
| LOTE | Lote fabricante, Lote interno |
| VALIDADE | Validade, DATA_VALIDADE |
| QUANTIDADE | Quantidade, QTDE, QTD |

---

## Lógica de análise

### Estoque ideal
- **Item com grade:** ideal = valor da grade
- **Item sem grade, com CMM:** ideal = CMM × (dias / 30)
- **Item sem grade, sem CMM:** ideal = 0

### Classificação
| Classificação | Significado |
|--------------|-------------|
| TRANSFERIVEL | Excesso real — CMM justifica; há demanda em outro hospital |
| TRANSFERIVEL - GRADE | Excesso de grade — o hospital tem mais do que a grade exige |
| RETORNO_CD | Sem consumo na rede — retornar ao CD |
| RETORNO_CD - GRADE | Sem consumo nem grade — retornar com prioridade |
| ADEQUADO/DEFICITARIO | Estoque dentro ou abaixo do ideal |
| ADEQUADO/DEFICITARIO - GRADE | Adequado conforme grade contratual |

### Risco de vencimento
- **Meses até vencer** = (data validade − hoje) / 30
- **Meses para consumir** = quantidade do lote / CMM rede
- **Risco real** = meses para consumir > meses até vencer
