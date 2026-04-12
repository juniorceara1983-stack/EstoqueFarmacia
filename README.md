# EstoqueFarmácia

Sistema integrado para controle de estoque e geração automática de lista de compras para farmácias.

## Arquitetura

| Camada | Tecnologia | Função |
|---|---|---|
| Banco de dados | Google Sheets | Armazena estoque atual e histórico de vendas |
| Backend / motor | Google Apps Script | Cálculo de média móvel, projeção e API JSON |
| Interface | GitHub Pages (HTML/JS) | Upload CSV, visualização, PDF, WhatsApp |

---

## Como usar

### 1. Configurar o backend (Google Apps Script)

1. Abra a planilha Google Sheets da farmácia.
2. Crie duas abas com os nomes exatos:
   - **`Estoque`** → Coluna A: Nome do Medicamento | Coluna B: Quantidade física atual
   - **`Vendas`** → Coluna A: Nome do Medicamento | Coluna B: Quantidade vendida no período
   - Célula **`D1`** da aba Vendas: número de dias que o histórico de vendas cobre (ex.: `30`)
3. Vá em **Extensões → Apps Script**.
4. Apague o conteúdo padrão e cole o conteúdo do arquivo `appsscript/Code.gs` deste repositório.
5. Salve e clique em **Implantar → Nova implantação**.
   - Tipo: **Aplicativo da Web**
   - Executar como: **Eu**
   - Quem tem acesso: **Qualquer pessoa**
6. Copie a URL de implantação (formato `https://script.google.com/macros/s/.../exec`).

### 2. Configurar o frontend

1. Abra o arquivo `docs/js/config.js`.
2. Substitua o valor de `APPS_SCRIPT_URL` pela URL copiada no passo anterior.
3. Altere `FARMACIA_NOME` para o nome da sua farmácia (aparece no PDF e no WhatsApp).
4. Faça commit e push das alterações.

### 3. Ativar o GitHub Pages

1. Vá em **Settings → Pages** no repositório.
2. Source: **Deploy from a branch** → branch `main` (ou `copilot/…`) → pasta `/docs`.
3. O site ficará disponível em `https://<usuario>.github.io/EstoqueFarmacia/`.

---

## Fluxo de uso diário

```
1. Exportar CSV de Estoque do sistema da farmácia
2. Exportar CSV de Vendas do sistema da farmácia
3. Abrir o app → arrastar os dois CSVs (upload em lote automático)
4. Selecionar período de cobertura (7 / 14 / 21 / 28 dias)
5. Clicar em "Calcular Lista de Compras"
6. Exportar PDF  ou  Enviar pelo WhatsApp
```

## Formato dos CSVs

Ambos os arquivos aceitam `,` ou `;` como separador:

```
Medicamento;Quantidade
Dipirona 500mg;120
Amoxicilina 500mg;45
...
```

A primeira linha com texto não numérico na coluna B é ignorada automaticamente (cabeçalho).

## Fórmula de cálculo

```
Média Diária  = Vendas Totais ÷ Dias do Período de Vendas
Projeção      = Média Diária × Dias de Cobertura Desejados
Necessidade   = Projeção − Estoque Atual   (apenas se > 0)
Comprar       = ⌈ Necessidade ⌉  (arredondamento para cima)
```

Itens cujo estoque já cobre o período solicitado **não aparecem** na lista.

## Estrutura do repositório

```
EstoqueFarmacia/
├── appsscript/
│   ├── Code.gs            ← Backend Google Apps Script
│   └── appsscript.json    ← Manifest do projeto Apps Script
├── docs/                  ← GitHub Pages (frontend)
│   ├── index.html
│   ├── css/style.css
│   └── js/
│       ├── config.js      ← ⚠️ Edite APPS_SCRIPT_URL aqui
│       └── app.js
└── README.md
```

