# 📦 Pasta `libs_local/` · Solução para erros de bibliotecas

## 🐛 Problema resolvido
1. Computadores corporativos bloqueiam CDN externo → libs não carregavam
2. SheetJS (Excel) tem caracteres internos que QUEBRAM quando embutido inline

## ✅ Solução final
- **Chart.js, PptxGenJS, ChartDataLabels**: embutidas inline no HTML (funcionam bem)
- **SheetJS (xlsx)**: servido como ARQUIVO EXTERNO ao lado do HTML
  (porque inline quebra o parser do browser)

## 📁 Como configurar (UMA VEZ)

### 1. Pasta `libs_local/` ao lado do script Python
```
INSIGHTS/
├── sales_dashboard_v10.py
├── dashboard_template_v10.html
├── libs_local/                      ← criar esta pasta
│   ├── xlsx.mini.min.js             (SheetJS - versão sem codepages)
│   ├── chart.umd.min.js
│   ├── pptxgen.bundle.js
│   └── chartjs-plugin-datalabels.min.js
```

### 2. Os arquivos .js estão no libs_local.zip
Descompacte e coloque os 4 arquivos na pasta.

### 3. Rode o Python
```bash
python sales_dashboard_v10.py
```

## ⚠️ IMPORTANTE: o arquivo xlsx.mini.min.js fica AO LADO do HTML
Quando o Python roda, ele:
1. Embute Chart.js, PptxGenJS e DataLabels DENTRO do HTML
2. **Copia o xlsx.mini.min.js para a mesma pasta do HTML gerado**
3. O HTML referencia: `<script src="xlsx.mini.min.js"></script>`

**Por isso**: ao mover/enviar o HTML, leve junto o arquivo `xlsx.mini.min.js`!
Se mandar só o HTML por email, o Excel não vai funcionar (mas PPT e gráficos sim).

### 💡 Dica para enviar por email
Zipe o HTML + xlsx.mini.min.js juntos:
```
dashboard.zip
├── dashboard_sales_insightsv10.html
└── xlsx.mini.min.js
```
A pessoa descompacta os 2 juntos e abre o HTML.

## ✅ Verificação
Ao abrir o HTML, pressione F12 → Console. Deve aparecer:
```
✅ Todas as bibliotecas carregadas: SheetJS, Chart.js, PptxGenJS
```
Se aparecer banner vermelho no topo → falta o xlsx.mini.min.js ao lado do HTML.

## 📌 Versões
- xlsx.mini.min.js → SheetJS 0.18.5 (mini - suporta escrita XLSX)
- chart.umd.min.js → Chart.js 4.4.0
- pptxgen.bundle.js → PptxGenJS 3.12.0
- chartjs-plugin-datalabels.min.js → 2.2.0
