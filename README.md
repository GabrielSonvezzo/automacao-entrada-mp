# 🚀 Automação de Entrada de Matéria Prima (XML para Excel)

**Autor:** Gabriel Sonvezzo  
**Versão:** 1.0.0 (Turbo Edition)  
**Data:** Janeiro / 2026

Este sistema foi idealizado e desenvolvido por **Gabriel Sonvezzo** para eliminar a digitação manual de notas fiscais. O robô realiza a leitura inteligente de arquivos XML (NF-e) e integra os dados diretamente em planilhas Excel, garantindo 100% de precisão e alta performance.

## 🛠️ Tecnologias Utilizadas

* **Python 3.10+**
* **CustomTkinter**: Interface gráfica (GUI) moderna com modo dark.
* **Openpyxl**: Manipulação de planilhas Excel.
* **XMLtodict**: Conversão eficiente de dados XML.
* **Pillow (PIL)**: Renderização da identidade visual.

## 📋 Funcionalidades de Destaque

* **Modo Turbo**: Carregamento de lotes em memória via dicionários, permitindo o processamento instantâneo de grandes volumes de notas.
* **Captura Inteligente de Lote**: Localização automática do lote através da tag de transporte (`<nVol>`), resolvendo falhas comuns em bobinas galvanizadas.
* **Interface Responsiva**: Log em tempo real e barra de progresso para acompanhamento do status.
* **Segurança de Fluxo**: Movimentação de arquivos apenas após a confirmação de gravação no Excel.

## 🚀 Como Utilizar

1. Coloque os arquivos XML na pasta `Notas_Pendentes`.
2. Certifique-se de que o arquivo `modelo_analise.xlsx` está na raiz com a aba "Arcelor Usina".
3. Execute o `interface.py` ou o executável gerado.
4. Clique em **"LANÇAR NOTAS NA PLANILHA"**.


## 🏗️ Estrutura do Projeto

* `interface.py`: Gerenciamento da interface e threads.
* `robo_notas.py`: Motor de lógica e processamento de dados.
* `logo.png`: Logotipo da empresa centralizado.
* `modelo_analise.xlsx`: Template para integração de dados.

---
Desenvolvido com foco em eficiência e automação por **Gabriel Sonvezzo**.