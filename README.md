# Projeto de Extração, Análise e Relatórios de Ensaios de Curva IxV

Este repositório reúne scripts, notebooks e planilhas utilizados no âmbito do Consórcio Marangatu (UFV), para automação e análise de ensaios de Curva IxV, conforme divulgado nos posts:

- [Consórcio Marangatú - Motrice Soluções](https://www.linkedin.com/posts/darieldon-de-brito-medeiros_ontem-no-cons%C3%B3rcio-marangatu-da-motrice-solu%C3%A7%C3%B5es-activity-7206242399352766464-g2oK?utm_source=social_share_send&utm_medium=member_desktop_web&rcm=ACoAACZ_e70BjDTM-fIO9bX--d7Woxcl08Tr3AE)
- [UFV Marangatú - Ensaios Automatizados](https://www.linkedin.com/posts/darieldon-de-brito-medeiros_boa-tarde-hoje-na-ufv-marangatu-constru%C3%ADdas-activity-7214685368128815105-LR0A?utm_source=social_share_send&utm_medium=member_desktop_web&rcm=ACoAACZ_e70BjDTM-fIO9bX--d7Woxcl08Tr3AE)

## Objetivo

Automatizar o processo de extração de dados de ensaios de Curva IxV, tratamento de imagens, extração de informações via OCR e geração de relatórios padronizados em planilhas Excel.

## Estrutura dos Arquivos

- **Notebooks (`.ipynb`)**:

  - `extImagem - MODELO 0X.ipynb` e `extImagem - MODELO XX.ipynb`: Automatizam a extração de imagens de arquivos PDF de relatórios de ensaio, recorte de regiões específicas das imagens, extração de texto via OCR (Tesseract) e inserção dos dados extraídos em planilhas Excel.
  - `MODELO ENTEC.ipynb` e `MODELO SOLMETRIC.ipynb`: Realizam o tratamento de arquivos CSV gerados por diferentes equipamentos de ensaio (ENTEC e SOLMETRIC), extraindo dados relevantes, formatando tabelas e gerando gráficos das curvas IxV (Corrente x Tensão) e PxV (Potência x Tensão), além de salvar imagens e tabelas para compor os relatórios.

- **Planilhas Excel (`.xlsm`)**:
  - `PLANILHA DE RENOMEAR - MODELO 0X.xlsm` e `PLANILHA DE RENOMEAR - MODELO XX.xlsm`: Auxiliam na organização e renomeação automática dos arquivos PDF dos ensaios, servindo de base para os scripts de extração.
  - `Relatório de Ensaio de Curva IxV - ENTEC.xlsm`, `Relatório de Ensaio de Curva IxV - SOLMETRIC.xlsm`, `Relatório de Ensaio de Curva IxV - MARXX-PCSY INVZZ.xlsm`: Modelos de relatórios finais, onde os dados extraídos e tratados são inseridos automaticamente pelos notebooks, padronizando a apresentação dos resultados dos ensaios.

## Bibliotecas Necessárias (Python)

Para executar os notebooks, é necessário instalar as seguintes bibliotecas Python:

- `pikepdf` — Extração de imagens de arquivos PDF.
- `openpyxl` — Manipulação de arquivos Excel (XLSX).
- `xlwings` — Integração avançada com Excel, inclusive arquivos com macros (XLSM).
- `opencv-python` (`cv2`) — Processamento e recorte de imagens.
- `pytesseract` — OCR para extração de texto de imagens.
- `Pillow` (`PIL`) — Manipulação de imagens.
- `re` — Expressões regulares (já incluída no Python).
- `pandas` — Manipulação e análise de dados tabulares.
- `matplotlib` — Geração de gráficos.
- `chardet` — Detecção de encoding de arquivos.

### Instalação das Bibliotecas

Você pode instalar todas as bibliotecas necessárias com o seguinte comando:

```bash
pip install pikepdf openpyxl xlwings opencv-python pytesseract Pillow pandas matplotlib chardet
```

> **Observação:**
>
> - Para o funcionamento do `pytesseract`, é necessário instalar o Tesseract-OCR no seu sistema operacional e garantir que o caminho para o executável esteja correto no script.
> - O `xlwings` requer que o Microsoft Excel esteja instalado na máquina.

## Fluxo de Trabalho

1. **Extração de Imagens e Dados**

   - Os notebooks de extração abrem os PDFs dos relatórios, extraem imagens das páginas desejadas e recortam regiões específicas contendo dados relevantes.
   - Utilizam OCR (Tesseract) para converter imagens em texto, extraindo valores como irradiância, temperatura, data e hora do ensaio.

2. **Tratamento e Análise dos Dados**

   - Os notebooks de análise (ENTEC/SOLMETRIC) processam arquivos CSV, extraem parâmetros elétricos (Pm, Voc, Isc, Vm, Im, FF, Irradiância, Temperatura), formatam tabelas e geram gráficos das curvas IxV e PxV.

3. **Geração de Relatórios**
   - Os dados extraídos e analisados são automaticamente inseridos nas planilhas-modelo, gerando relatórios padronizados e prontos para apresentação ou arquivamento.

## Como Utilizar

1. Certifique-se de ter Python e as bibliotecas necessárias instaladas.
2. Execute os notebooks conforme o modelo do equipamento (ENTEC ou SOLMETRIC) e o padrão do relatório (0X ou XX).
3. Os arquivos de entrada (PDFs, CSVs) devem estar na mesma pasta dos notebooks ou conforme especificado nos scripts.
4. Os relatórios finais serão gerados automaticamente nas planilhas-modelo correspondentes.

## Sobre o Projeto

Este projeto faz parte das atividades do Consórcio Marangatu, em parceria com a Motrice Soluções e a BN Engenharia, visando a automação e padronização dos ensaios de Curva IxV, promovendo eficiência, rastreabilidade e qualidade nos resultados.

---

## Contato

Quaisquer dúvidas, entrem em contato comigo:

[![Email](https://img.shields.io/badge/Email-darieldonbm99@outlook.com-blue?style=for-the-badge&logo=microsoftoutlook)](mailto:darieldonbm99@outlook.com)
[![WhatsApp](<https://img.shields.io/badge/WhatsApp-+55%20(86)%2099466--6480-25D366?style=for-the-badge&logo=whatsapp>)](https://wa.me/5586994666480)
