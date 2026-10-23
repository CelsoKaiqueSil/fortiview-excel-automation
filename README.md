# FortiView Excel Automation

[Baixar Excel sintetico de exemplo](docs/examples/traffic_weekly.xlsx) · [Historico de mudancas](CHANGELOG.md)

![Previa dos totais de trafego sintetico](docs/images/preview.png)

**Problema:** atualizar uma planilha recorrente sem duplicar amostras ou perder o arquivo anterior.

- 8 amostras e 2 equipamentos ficticios no exemplo; repetir a entrada insere zero linhas.
- Trava exclusiva e substituicao atomica protegem contra concorrencia e falha de gravacao.
- Correcao de uma amostra atualiza o valor e registra antes/depois na aba Alteracoes.

A imagem e uma previa Python. Fatores de entrada sao binarios (1024), documentados
explicitamente; o resumo usa `total_gib`, nao chama GiB de GB decimal.

[![CI](https://github.com/CelsoKaiqueSil/fortiview-excel-automation/actions/workflows/ci.yml/badge.svg)](https://github.com/CelsoKaiqueSil/fortiview-excel-automation/actions/workflows/ci.yml)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](pyproject.toml)
[![Excel](https://img.shields.io/badge/Excel-openpyxl-217346?logo=microsoftexcel&logoColor=white)](src/fortiview_excel/pipeline.py)
[![PowerShell](https://img.shields.io/badge/PowerShell-agendamento-5391FE?logo=powershell&logoColor=white)](scripts/run_windows.ps1)
[![Graph](https://img.shields.io/badge/Microsoft%20Graph-integração-5E5E5E?logo=microsoft)](.env.example)
[![Data](https://img.shields.io/badge/dados-100%25%20sintéticos-14804A)](data/traffic_sample.csv)

Automação demonstrativa para transformar uma rotina semanal de atualização de
planilhas em um processo agendado, padronizado e protegido por backup.

## Fluxo

```mermaid
flowchart LR
    A["Exportação FortiView<br>sintética"] --> B["Leitura"]
    B --> C["Normalização de unidades"]
    C --> D["Deduplicação"]
    D --> E["Backup do arquivo atual"]
    E --> F["Atualização do Excel"]
    F --> G["OneDrive / SharePoint<br>camada opcional"]
```

## Recursos

- conversão de B, KB, MB, GB e TB para bytes;
- preservação do histórico e prevenção de duplicidades;
- backup com data e hora antes de alterar um arquivo existente;
- páginas `Dados` e `Resumo`;
- cabeçalhos, filtros, congelamento de painel e larguras configuradas;
- execução pelo Windows Task Scheduler ou cron;
- configuração externa por variáveis de ambiente.

## Executar

```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
python -m pip install -e ".[dev]"
fortiview-excel --input data/traffic_sample.csv --output output/traffic_weekly.xlsx
pytest -q
```

No Windows PowerShell: `.\.venv\Scripts\Activate.ps1` em lugar de `source`.
Sem ativacao: `.\.venv\Scripts\python.exe -m fortiview_excel.pipeline`.
Use o mesmo Python com `-m pytest`; nao e necessario alterar politicas globais.

Resultado esperado na primeira execucao: `inserted=8`; na segunda, `inserted=0` e
backup distinto. `python scripts/validate_exports.py` confere os arquivos.
`python scripts/preview.py` regenera a imagem dos totais.

## Concorrencia e dados corrigidos

Uma segunda execucao falha enquanto existir a trava `.xlsx.lock`. Nao apague uma
trava sem confirmar que o processo anterior terminou. Nao ha remocao automatica
de trava por idade. Salve e feche o Excel antes de atualizar no Windows.

A identidade e `collected_at + device + interface + direction`. Conflitos dentro
da mesma entrada bloqueiam a atualizacao; dados corrigidos numa execucao posterior
substituem o valor e preservam o anterior na auditoria e no backup. Nao rode duas
copias em pastas sincronizadas esperando que a trava local funcione como trava distribuida.

No Windows, o script [run_windows.ps1](scripts/run_windows.ps1) pode ser associado
ao Agendador de Tarefas. Em Linux, use [run_linux.sh](scripts/run_linux.sh).

## Segurança

A edição pública não realiza login em firewall, não contém IPs reais e não envia
arquivos a um tenant. A integração com Microsoft Graph é apresentada como fronteira
arquitetural e deve utilizar uma identidade de aplicativo com o menor privilégio
possível.

## Contexto profissional

O estudo representa uma automação recorrente de tráfego enviado e recebido,
mantendo o padrão visual da planilha operacional e reduzindo cópias manuais.
