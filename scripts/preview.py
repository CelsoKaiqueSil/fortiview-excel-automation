"""Preview deterministic synthetic traffic totals, expressed in GiB."""
from pathlib import Path
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from fortiview_excel.pipeline import load_traffic

ROOT = Path(__file__).resolve().parents[1]
frame = load_traffic(ROOT / 'data/traffic_sample.csv')
data = frame.groupby(['device', 'direction'])['bytes'].sum().unstack() / 1024**3
fig, ax = plt.subplots(figsize=(13, 7), facecolor='#f6f8f9')
fig.subplots_adjust(top=.70, bottom=.20, left=.10, right=.95)
fig.text(.06, .93, 'FortiView Excel Automation', fontsize=24, weight='bold', color='#173631')
fig.text(.06, .85, '8 amostras | 2 equipamentos sinteticos | reaplicacao sem duplicidade', fontsize=13, color='#52636b')
data.plot.bar(ax=ax, color=['#147e57', '#008aab'], rot=0)
ax.set_facecolor('#f6f8f9')
ax.set_ylabel('GiB acumulados na amostra')
ax.set_xlabel('Equipamento')
ax.spines[['top', 'right']].set_visible(False)
ax.legend(title='Direcao', frameon=False)
fig.text(.06, .07, 'Previa Python | fatores binarios (1024) | backup + trava exclusiva + gravacao atomica + auditoria', fontsize=11, color='#52636b')
output = ROOT / 'docs/images/preview.png'
output.parent.mkdir(parents=True, exist_ok=True)
fig.savefig(output, dpi=140)
plt.close(fig)
