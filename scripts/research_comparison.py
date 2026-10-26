"""Comparación explícita entre compras: correspondencias visuales y un mismo MAT."""
import re
import unicodedata


def category_key(label):
    text = re.sub(r'\.[CV]$', '', str(label), flags=re.I)
    text = re.sub(r'^T\.\s*', '', text, flags=re.I)
    text = ''.join(ch for ch in unicodedata.normalize('NFD', text.lower()) if not unicodedata.combining(ch))
    key = re.sub(r'[^a-z0-9]', '', text)
    # Equivalencias de etiquetas de canales; no agrupan ni suman categorías.
    return {'autoservicos':'autosservicos', 'farmaciadrogaria':'farmacia',
            'perfumariafranquia':'perfumaria', 'lojadedepart':'lojasdedepto',
            'lojasdedepart':'lojasdedepto'}.get(key, key)


def alignment(configs):
    if len(configs) != 2 or not all(c.data.attrs.get('source_kind') == 'Compras' for c in configs):
        return None
    left, right = configs
    columns = [list(c.data.columns[1:]) for c in configs]
    groups = []
    for labels in columns:
        grouped = {}
        for label in labels:
            grouped.setdefault(category_key(label), []).append(label)
        groups.append(grouped)
    pairs = []
    used = set()
    for label in columns[0]:
        key = category_key(label)
        candidates = groups[1].get(key, [])
        other = candidates[0] if len(candidates) == len(groups[0][key]) == 1 else None
        if other: used.add(other)
        pairs.append((label, other))
    pairs.extend((None, label) for label in columns[1] if label not in used)
    return {'labels':[c.data.attrs.get('source_label', c.display_tipo) for c in configs], 'pairs':pairs}


def palette(configs, colors):
    model = alignment(configs)
    if not model: return {}
    return {label:colors[index % len(colors)] for index,pair in enumerate(model['pairs']) for label in pair if label}


def mat_model(configs, pd, cutoff=None):
    model = alignment(configs)
    if not model: return None
    frames = [c.data.set_index(c.data.columns[0]).copy() for c in configs]
    for frame in frames:
        frame.index = pd.to_datetime(frame.index).to_period('M')
    if any(frame.index.has_duplicates for frame in frames):
        return dict(model, ready=False, note='No se compara MAT con meses duplicados.')
    common = frames[0].index.intersection(frames[1].index)
    if cutoff is not None: common = common[common <= pd.Period(cutoff, freq='M')]
    if common.empty: return dict(model, ready=False, note='No hay meses comunes entre las dos fuentes.')
    end = common.max()
    months = pd.period_range(end=end, periods=12, freq='M')
    if not months.isin(common).all():
        return dict(model, ready=False, note='El comparativo necesita 12 meses comunes consecutivos.')
    values = []
    for frame in frames:
        numeric = frame.loc[months].apply(pd.to_numeric, errors='coerce')
        sums = numeric.sum(min_count=12)
        values.append({label:float(value) if pd.notna(value) else None for label,value in sums.items()})
    def row(left, right):
        a, b = values[0].get(left), values[1].get(right)
        delta = b-a if a is not None and b is not None else None
        percent = delta/a*100 if delta is not None and a != 0 else 0.0 if delta == 0 else None
        return {'left_label':left, 'right_label':right, 'label':re.sub(r'\.[CV]$', '', str(left or right)),
                'base':a, 'comparison':b, 'delta':delta, 'percent':percent}
    rows = [row(a,b) for a,b in model['pairs']]
    totals = [sum(data.values()) if all(value is not None for value in data.values()) else None for data in values]
    a,b = totals
    total = {'label':'Total', 'base':a, 'comparison':b, 'delta':b-a if a is not None and b is not None else None,
             'percent':(b/a-1)*100 if a not in (None,0) and b is not None else 0.0 if a == b == 0 else None}
    return dict(model, ready=True, start=str(months[0]), end=str(end), rows=rows, total=total,
                note='Diferencia entre fuentes del mismo MAT; no es variación interanual. Sin equivalente no significa cero.')


def figure(model, engine, language='E'):
    """Dos paneles alineados: niveles MAT y diferencia porcentual frente al bloque base."""
    pt = language == 'P'
    fig = engine.plt.figure(figsize=(12.6, 3.8), facecolor='white')
    left = fig.add_axes((.20,.16,.42,.70)); right = fig.add_axes((.73,.16,.22,.70))
    rows = list(model['rows']) + [model['total']]
    positions = list(range(len(rows)))
    def fmt(value):
        if value is None: return '—'
        result = f'{value:+.1f}%'
        return result.replace('.',',') if pt else result
    for index,row in enumerate(rows):
        a,b = row['base'],row['comparison']
        if a is not None:left.barh(index-.16,a,height=.28,color='#2c3e50')
        if b is not None:left.barh(index+.16,b,height=.28,color='#078c86')
        delta = row['percent']
        if delta is None:
            right.text(0,index,'Sem base comparável' if pt else 'Sin base comparable',fontsize=7.5,va='center',color='#637888')
        else:
            right.barh(index,delta,height=.55,color='#3c9c63' if delta>0 else '#ce5b58' if delta<0 else '#8c9aa3')
            right.annotate(fmt(delta),(delta,index),xytext=(4 if delta>=0 else -4,0),textcoords='offset points',
                           ha='left' if delta>=0 else 'right',va='center',fontsize=8,fontweight='bold')
    labels=[engine.textwrap.fill(re.sub(r'^T\.\s*','',row['label']),26) for row in rows]
    left.set_yticks(positions,labels,fontsize=8.5)
    for axis in (left,right):
        axis.set_ylim(len(rows)-.5,-.5); axis.set_axisbelow(True); axis.grid(axis='x',color='#e3eaee',linewidth=.6)
        axis.spines[['top','right','left']].set_visible(False);axis.tick_params(axis='y',length=0)
        axis.tick_params(axis='x',labelsize=8)
    right.set_yticks(positions,[]);right.axvline(0,color='#859aa8',linewidth=.8)
    right.margins(x=.35); left.margins(x=.10)
    max_value=max([r[k] or 0 for r in rows for k in ('base','comparison')])
    divisor,suffix=(1e6,'M') if max_value>=1e6 else (1e3,'K') if max_value>=1e3 else (1,'')
    left.xaxis.set_major_formatter(engine.FuncFormatter(lambda v,_:f'{v/divisor:g}{suffix}'))
    left.set_title('Volume MAT' if pt else 'Volumen MAT',fontsize=11,fontweight='bold',loc='left',pad=12)
    right.set_title('Diferença vs. base (%)' if pt else 'Diferencia vs. base (%)',fontsize=11,fontweight='bold',loc='left',pad=12)
    from matplotlib.patches import Patch
    fig.legend(handles=[Patch(color='#2c3e50',label=model['labels'][0]),Patch(color='#078c86',label=model['labels'][1])],
               loc='lower left',bbox_to_anchor=(.20,.01),ncol=2,frameon=False,fontsize=8.5)
    return engine.figure_to_stream(fig,bbox_inches=None,transparent=False)
