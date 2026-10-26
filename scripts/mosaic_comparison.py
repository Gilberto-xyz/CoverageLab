"""Comparación por parejas Original/Mosaic, usando los cálculos del explorador."""
import hashlib
import re
import unicodedata


def _base(name):
    name = unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z0-9]', '', re.sub(r'^p[0-6]_', '', re.sub(r'\(mosaic\)', '', name)))


def detect_pairs(sheets):
    """El nombre ayuda; las fechas, unidad y Sell-in deben coincidir siempre.

    Una firma única de datos resuelve abreviaturas sin emparejar por posición.
    Las coincidencias ambiguas quedan fuera para evitar comparar otra marca.
    """
    originals = [s for s in sheets if not s.get('error') and not re.search(r'\(mosaic\)', s['name'], re.I)]
    mosaics = [s for s in sheets if not s.get('error') and re.search(r'\(mosaic\)', s['name'], re.I)]
    pairs, used = [], set()
    for original in originals:
        candidates = [s for s in mosaics if s['name'] not in used
                      and s.get('unit') == original.get('unit')
                      and s.get('dates') and s['dates'] == original.get('dates')
                      and s.get('sell_in') == original.get('sell_in')]
        exact = [s for s in candidates if _base(s['name']) == _base(original['name'])]
        candidates = exact if exact else candidates
        if len(candidates) != 1:
            continue
        mosaic = candidates[0]
        used.add(mosaic['name'])
        pairs.append({'id': hashlib.sha256(original['name'].encode()).hexdigest()[:16],
                      'original': original['name'], 'mosaic': mosaic['name'],
                      'label': original.get('brand', original['name']),
                      'category': original.get('category', ''), 'unit': original.get('unit', '')})
    return pairs


def selected_pairs(analysis, selections):
    sheets = {s['name']: s for s in analysis['sheets']}
    pairs = []
    for pair in analysis.get('mosaic_pairs', []):
        if any(pair[key] not in selections for key in ('original', 'mosaic')):
            continue
        dates = []
        for key in ('original', 'mosaic'):
            name = pair[key]
            p = selections[name]
            dates.append({r['date'] for r in sheets[name]['dashboard']['mat_periods']
                          if all(v is not None for v in (r['worldpanel'], r['sell_in'][p], r['coverage'][p]))})
        common = sorted(dates[0] & dates[1])
        if common:
            pairs.append({**pair, 'dates': common,
                          'original_pipeline': selections[pair['original']],
                          'mosaic_pipeline': selections[pair['mosaic']]})
    return pairs


def write_presentation(core, template, target, pairs, images, progress=None, language='ES'):
    """Extiende el motor PPT existente con las capturas comparativas de la web."""
    ppt = core.Presentation(str(template))
    ppt.core_properties.language={'PT':'pt-BR','EN':'en-US','ES':'es-MX'}.get(language,'es-MX')
    for index in reversed(range(len(ppt.slides))):
        core._delete_slide(ppt, index)
    width, height = ppt.slide_width, ppt.slide_height
    for index, pair in enumerate(pairs):
        if progress:
            progress('comparison', index, len(pairs), pair['label'])
        slide = ppt.slides.add_slide(ppt.slide_layouts[core.PPT_LAYOUT_INDEX])
        slide.name = f"Original-Mosaic: {pair['label']}"
        for shape in list(slide.shapes):
            if shape.is_placeholder:
                shape._element.getparent().remove(shape._element)
        # Área superior libre para el contenido, respetando el pie de la plantilla.
        from PIL import Image
        with Image.open(images[pair['id']]) as capture:
            iw, ih = capture.size
        scale = min(width*.96/iw, height*.90/ih)
        pw, ph = int(iw*scale), int(ih*scale)
        slide.shapes.add_picture(str(images[pair['id']]), int((width-pw)/2), int(height*.015), width=pw, height=ph)
        notes = (
            f"Original: {pair['original']} · P{pair['original_pipeline']}\n"
            f"Mosaic: {pair['mosaic']} · P{pair['mosaic_pipeline']}\n"
            f"Cortes MAT compartidos: {', '.join(pair['dates'])}. "
            "Mismas escalas de volumen y cobertura. Correlaciones r en porcentaje, "
            "ventanas y pipelines con datos. Rango de color independiente por matriz. "
            "Debajo: variaciones MAT de Sell-in P0–P6 y Worldpanel frente al MAT anterior, "
            "en porcentaje. Verde aumenta, rojo disminuye. "
            "Datos y selección de pipelines del archivo de entrada registrado en _sistema/selection.json.")
        if language=='PT':
            notes=(f"Original: {pair['original']} · P{pair['original_pipeline']}\n"
                   f"Mosaic: {pair['mosaic']} · P{pair['mosaic_pipeline']}\n"
                   f"Cortes MAT comuns: {', '.join(pair['dates'])}. Eixos de volume e cobertura comuns. "
                   "Correlações r em porcentagem, com escalas de cores próprias para cada matriz. "
                   "Variações MAT de Sell-in P0–P6 e Worldpanel em relação ao MAT anterior. "
                   "Verde indica aumento e vermelho indica queda. "
                   "Fonte e seleção de pipelines registradas em _sistema/selection.json.")
        elif language=='EN':
            notes=(f"Original: {pair['original']} · P{pair['original_pipeline']}\n"
                   f"Mosaic: {pair['mosaic']} · P{pair['mosaic_pipeline']}\n"
                   f"Shared MAT cutoffs: {', '.join(pair['dates'])}. Shared volume and coverage axes. "
                   "Correlation r as a percentage, with separate color scales for each matrix. "
                   "Sell-in P0–P6 and Worldpanel MAT changes vs. previous MAT. "
                   "Green indicates growth; red indicates decline. Source and pipeline choices in _sistema/selection.json.")
        slide.notes_slide.notes_text_frame.text=notes
    ppt.save(str(target))
    if progress:
        progress('comparison', len(pairs), len(pairs), 'PPT Original/Mosaic lista')
    return str(target)
