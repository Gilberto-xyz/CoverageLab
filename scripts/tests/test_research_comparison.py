from datetime import datetime
import io
import json
from pathlib import Path
import unittest

import research_comparison as comparison
from studio_5w1h import engine


class PurchaseComparisonTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls): engine.ensure_heavy_modules_loaded()

    def configs(self, zero=False):
        dates=engine.pd.date_range('2023-01-01',periods=36,freq='MS')
        left=engine.pd.DataFrame({'Fecha':dates,'T. Venda Direta.C':[100]*36,'Norte.C':[0 if zero else 100]*36})
        right=engine.pd.DataFrame({'Fecha':dates,'Norte.V':[50]*36,'Venda Direta.V':[150]*36,'Internet.V':[20]*36})
        for frame,label in [(left,'Compras'),(right,'Compras_Mosaico')]:frame.attrs.update(source_kind='Compras',source_label=label)
        return [engine.SeriesConfig(left,'Compras','Compras',0),engine.SeriesConfig(right,'Ventas','Ventas',0)]

    def test_same_mat_signed_differences_and_unmatched_are_not_zero(self):
        configs=self.configs()
        model=comparison.mat_model(configs,engine.pd)
        self.assertTrue(model['ready'])
        self.assertEqual((model['start'],model['end']),('2025-01','2025-12'))
        first,second,unmatched=model['rows']
        self.assertEqual((first['base'],first['comparison'],first['delta'],first['percent']),(1200,1800,600,50))
        self.assertEqual(second['percent'],-50)
        self.assertIsNone(unmatched['base']); self.assertIsNone(unmatched['percent'])
        self.assertEqual(model['total']['base'],2400);self.assertEqual(model['total']['comparison'],2640)
        self.assertAlmostEqual(model['total']['percent'],10)
        colors=comparison.palette(configs,engine.TREND_COLOR_SEQUENCE)
        self.assertEqual(colors['T. Venda Direta.C'],colors['Venda Direta.V'])
        self.assertEqual(colors['Norte.C'],colors['Norte.V'])
        self.assertNotEqual(colors['Internet.V'],colors['Norte.V'])
        zero=comparison.mat_model(self.configs(zero=True),engine.pd)
        self.assertIsNone(zero['rows'][1]['percent'])
        json.dumps(zero,allow_nan=False)

    def test_visual_channel_aliases_are_explicit_and_ambiguous_names_stay_unmatched(self):
        for left,right in [('T. Autosserviços.C','Auto Servicos.V'),('T. Farmacia.C','Farmacia/Drogaria.V'),
                           ('T. Perfumaria/Franquia.C','Perfumaria.V'),('T. Loja de Depart..C','Lojas De Depto.V')]:
            self.assertEqual(comparison.category_key(left),comparison.category_key(right))
            self.assertEqual(engine.build_semantic_label_signature(left).compact,engine.build_semantic_label_signature(right).compact)
        configs=self.configs()
        configs[0].data['Venda Direta.C']=configs[0].data['T. Venda Direta.C']
        pairs=comparison.alignment(configs)['pairs']
        self.assertIn(('T. Venda Direta.C',None),pairs)
        self.assertIn(('Venda Direta.C',None),pairs)
        self.assertIn((None,'Venda Direta.V'),pairs)

    def test_common_cutoff_missing_data_and_duplicate_dates(self):
        configs=self.configs()
        configs[1].data.drop(index=35,inplace=True)
        model=comparison.mat_model(configs,engine.pd)
        self.assertEqual(model['end'],'2025-11')
        limited=comparison.mat_model(configs,engine.pd,'2024-05-01')
        self.assertEqual(limited['end'],'2024-05')
        configs[1].data.loc[30,'Norte.V']=float('nan')
        self.assertIsNone(comparison.mat_model(configs,engine.pd)['rows'][1]['comparison'])
        configs[1].data.loc[29,'Fecha']=configs[1].data.loc[28,'Fecha']
        self.assertFalse(comparison.mat_model(configs,engine.pd)['ready'])

    def test_combined_trend_preserves_paired_colors_styles_and_reduces_legend(self):
        configs=self.configs()
        frame=engine.combine_research_series(configs)
        pairs=comparison.alignment(configs)['pairs']
        keys={label:a or b for a,b in pairs for label in (a,b) if label}
        palette=comparison.palette(configs,engine.TREND_COLOR_SEQUENCE)
        captured={}
        original=engine.figure_to_stream
        def capture(fig,**kwargs):
            captured['lines']=[(line.get_label(),line.get_color(),line.get_linestyle()) for line in fig.axes[0].lines if line.get_label() in frame.columns]
            captured['legend']=[text.get_text() for text in fig.axes[0].get_legend().get_texts()]
            return original(fig,**kwargs)
        from unittest.mock import patch
        with patch.object(engine,'figure_to_stream',capture):
            engine.line_graf(frame,0,'Comparación',1,2,color_overrides=palette,comparison_keys=keys,
                             linestyle_overrides={col:'-' if col.endswith('.C') else '--' for col in frame.columns[1:]},show_title=False)
        lines={label:(color,style) for label,color,style in captured['lines']}
        self.assertEqual(lines['T. Venda Direta.C'][0],lines['Venda Direta.V'][0])
        self.assertEqual(lines['T. Venda Direta.C'][1],'-');self.assertEqual(lines['Venda Direta.V'][1],'--')
        self.assertEqual(len(captured['legend']),len(pairs)+2)
        self.assertIn('Compras_Mosaico',captured['legend'])


if __name__=='__main__':unittest.main()
