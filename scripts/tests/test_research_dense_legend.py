"""Leyendas densas: más área de trazado, texto legible y mismas observaciones."""
import unittest
from unittest import mock

from studio_5w1h_vendor import report_engine as engine


class DenseLegendTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        engine.ensure_heavy_modules_loaded()

    def render(self, labels, width=12.6, multi=False, previous=False):
        frame = engine.pd.DataFrame({'Compras': engine.pd.date_range('2024-01-01', periods=32, freq='MS')})
        for index, label in enumerate(labels):
            frame[label] = [100 + index * 30 + month * (index + 1) for month in range(32)]
        measurements = {}
        original_export = engine.figure_to_stream

        def inspect(figure, **kwargs):
            figure.canvas.draw()
            renderer = figure.canvas.get_renderer()
            ax = figure.axes[0]
            legend = ax.get_legend()
            box = legend.get_window_extent(renderer)
            measurements.update(columns=legend._ncols, labels=[t.get_text() for t in legend.get_texts()],
                font=min(t.get_fontsize() for t in legend.get_texts()),
                plot_height=ax.get_position().height, legend_box=box.bounds, canvas=figure.bbox.bounds,
                plot_bottom=ax.get_window_extent(renderer).y0,
                lines=[list(line.get_ydata()) for line in ax.lines])
            return original_export(figure, **kwargs)

        colors = {}
        with mock.patch.object(engine, 'figure_to_stream', side_effect=inspect):
            if previous:
                with mock.patch.object(engine, '_trend_legend_columns', side_effect=lambda labels, width, font, default, fig: min(len(labels), default)), \
                     mock.patch.object(engine, '_trend_legend_height_inches', side_effect=lambda ax, handles, labels, cols, font: font * 1.35 / 72 * engine.math.ceil(len(labels) / cols)):
                    stream = engine.line_graf(frame, 0, '', 9801, 0, width_emu=engine.Inches(width),
                        height_emu=engine.Cm(10), multi_chart=multi, color_collector=colors, show_title=False)
            else:
                stream = engine.line_graf(frame, 0, '', 9801, 0, width_emu=engine.Inches(width),
                    height_emu=engine.Cm(10), multi_chart=multi, color_collector=colors, show_title=False)
        self.assertEqual(stream.getvalue()[:8], b'\x89PNG\r\n\x1a\n')
        measurements['colors'] = colors
        return measurements

    def test_many_short_labels_gain_plot_height_without_changing_values_or_colors(self):
        labels = [f'{index} Und' for index in range(1, 33)]
        old = self.render(labels, previous=True)
        new = self.render(labels)
        self.assertEqual(old['columns'], 4)
        self.assertGreaterEqual(new['columns'], 8)
        self.assertGreater(new['plot_height'], old['plot_height'] * 1.25)
        self.assertEqual(new['labels'], labels)
        self.assertEqual(new['lines'], old['lines'])
        self.assertEqual(new['colors'], old['colors'])
        self.assertGreaterEqual(new['font'], 8)
        x, y, width, height = new['legend_box']
        self.assertGreaterEqual(x, 0)
        self.assertGreaterEqual(y, 0)
        self.assertLessEqual(x + width, new['canvas'][2] + 1)
        self.assertLess(y + height, new['plot_bottom'])

    def test_columns_adapt_to_label_length_and_half_width_panels(self):
        short = [f'{index} Und' for index in range(1, 25)]
        wide = self.render(short)
        narrow = self.render(short, width=6, multi=True)
        long = self.render([f'Presentación familiar {index} unidades' for index in range(1, 25)])
        self.assertLess(narrow['columns'], wide['columns'])
        self.assertLess(long['columns'], wide['columns'])
        self.assertEqual(len(narrow['labels']), 24)
        self.assertEqual(len(long['labels']), 24)
        for layout in [narrow, long]:
            x, y, width, height = layout['legend_box']
            self.assertGreaterEqual(y, 0)
            self.assertLessEqual(x + width, layout['canvas'][2] + 1)
            self.assertLess(y + height, layout['plot_bottom'])

    def test_small_legends_keep_their_original_layout(self):
        labels = ['Moderno', 'Tradicional', 'Farmacia', 'Otros']
        for multi in [False, True]:
            old = self.render(labels, width=6 if multi else 12.6, multi=multi, previous=True)
            new = self.render(labels, width=6 if multi else 12.6, multi=multi)
            self.assertEqual(new['columns'], old['columns'])
            self.assertEqual(new['plot_height'], old['plot_height'])


if __name__ == '__main__':
    unittest.main()
