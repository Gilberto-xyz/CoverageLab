import copy
from pathlib import Path
import shutil
import subprocess
import unittest

from mosaic_comparison import detect_pairs, selected_pairs

ROOT=Path(__file__).resolve().parents[1]


class MosaicComparisonTests(unittest.TestCase):
    def sheet(self, name):
        return {'name':name,'brand':name,'unit':'UNITS','dates':['2024-01-01','2024-02-01'],
                'sell_in':[123,456],'error':None,'dashboard':{'mat_periods':[
                    {'date':'2025-06-01','worldpanel':80,'sell_in':[100]*7,'coverage':[80]*7},
                    {'date':'2026-06-01','worldpanel':90,'sell_in':[120]*7,'coverage':[75]*7}]}}

    def test_pairing_uses_data_to_resolve_abbreviations_independently_of_order(self):
        a=self.sheet('P0_BDCR_Natura Corpo+Óleo')
        b=self.sheet('P0_BDCR_Natura Cpo+Óleo(Mosaic)')
        c=self.sheet('P1_TOIL_Natura');c['sell_in']=[1,2]
        d=copy.deepcopy(c);d['name']='P1_TOIL_Natura(Mosaic)'
        pairs=detect_pairs([b,c,a,d])
        self.assertEqual([(p['original'],p['mosaic']) for p in pairs],[(c['name'],d['name']),(a['name'],b['name'])])
        self.assertEqual(len({p['id'] for p in pairs}),2)

    def test_ambiguous_or_different_inputs_never_form_a_pair(self):
        a=self.sheet('P0_Natura');b=self.sheet('P0_MosaicName(Mosaic)')
        c=copy.deepcopy(b);c['name']='P0_Other(Mosaic)'
        self.assertEqual(detect_pairs([a,b,c]),[])
        for key,value in [('dates',['2024-03-01']),('sell_in',[999,456]),('unit','KG'),('error','Sin datos')]:
            changed=copy.deepcopy(b);changed[key]=value
            self.assertEqual(detect_pairs([a,changed]),[])

    def test_selection_requires_both_sheets_and_uses_each_pipeline_and_common_cutoffs(self):
        a=self.sheet('P0_Natura');b=self.sheet('P0_Natura(Mosaic)')
        b['dashboard']['mat_periods'][0]['sell_in'][3]=None
        analysis={'sheets':[a,b],'mosaic_pairs':detect_pairs([a,b])}
        self.assertEqual(selected_pairs(analysis,{a['name']:0}),[])
        pair=selected_pairs(analysis,{a['name']:0,b['name']:3})[0]
        self.assertEqual(pair['dates'],['2026-06-01'])
        self.assertEqual((pair['original_pipeline'],pair['mosaic_pipeline']),(0,3))

    def test_browser_models_share_scales_and_dates_without_changing_inputs(self):
        node=shutil.which('node')
        if not node:self.skipTest('Node no está disponible')
        script=r'''
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const context={window:{}};vm.createContext(context);
vm.runInContext(fs.readFileSync('explorer_web/dashboard.js','utf8'),context);
context.CoverageDashboard=context.window.CoverageDashboard;
vm.runInContext(fs.readFileSync('explorer_web/mosaic.js','utf8'),context);
const api=context.window.MosaicComparison;
function sheet(name,wp){return {name,dates:[],pipelines:Array.from({length:7},()=>({})),
 dashboard:{kpis:[],mat_periods:['2025-06-01','2026-06-01'].map((date,i)=>({date,worldpanel:wp+i,
 worldpanel_yoy:i?3:null,sell_in:Array.from({length:7},(_,p)=>100+p),sell_in_yoy:Array(7).fill(i?4:null),coverage:Array(7).fill(wp)}))},
 template:{correlations:[{label:'Actual',pipelines:[.1,null,.3,null,null,null,null]},
 {label:'Anterior',pipelines:Array(7).fill(null)}]}};}
const a=sheet('A',40),b=sheet('B',90),pair={original:'A',mosaic:'B',id:'x'};
b.template.correlations[0].pipelines=[.8,null,.95,null,null,null,null];
b.dashboard.mat_periods[0].sell_in[3]=null;
const analysis={sheets:[a,b],mosaic_pairs:[pair],chart_colors:{}},before=JSON.stringify(analysis);
const prepared=api.prepare(analysis,{A:0,B:3},pair);
assert.deepEqual([...prepared.dates],['2026-06-01']);
assert.equal(prepared.models[0].series[0].values[0],100);
assert.equal(prepared.models[1].series[0].values[0],103);
assert.equal(prepared.models[0].scales.left.max,prepared.models[1].scales.left.max);
assert.equal(prepared.models[0].scales.right.max,prepared.models[1].scales.right.max);
assert.deepEqual([...prepared.rows],['Actual']);assert.deepEqual([...prepared.pipelines],[0,1,2,3,4,5,6]);
assert.equal(prepared.variations[0].worldpanel,3);assert.equal(prepared.variations[1].pipelines[3],4);
assert.deepEqual([...prepared.samples[0]],[.1,.3]);
assert.deepEqual([...prepared.samples[1]],[.8,.95]);
const style=context.CoverageDashboard.correlationStyle,palette=['#F8696B','#FFEB84','#63BE7B'];
for(const samples of prepared.samples){
 assert.equal(style(Math.min(...samples),samples,palette).fill,palette[0]);
 assert.equal(style(Math.max(...samples),samples,palette).fill,palette[2]);
}
// La escala de la original no depende del nivel de correlación de Mosaic.
b.template.correlations[0].pipelines=[-.9,null,-.8,null,null,null,null];
const changed=api.prepare(analysis,{A:0,B:3},pair);
assert.equal(style(.3,prepared.samples[0],palette).fill,style(.3,changed.samples[0],palette).fill);
b.template.correlations[0].pipelines=[.8,null,.95,null,null,null,null];
assert.equal(api.eligible(analysis,{A:0}).length,0);assert.equal(api.eligible(analysis,{A:0,B:3}).length,1);
assert.equal(JSON.stringify(analysis),before);
// Reutiliza la variación validada del template, que ya viene en porcentaje.
const templateCase=JSON.parse(JSON.stringify(analysis));
templateCase.sheets.forEach(s=>{s.dates=['2026-06-01T00:00:00'];s.template.variations=[{
 label:'MAT vs. MAT anterior',worldpanel:-1.4,pipelines:[-9.3,-8.2,-5.95,0,2,null,null]}];});
const mat=api.prepare(templateCase,{A:0,B:3},pair).variations;
assert.equal(mat[0].worldpanel,-1.4);assert.equal(mat[0].pipelines[2],-5.95);
assert.equal(mat[1].pipelines[3],0);assert.equal(mat[1].pipelines[5],null);
assert.equal(api.variationPercent(-5.95),'-6.0%');assert.equal(api.variationPercent(.5),'+0.5%');
assert.equal(api.variationPercent(0),'0.0%');assert.equal(api.variationPercent(null),'—');
assert.equal(api.variationStyle(-1).fill,'#FFC7CE');assert.equal(api.variationStyle(1).fill,'#C6EFCE');
assert.equal(api.variationStyle(0).fill,api.variationStyle(null).fill);
const custom={positive:{fill:'#123456',ink:'#FFFFFF'}};
assert.equal(api.variationStyle(1,custom).fill,'#123456');
'''
        subprocess.run([node,'-e',script],cwd=ROOT,check=True,capture_output=True,text=True)
