from pathlib import Path
import shutil
import subprocess
import unittest

ROOT=Path(__file__).resolve().parents[1]


class ExportLocalizationTests(unittest.TestCase):
    def test_all_extra_models_and_mosaic_terms_support_pt_es_en(self):
        node=shutil.which('node')
        if not node:self.skipTest('Node no está disponible')
        script=r'''
const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict');
const c={window:{}};vm.createContext(c);vm.runInContext(fs.readFileSync('explorer_web/i18n.js','utf8'),c);
const i18n=c.window.CoverageI18n,pt=i18n.get('PT'),en=i18n.get('EN'),es=i18n.get('ES');
assert.equal(pt.name,'Português (Brasil)');assert.equal(pt.number(12.9,1,'%'),'12,9%');
assert.equal(en.number(12.9,1,'%'),'12.9%');assert.equal(pt.signed(-5.95),'-6,0%');
assert.equal(pt.month('2026-01-01'),'jan. de 26');assert.equal(en.month('2026-01-01'),'Jan 26');
assert.equal(pt.correlationLabel('Año actual · 12 meses'),'Ano atual · 12 meses');
assert.equal(pt.correlationLabel('Año actual · 12 meses',true),'Atual · 12m');
assert.equal(en.correlationLabel('Últimos 24 meses'),'Last 24 months');
assert.equal(pt.category('Cremas Corporales'),'Cremes corporais');assert.equal(pt.mosaic,'Expansão Mosaic');
const titles={annual_mat:'Volume MAT, YoY e cobertura por ano',frequency:'Frequência',
 purchase_average:'Compra média',purchase_occasion:'Compra por ocasião',penetration:'Penetração',buyers_series:'Compradores'};
for(const [key,title] of Object.entries(titles)){
 const input={key,title:'español',description:'español',kpi:key==='annual_mat'?undefined:{},series:[{name:'Cobertura P2',values:[1,2]}]};
 const before=JSON.stringify(input),model=pt.model(input);
 assert.equal(model.title,title);assert.ok(!model.description.includes('español'));assert.equal(JSON.stringify(input),before);
 assert.ok(!en.model(input).description.includes('español'));
}
assert.ok(pt.model({key:'monthly',description:'UNITS · Sell-in desplazado 2 mes(es)',series:[]}).description.includes('deslocado 2'));
assert.ok(en.model({key:'monthly',description:'kg · Sell-in desplazado 2 mes(es)',series:[]}).description.includes('shifted by 2'));
assert.equal(pt.model({key:'coverage',series:[{name:'Penetración mensual'}]}).series[0].name,'Penetração mensal');
assert.equal(es.title('annual_mat'),'Volumen MAT, YoY y cobertura por año');
'''
        subprocess.run([node,'-e',script],cwd=ROOT,check=True,capture_output=True,text=True)
