'use strict'
const assert=require('assert'),fs=require('fs'),path=require('path'),{createRequire}=require('module')
const root=path.resolve(__dirname,'../..'),rf=createRequire(path.join(root,'package.json')),babel=rf('@babel/core'),cache=new Map()
function load(file){if(cache.has(file))return cache.get(file).exports;const module={exports:{}};cache.set(file,module);const source=babel.transformSync(fs.readFileSync(file,'utf8'),{filename:file,babelrc:false,configFile:false,plugins:[rf.resolve('@babel/plugin-transform-modules-commonjs')]}).code;new Function('require','module','exports',source)(name=>name.startsWith('.')?load(path.resolve(path.dirname(file),name+'.js')):rf(name),module,module.exports);return module.exports}
const controlled=load(path.join(root,'src/views/application/agentReport/controlled-write-executor.js'))
const core=rf('@univerjs/core/lib/cjs/index.js'),sheets=rf('@univerjs/sheets/lib/cjs/index.js'),engine=rf('@univerjs/engine-formula/lib/cjs/index.js'),formulaPlugin=rf('@univerjs/sheets-formula/lib/cjs/index.js')
rf('@univerjs/sheets/lib/cjs/facade.js');rf('@univerjs/sheets-formula/lib/cjs/facade.js')
const {FUniver}=rf('@univerjs/core/lib/cjs/facade.js')
function setup(){const univer=new core.Univer({locale:core.LocaleType.EN_US,locales:{[core.LocaleType.EN_US]:{}}});univer.registerPlugin(engine.UniverFormulaEnginePlugin);univer.registerPlugin(sheets.UniverSheetsPlugin);univer.registerPlugin(formulaPlugin.UniverSheetsFormulaPlugin);univer.createUnit(core.UniverInstanceType.UNIVER_SHEET,{id:'b7',name:'B7',sheetOrder:['s'],sheets:{s:{id:'s',name:'Sheet',rowCount:8,columnCount:5,cellData:{0:{0:{v:2,t:2},1:{v:3,t:2}},1:{0:{v:4,t:2},1:{v:5,t:2}}}}}});const api=FUniver.newAPI(univer),book=api.getActiveWorkbook();return{univer,api,book}}
const action=(range,formula)=>({schemaVersion:'2.0',type:'setFormula',target:{sheetId:'s',range},formula})
function manifest(a,cells){const patch={sheetId:'s',range:a.target.range,cells};return{transportVersion:2,action:a,actions:[a],formulaPatch:patch,formulaReadback:cells}}
async function main(){
  const fill=action('C1:C2','=A1+B1'),sum=action('C1','=SUM(A1:B2)')
  assert.throws(()=>controlled.classifyAiActions([fill]),/AI_CAPABILITY_BLOCKED/)
  assert.equal(controlled.classifyAiActions([fill],['setFormula']),'controlled-values')
  assert.throws(()=>controlled.classifyAiActions([fill,fill],['setFormula']),/AI_CAPABILITY_BLOCKED/)
  for(const [a,cells] of [[fill,[{row:0,column:2,formula:'=A1+B1',value:5},{row:1,column:2,formula:'=A2+B2',value:9}]],[sum,[{row:0,column:2,formula:'=SUM(A1:B2)',value:14}] ]]){
    const {univer,api,book}=setup();try{const before=book.save(),m=manifest(a,cells),result=await controlled.executeControlledWrite(api,book,m);assert.equal(result.ok,true);assert.deepStrictEqual(result.formulaReadback,cells);assert.equal(controlled.sameControlledFormulaTransition(before,book.save(),m.formulaPatch),true);const bad=JSON.parse(JSON.stringify(book.save()));bad.sheets.s.cellData[4]={0:{v:99,t:2}};assert.equal(controlled.sameControlledFormulaSnapshot(bad,book.save(),m.formulaPatch),false);console.log('PASS B7 actual SDK '+a.formula)}finally{univer.dispose()}
  }
  const readOnly=setup();try{const before=readOnly.book.save();assert.deepStrictEqual(readOnly.book.getSheetBySheetId('s').getRange('C1').getFormulas(),[['']]);assert.deepStrictEqual(readOnly.book.save(),before)}finally{readOnly.univer.dispose()}
  const wrong=setup();try{const cells=[{row:0,column:2,formula:'=A1+B1',value:6},{row:1,column:2,formula:'=A2+B2',value:9}];await assert.rejects(controlled.executeControlledWrite(wrong.api,wrong.book,manifest(fill,cells)),/RESULT_UNKNOWN/);console.log('PASS B7 wrong calculation cache cannot return successful receipt')}finally{wrong.univer.dispose()}
}
main().catch(error=>{console.error(error);process.exitCode=1})
