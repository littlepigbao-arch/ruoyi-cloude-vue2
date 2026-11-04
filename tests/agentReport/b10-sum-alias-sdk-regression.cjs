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
const alias=()=>({schemaVersion:'2.0',type:'sumToCell',target:{sheetId:'s',range:'C1'},sourceRange:'A1:B2'})
function sumManifest(a,value=14){const c=controlled.normalizeControlledFormulaAction(a),cells=[{row:0,column:2,formula:c.formula,value}];return{transportVersion:2,action:a,actions:[a],formulaPatch:{sheetId:'s',range:c.target.range,cells},formulaReadback:cells}}
async function main(){
 const base=alias(),old=alias(),both=alias();old.targetRange='C1';delete old.target.range;both.targetRange='C1';const single={...alias(),sourceRange:'A1'}
 assert.throws(()=>controlled.classifyAiActions([base],['setFormula']),/AI_CAPABILITY_BLOCKED/)
 assert.equal(controlled.classifyAiActions([base],['setFormula','sumToCell']),'controlled-values')
 assert.throws(()=>controlled.classifyAiActions([base,base],['setFormula','sumToCell']),/AI_CAPABILITY_BLOCKED/)
 let passed=3
 for(const [a,value]of [[base,14],[old,14],[both,14],[single,2]]){const{univer,api,book}=setup();try{const before=book.save(),m=sumManifest(a,value),result=await controlled.executeControlledWrite(api,book,m);assert.equal(result.ok,true);assert.deepStrictEqual(result.formulaReadback,m.formulaReadback);assert.equal(book.getSheetBySheetId('s').getRange('C1').getValues()[0][0],value);assert.equal(controlled.sameControlledFormulaTransition(before,book.save(),m.formulaPatch),true);assert.equal(m.actions[0].type,'sumToCell');await assert.rejects(controlled.executeControlledWrite(api,book,m),/目标不再为空/);passed++;console.log('PASS sumToCell installed SDK '+JSON.stringify(a))}finally{univer.dispose()}}
 const invalid=[{...alias(),formula:'=1+1'},{...alias(),targetRange:'D1'},{...alias(),sourceRange:'Other!A1:B2'},{...alias(),sourceRange:'B2:A1'},{...alias(),sourceRange:'A1:A1001'},{...alias(),sourceRange:'A1:C1'},{...alias(),schemaVersion:'1.0'},{...alias(),target:{sheetId:'s',range:'C1:C2'}},{...alias(),sourceRange:null},{...alias(),background:'#123456'}]
 for(const a of invalid){const{univer,api,book}=setup();try{const before=book.save();assert.throws(()=>controlled.normalizeControlledFormulaAction(a));assert.deepStrictEqual(book.save(),before);passed++}finally{univer.dispose()}}
 const wrong=setup();try{await assert.rejects(controlled.executeControlledWrite(wrong.api,wrong.book,sumManifest(alias(),15)),/RESULT_UNKNOWN/);passed++}finally{wrong.univer.dispose()}
 console.log(JSON.stringify({passed,failures:0,installedSdk:true,realDatabase:false}))
}
main().catch(e=>{console.error(e);process.exitCode=1})
