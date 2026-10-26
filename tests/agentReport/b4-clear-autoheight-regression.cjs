'use strict';
// Pure memory unit tests. The fake high-level clear reproduces the observed SDK side effect.
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'../..'),babel=require(path.join(root,'node_modules/@babel/core')),plugin=require(path.join(root,'node_modules/@babel/plugin-transform-modules-commonjs')),fixtureSdk=require('./clear-content-sdk-fixture.cjs'),cache=new Map();
function load(file){if(cache.has(file))return cache.get(file);const m={exports:{}};vm.runInNewContext(babel.transformSync(fs.readFileSync(file,'utf8'),{configFile:false,babelrc:false,plugins:[plugin]}).code,{module:m,exports:m.exports,console,performance,require:n=>fixtureSdk.modules[n]||load(path.resolve(path.dirname(file),n+'.js'))});cache.set(file,m.exports);return m.exports;}
const api=load(path.join(root,'src/views/application/agentReport/controlled-write-executor.js')),clone=v=>JSON.parse(JSON.stringify(v));
let passed=0,failed=0;
async function test(name,fn){try{await fn();passed++;console.log('PASS '+name);}catch(e){failed++;console.error('FAIL '+name+': '+e.stack);}}
function fixture(options={}){
 const cell={v:0,t:3,s:{fs:16,bg:{rgb:'#abcdef'}},...options.cell};
 const snapshot={id:'unit',sheets:{s:{id:'s',name:'Sheet1',rowCount:30,columnCount:10,cellData:{9:{1:cell,2:{v:'outside',t:1}}},rowData:{9:{ah:26,h:35,ia:1}},columnData:{1:{w:90}},mergeData:[]}},resources:[{name:'KEEP',data:'unchanged'}]};
 let highLevel=0;
 const range={clearContent(){highLevel++;for(const k of ['v','t','p','f','si'])delete cell[k];snapshot.sheets.s.rowData[9].ah=24;},getValues:()=>[[cell.v??null]],getFormulas:()=>[[cell.f||'']]};
 const sheet={getSheetId:()=> 's',getSheetName:()=> 'Sheet1',getMaxRows:()=>30,getMaxColumns:()=>10,getRange:()=>range};
 const sdk=fixtureSdk.attach(sheet,{...options,mutate:params=>{for(const [row,cols]of Object.entries(params.cellValue))for(const [col,value]of Object.entries(cols)){const target=snapshot.sheets.s.cellData[row][col];for(const [k,v]of Object.entries(value))if(v===null)delete target[k];if(value.v===null)delete target.t;}}});
 const book={getSheetBySheetId:id=>id==='s'?sheet:null,getActiveSheet:()=>sheet,save:()=>clone(snapshot)};
 const manifest={transportVersion:2,actions:[{schemaVersion:'2.0',type:'clearRange',clearWhat:'content',target:{sheetId:'s',range:'B10'}}],cells:[{sheetId:'s',row:9,column:1,value:null,operation:'clear-content',before:clone(cell)}]};
 return{book,manifest,snapshot,cell,sdk,highLevel:()=>highLevel,sheet};
}
(async()=>{
 await test('false/font16 clear preserves ah=26 and every non-content field',async()=>{const f=fixture(),before=clone(f.snapshot),r=await api.executeControlledWrite({},f.book,f.manifest);assert.equal(r.ok,true);assert.equal(r.cells[0].value,null);const expected=clone(before);for(const k of ['v','t','p','f','si'])delete expected.sheets.s.cellData[9][1][k];assert.deepEqual(f.snapshot,expected);assert.equal(f.highLevel(),0);assert.equal(f.sdk.sent.length,1);assert.equal(f.sdk.history.length,1);});
 for(const value of [false,0,'','text',42])await test('content-only clear preserves dimensions for '+JSON.stringify(value),async()=>{const f=fixture({cell:{v:value,t:typeof value==='boolean'?3:typeof value==='number'?2:1}});await api.executeControlledWrite({},f.book,f.manifest);assert.equal(f.snapshot.sheets.s.rowData[9].ah,26);assert.equal(f.highLevel(),0);});
 await test('rich text and formula clear preserve style',async()=>{const f=fixture({cell:{p:{body:{dataStream:'old\r\n'}},f:'=1+1',si:'shared'}});const style=clone(f.cell.s);await api.executeControlledWrite({},f.book,f.manifest);assert.deepEqual(f.cell.s,style);for(const k of ['v','p','f','si','t'])assert.equal(f.cell[k],undefined);});
 for(const options of [{permitted:false},{missingMutation:true},{intercepted:{redos:[{id:'other'}],undos:[]}}])await test('unsafe SDK preflight zero writes '+JSON.stringify(options),async()=>{const f=fixture(options),before=clone(f.snapshot);await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest));assert.deepEqual(f.snapshot,before);assert.equal(f.sdk.sent.length,0);assert.equal(f.highLevel(),0);});
 for(const mutationResult of [false,undefined,null,{},Promise.resolve(true)])await test('non-true mutation result never reports success '+typeof mutationResult,async()=>{const f=fixture({mutationResult});await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest));assert.equal(f.sdk.history.length,0);assert.equal(f.highLevel(),0);});
 await test('missing bridge refuses before supported prefix writes',async()=>{const f=fixture();delete f.sheet.getInject;const before=clone(f.snapshot);f.manifest.actions.unshift({schemaVersion:'2.0',type:'setCell',target:{sheetId:'s',range:'C10'},value:'new'});f.manifest.cells.unshift({sheetId:'s',row:9,column:2,value:'new',before:{v:'outside',t:1}});await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest));assert.deepEqual(f.snapshot,before);assert.equal(f.highLevel(),0);});
 await test('mutation exception and undo exception cannot yield successful receipt',async()=>{for(const options of [{mutationThrows:true},{undoThrows:true}]){const f=fixture(options);let receipts=0;await api.submitControlledCall({applyControlled:()=>api.executeControlledWrite({},f.book,f.manifest)},{callId:'c',claimToken:'fixture',argsHash:'h',fencingToken:2},'r',async(_p,body)=>{receipts++;assert.equal(body.ok,false);return{status:'PAUSED',persistenceStatus:'NOT_SAVED'};},()=>true);assert.equal(receipts,1);assert.equal(f.highLevel(),0);}});
 console.log(JSON.stringify({passed,failed,realSdk:false,realBrowser:false,realDatabase:false}));process.exitCode=failed?1:0;
})();
