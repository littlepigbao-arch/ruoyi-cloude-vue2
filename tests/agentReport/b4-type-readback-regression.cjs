'use strict';
// Offline-only fault injection: a setter/getter can agree on v but retain a wrong Univer t.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert');
const root=path.resolve(__dirname,'../..'),babel=require(path.join(root,'node_modules/@babel/core')),plugin=require(path.join(root,'node_modules/@babel/plugin-transform-modules-commonjs')),cache=new Map();
function load(file){if(cache.has(file))return cache.get(file);const m={exports:{}};vm.runInNewContext(babel.transformSync(fs.readFileSync(file,'utf8'),{configFile:false,babelrc:false,plugins:[plugin]}).code,{module:m,exports:m.exports,console,performance,require:n=>require('./clear-content-sdk-fixture.cjs').modules[n]||load(path.resolve(path.dirname(file),n+'.js'))});cache.set(file,m.exports);return m.exports;}
const api=load(path.join(root,'src/views/application/agentReport/controlled-write-executor.js'));
const clone=v=>JSON.parse(JSON.stringify(v));let passed=0,failed=0;
async function test(name,fn){try{await fn();passed++;console.log('PASS '+name);}catch(e){failed++;console.error('FAIL '+name+': '+e.message);}}
function fixture(options={}){
 const state={
  id:'unit',resources:[{name:'KEEP',data:'unchanged'}],
  sheets:{s:{id:'s',name:'Orders',rowCount:20,columnCount:10,cellData:{
   0:{0:{v:'outside',t:1}},1:{1:{v:'old',t:1,s:{bg:{rgb:'#abc'}}}}
  }}}
 };
 const cell=()=>state.sheets.s.cellData[1][1];let writes=0;
 function set(value){writes++;const typed=value!==null&&typeof value==='object'?value:{v:value,t:typeof value==='number'?2:typeof value==='boolean'?3:1};Object.assign(cell(),typed);if(options.missingType)delete cell().t;else if(Object.hasOwn(options,'wrongType'))cell().t=options.wrongType;if(options.numericBoolean&&typeof cell().v==='boolean')cell().v=cell().v?1:0;}
 const range={setValue:set,setValueForCell:set,setValues:values=>set(values[0][0]),clearContent(){writes++;for(const k of ['v','t','p','f','si'])delete cell()[k];},getValues:()=>[[cell().v??null]],getFormulas:()=>[['']]};
 const sheet={getSheetId:()=> 's',getSheetName:()=> 'Orders',getMaxRows:()=>20,getMaxColumns:()=>10,getRange:()=>range};require('./clear-content-sdk-fixture.cjs').attach(sheet);
 const book={getSheetBySheetId:id=>id==='s'?sheet:null,getActiveSheet:()=>sheet,getSheets:()=>[sheet],save:()=>clone(state)};
 const manifest=(type,value)=>({transportVersion:2,actions:[{schemaVersion:'2.0',type,target:{sheetId:'s',range:'B2'},range:'B2',...(type==='clearRange'?{clearWhat:'content'}:type==='setValues'?{values:[[value]]}:{value})}],cells:[{sheetId:'s',row:1,column:1,before:clone(cell()),value:type==='clearRange'?null:value,...(type==='clearRange'?{operation:'clear-content'}:{})}]});
 return{book,cell,state,manifest,writes:()=>writes};
}
(async()=>{
 for(const [name,value,wrongType] of [['numeric text','001',2],['number',42,1],['boolean false',false,2],['text tagged as boolean','plain',3],['text tagged as rich text','plain',4],['unknown type','plain',99],['string type tag','plain','1']])
  await test('alias rejects correct getter/value but wrong stored type: '+name,async()=>{const f=fixture({wrongType});await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest('setCellForCell',value)));assert.equal(f.writes(),1);});
 await test('missing scalar type is uncertainty, not a verified alias',async()=>{const f=fixture({missingType:true});await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest('setCellForCell','001')));});
 await test('setCell and setValues share the same type readback boundary',async()=>{for(const type of ['setCell','setValues']){const f=fixture({wrongType:2});await assert.rejects(()=>api.executeControlledWrite({},f.book,f.manifest(type,'001')));}});
 await test('both entry receipts remain failed when scalar type is inconsistent',async()=>{for(const entrance of ['ordinary-ai','agent']){const f=fixture({wrongType:2}),receipts=[];const result=await api.submitControlledCall({applyControlled:()=>api.executeControlledWrite({},f.book,f.manifest('setCellForCell','001'))},{callId:entrance,claimToken:'fixture',argsHash:'fixture',fencingToken:1},entrance,async(endpoint,body)=>{receipts.push(body);return{status:'PAUSED',persistenceStatus:'NOT_SAVED'};},()=>true);assert.equal(receipts.length,1);assert.strictEqual(receipts[0].ok,false);assert.equal(result.persistenceStatus,'NOT_SAVED');assert.equal(f.writes(),1);}});
 await test('verified scalar tags preserve values, styles, resources and workbook identity',async()=>{for(const value of ['',0,false,true,'plain','001',42]){const f=fixture({numericBoolean:true}),before=clone(f.state),book=f.book,result=await api.executeControlledWrite({},f.book,f.manifest('setCellForCell',value));assert.strictEqual(result.cells[0].value,value);assert.strictEqual(f.book,book);assert.deepStrictEqual(f.cell().s,before.sheets.s.cellData[1][1].s);assert.deepStrictEqual(f.state.resources,before.resources);assert.deepStrictEqual(f.state.sheets.s.cellData[0],before.sheets.s.cellData[0]);}});
 await test('content clear remains a null result without a scalar type tag',async()=>{const f=fixture(),r=await api.executeControlledWrite({},f.book,f.manifest('clearRange'));assert.strictEqual(r.cells[0].value,null);assert.strictEqual(f.cell().t,undefined);assert.equal(f.cell().s.bg.rgb,'#abc');});
 console.log(JSON.stringify({passed,failed,scope:'OFFLINE_MEMORY_FACADE_NO_DATABASE_NO_SERVICE',realBrowser:false,realDatabase:false}));process.exitCode=failed?1:0;
})();
