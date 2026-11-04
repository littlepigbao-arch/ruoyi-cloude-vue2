const assert=require('assert')
const fs=require('fs'),path=require('path'),{createRequire}=require('module')
const root=path.resolve(__dirname,'../..'),rf=createRequire(path.join(root,'package.json')),babel=rf('@babel/core'),cache=new Map()
function load(file){if(cache.has(file))return cache.get(file).exports;const module={exports:{}};cache.set(file,module);const source=babel.transformSync(fs.readFileSync(file,'utf8'),{filename:file,babelrc:false,configFile:false,plugins:[rf.resolve('@babel/plugin-transform-modules-commonjs')]}).code;new Function('require','module','exports',source)(name=>name.startsWith('.')?load(path.resolve(path.dirname(file),name+'.js')):rf(name),module,module.exports);return module.exports}
const {submitControlledCall}=load(path.join(root,'src/views/application/agentReport/controlled-write-executor.js'))
async function run(){
  const call={callId:'call-1',effectId:'effect-stable',claimToken:'claim',argsHash:'args',fencingToken:3}
  const requests=[];let executions=0
  const result=await submitControlledCall({applyControlled:async()=>{executions++;return{ok:true,cells:[]}}},call,'run-1',async(url,body)=>{requests.push({url,body});return{status:'COMPLETED',persistenceStatus:'SAVED'}},()=>true)
  assert.strictEqual(result.status,'COMPLETED');assert.strictEqual(executions,1);assert.strictEqual(requests[0].body.effectId,'effect-stable');assert.strictEqual(requests[0].body.callId,'call-1')
  console.log('PASS B8 stable effect identity passes through the shared receipt transport')
  const legacy={...call};delete legacy.effectId
  const old=[];await submitControlledCall({applyControlled:async()=>({ok:true,cells:[]})},legacy,'run-2',async(url,body)=>{old.push(body);return{status:'COMPLETED',persistenceStatus:'SAVED'}},()=>true)
  assert.strictEqual(Object.prototype.hasOwnProperty.call(old[0],'effectId'),false)
  console.log('PASS B8 old controlled calls keep the legacy receipt shape')
}
run().catch(e=>{console.error(e);process.exitCode=1})
