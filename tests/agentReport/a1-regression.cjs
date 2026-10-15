// Run real source methods with Vue's reactive watchers and in-memory boundaries.
// This is not a real-browser / MySQL end-to-end test.
const fs=require('fs'),path=require('path'),vm=require('vm'),assert=require('assert');
const root=path.resolve(__dirname,'../..');
const babel=require(path.join(root,'node_modules/@babel/core'));
const plugin=require(path.join(root,'node_modules/@babel/plugin-transform-modules-commonjs'));
const Vue=require(path.join(root,'node_modules/vue'));
const sourceRoot=process.env.A1_SOURCE_ROOT||root;
const dir=path.join(sourceRoot,'src/views/application/agentReport');
let api={},wbApi={},storage=new Map(),cache=new Map(),failures=0,passed=0;
function load(file){
 if(cache.has(file))return cache.get(file);
 let src=fs.readFileSync(file,'utf8');if(file.endsWith('.vue'))src=src.match(/<script>([\s\S]*?)<\/script>/)[1];
 const module={exports:{}};
 const code=babel.transformSync(src,{configFile:false,babelrc:false,plugins:[plugin]}).code;
 vm.runInNewContext(code,{module,exports:module.exports,console,performance,Blob,sessionStorage:{setItem:(k,v)=>storage.set(k,v)},
 setTimeout:()=>0,clearTimeout:()=>{},setInterval:()=>0,clearInterval:()=>{},
 require:name=>{
  if(name==='@/api/ai/agent')return new Proxy({}, {get:(_,k)=>api[k]});
  if(name==='@/api/ai/workbook')return new Proxy({}, {get:(_,k)=>wbApi[k]});
  if(name.startsWith('.')){const target=path.resolve(path.dirname(file),name+(path.extname(name)?'':'.js'));if(target.endsWith('.vue')&&!target.endsWith('AgentTaskPanel.vue'))return {};return load(target);}
  return {};
 }},{filename:file});cache.set(file,module.exports);return module.exports;
}
const executor=load(path.join(dir,'univer-action-executor.js'));
let writes=0;
const makeSheet=(id,name)=>({getSheetId:()=>id,getSheetName:()=>name,getMaxRows:()=>100,getMaxColumns:()=>26,
 getRange:()=>({setValue:()=>writes++,setValues:()=>writes++})});
const a=makeSheet('a','Orders'),b=makeSheet('b','Arrivals');
const book={getActiveSheet:()=>b,getSheets:()=>[a,b],getSheetBySheetId:id=>({a,b}[id]),getSheetByName:name=>({Orders:a,Arrivals:b}[name])};
async function test(name,fn){try{writes=0;await fn();passed++;console.log('PASS '+name);}catch(e){failures++;console.error('FAIL '+name+': '+e.message);}}
const defer=()=>{let resolve,reject;const promise=new Promise((a,b)=>{resolve=a;reject=b});return {promise,resolve,reject}};
const plain=v=>JSON.parse(JSON.stringify(v));
function panel(){
 const component=load(path.join(dir,'components/AgentTaskPanel.vue')).default;
 const locks=[];let reloads=0;
 const workspace={lock:v=>locks.push(v),reload:async()=>{reloads++;return {revision:'new'}},saved:async()=>{},prepare:async()=>({})};
 const instance=new Vue({...component,propsData:{workbookId:1,workspace}});
 instance.enabled=true;instance.recoveryAvailable=true;instance.generation=1;instance.syncRequired=true;instance.run={runId:'r',status:'EXECUTING'};
 instance.schedule=()=>{};
 return {instance,locks,workspace,reloads:()=>reloads};
}
function page(){
 const c=load(path.join(dir,'index.vue')).default;
 const obj=c.data();Object.entries(c.methods).forEach(([k,f])=>obj[k]=f.bind(obj));
 Object.assign(obj,{currentDocumentId:1,currentBranchId:2,headCommitId:'h',branchVersion:4,draftVersion:7,loadedRevision:'old',
 workspaceGeneration:1,loadedSnapshotFingerprint:JSON.stringify({v:'old'}),getSnapshot:()=>({v:'old'}),$nextTick:async()=>{},$modal:{msgError:()=>{},msgSuccess:()=>{}},
 runPersistTask:async(_title,fn)=>fn(()=>{}),manualCommands:[],pendingChanges:[{changeId:'old'}],pendingActions:[]});
 return obj;
}
(async()=>{
 await test('V2 missing ID: zero writes',()=>{const r=executor.executeActions({},book,[{schemaVersion:'2.0',type:'setCell',target:{range:'A1'},value:1}]);assert.equal(r.summary.ok,0);assert.equal(writes,0);});
 await test('V2 invalid ID cannot fall back to active Sheet',()=>{const r=executor.executeAction({},book,{schemaVersion:'2.0',type:'setCell',target:{sheetId:'missing',range:'A1'},value:1});assert(!r.ok);assert.equal(writes,0);});
 await test('V2 ID/name/flat target conflicts are rejected',()=>{for(const extra of [{sheetName:'Arrivals'},{}]){const r=executor.executeAction({},book,{schemaVersion:'2.0',type:'setCell',sheetId:extra.sheetName?undefined:'b',target:{sheetId:'a',range:'A1',...extra},value:1});assert(!r.ok)}assert.equal(writes,0);});
 await test('Invalid later V2 target prevents any group writes',()=>{const r=executor.executeActions({},book,[{schemaVersion:'2.0',type:'setCell',target:{sheetId:'a',range:'A1'},value:1},{schemaVersion:'2.0',type:'setCell',target:{range:'A1'},value:2}]);assert.equal(r.summary.ok,0);assert.equal(writes,0);});
 await test('V1 active/name/prefixed and valid V2 stay compatible',()=>{for(const action of [{type:'setCell',range:'A1',value:1},{schemaVersion:'1.0',type:'setCell',range:'Orders!A1',value:1},{type:'setCell',sheetName:'Orders',range:'A1',value:1},{schemaVersion:'2.0',type:'setCell',target:{sheetId:'a',range:'A1'},value:1}])assert(executor.executeActions({},book,[action]).results[0].ok);assert.equal(writes,4);});
 await test('Completion watcher cannot unlock before content reload',async()=>{const p=panel(),d=defer();p.workspace.reload=()=>d.promise;api.getAgent=async()=>({runId:'r',status:'COMPLETED'});api.agentRequest=async()=>[];const polling=p.instance.poll();await Vue.nextTick();await Promise.resolve();assert(!p.locks.includes(false));d.resolve({revision:'new'});await polling;await Vue.nextTick();assert.equal(p.instance.busy,false);assert.equal(p.locks.at(-1),false);});
 await test('Sync failure retains protection',async()=>{const p=panel();p.workspace.reload=async()=>{throw Error('offline')};api.getAgent=async()=>({runId:'r',status:'FAILED'});api.agentRequest=async()=>[];await p.instance.poll();await Vue.nextTick();assert(p.instance.busy);assert(!p.locks.includes(false));});
 await test('Old run response cannot replace newer page or unlock',async()=>{const p=panel(),d=defer();api.getAgent=()=>d.promise;const polling=p.instance.poll();p.instance.generation++;p.instance.run={runId:'new',status:'EXECUTING'};d.resolve({runId:'r',status:'COMPLETED'});await polling;assert.equal(p.instance.run.runId,'new');assert(!p.locks.includes(false));});
 await test('Cancellation terminal state reloads before unlock',async()=>{const p=panel();api.agentRequest=async()=>({runId:'r',status:'CANCELLED'});await p.instance.stop();assert.equal(p.reloads(),1);assert.equal(p.locks.at(-1),false);});
 await test('Disabled entry can attach and cancel existing run',async()=>{const p=panel();p.instance.enabled=false;api.activeAgent=async()=>({runId:'r',status:'PAUSED'});await p.instance.attach();assert.equal(p.instance.run.runId,'r');api.agentRequest=async()=>({runId:'r',status:'CANCELLED'});await p.instance.stop();assert.equal(p.locks.at(-1),false);});
 await test('Recover UNKNOWN cancels, reloads and reconciles without replay',async()=>{const p=panel();p.instance.run={runId:'r',status:'PAUSED',call:{status:'UNKNOWN'}};const calls=[];api.agentRequest=async(url)=>{calls.push(url);return {runId:'r',status:url.endsWith('/cancel')?'CANCELLING':'CANCELLED'}};await p.instance.recover();assert.equal(calls.length,2);assert(calls[1].endsWith('/resume'));assert.equal(p.reloads(),2);assert.equal(p.locks.at(-1),false);});
 await test('Version dialog does not bless old local content',()=>{const p=page();p.applyVersionStatus({branchId:2,version:99,headCommitId:'new',loadedRevision:'new',workingState:{version:99}});assert.equal(p.draftVersion,7);assert.equal(p.loadedRevision,'old');assert.equal(p.branchVersion,4);});
 await test('Restoring without a draft loads committed content before adopting version',async()=>{const p=page();wbApi.getVersionStatus=async()=>({code:200,data:{branchId:2,version:5,headCommitId:'new',content:'{"v":"committed"}',loadedRevision:'new',workingState:null}});let loaded;p.loadUniverWorkbook=s=>{loaded=s;p.getSnapshot=()=>s};await p.loadVersionStatus(true);assert.equal(loaded.v,'committed');assert.equal(p.loadedRevision,'new');assert.equal(p.draftVersion,-1);});
 await test('Autosave sends loaded versions and advances only its own response',async()=>{const p=page();let sent;wbApi.saveWorkingState=async(id,payload)=>{sent=payload;return {code:200,data:{version:8,contentHash:'saved'}}};await p.performAutosaveWorkingState(true);assert.equal(sent.expectedDraftVersion,7);assert.equal(sent.loadedRevision,'old');assert.equal(p.draftVersion,8);});
 await test('Old autosave response cannot bless new workspace',async()=>{const p=page(),d=defer();wbApi.saveWorkingState=()=>d.promise;const saving=p.performAutosaveWorkingState(true);p.currentDocumentId=9;p.workspaceGeneration++;d.resolve({code:200,data:{version:99,contentHash:'new'}});await assert.rejects(saving);assert.equal(p.draftVersion,7);assert.equal(p.loadedRevision,'old');});
 await test('File switch sends original draft version, not latest server version',async()=>{const p=page();let sent;wbApi.saveWorkingState=async(id,payload)=>{sent=payload;return {code:409,msg:'conflict'}};await p.loadDocument({workbookId:9});assert.equal(sent.expectedDraftVersion,7);assert.equal(sent.loadedRevision,'old');assert.equal(p.currentDocumentId,1);});
 await test('Branch checkout saves with the loaded draft version first',async()=>{const p=page();let sent;wbApi.saveWorkingState=async(id,payload)=>{sent=payload;return {code:409,msg:'conflict'}};await p.checkoutSelectedBranch(9);await Promise.resolve();assert.equal(sent.expectedDraftVersion,7);assert.equal(sent.loadedRevision,'old');assert.equal(p.currentBranchId,2);});
 await test('Reload preserves unsaved edits and adopts one content/version tuple',async()=>{const p=page();p.getSnapshot=()=>({v:'local-edit'});let applied;api.agentWorkspace=async()=>({workbookId:1,branchId:2,snapshot:{v:'server'},revision:'server',loadedRevision:'server-hash',draftVersion:8,branchVersion:4,headCommitId:'h',workingState:{version:8}});p.applyAgentSnapshot=async s=>{applied=s;p.getSnapshot=()=>s};await p.reloadAgentDraft();assert.equal(p.conflictDrafts[0].snapshot.v,'local-edit');assert.equal(applied.v,'server');assert.equal(p.loadedRevision,'server-hash');assert.equal(p.draftVersion,8);assert.equal(p.agentSyncRequired,false);});
 await test('Old workspace sync is ignored before applying content',async()=>{const p=page(),d=defer();api.agentWorkspace=()=>d.promise;let applied=0;p.applyAgentSnapshot=async()=>applied++;const reload=p.reloadAgentDraft();p.workspaceGeneration++;d.resolve({workbookId:1,branchId:2,snapshot:{},loadedRevision:'new',draftVersion:9});await assert.rejects(reload);assert.equal(applied,0);assert(p.agentLocked);});
 await test('Reload failure cannot unlock page or update version',async()=>{const p=page();api.agentWorkspace=async()=>{throw Error('offline')};await assert.rejects(p.reloadAgentDraft());p.lockAgentWorkspace(false);assert(p.agentLocked);assert.equal(p.draftVersion,7);});
 await test('Failure to persist local recovery copy prevents replacement',async()=>{const p=page();p.getSnapshot=()=>({v:'edit'});const originalSet=storage.set;storage.set=()=>{throw Error('quota')};try{assert.throws(()=>p.preserveLocalDraft());assert.equal(p.conflictDrafts[0].snapshot.v,'edit');}finally{storage.set=originalSet}});
 console.log(JSON.stringify({passed,failures,realBrowser:false,realDatabase:false}));if(failures)process.exitCode=1;
})().catch(e=>{console.error(e);process.exitCode=1});
