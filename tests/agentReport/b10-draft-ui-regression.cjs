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
const defer=()=>{let resolve;const promise=new Promise(r=>{resolve=r});return{promise,resolve}}
function page(){const c=load(path.join(dir,'index.vue')).default,p=c.data();for(const[k,v]of Object.entries(c.methods))p[k]=v.bind(p);p.errors=[];p.snapshot={v:'local'};p.backups=[];Object.assign(p,{currentDocumentId:1,currentBranchId:2,branchVersion:4,headCommitId:'head',draftVersion:7,loadedRevision:'loaded',currentDocumentVersion:11,workspaceGeneration:1,pendingChanges:[],pendingActions:[],branches:[{branchId:3,branchName:'main',headCommitId:'main-head',version:8,draftVersion:-1,loadedRevision:'main-hash'}],getSnapshot:()=>p.snapshot,loadUniverWorkbook:s=>{p.snapshot=s},preserveLocalDraft:force=>p.backups.push({force,snapshot:JSON.parse(JSON.stringify(p.snapshot))}),flushManualChanges(){},openVersionDialog:async()=>{},$confirm:()=>Promise.resolve(),$modal:{msgError:e=>p.errors.push(e),msgWarning:e=>p.errors.push(e),msgSuccess(){}}});p.runPersistTask=(_title,fn)=>fn(()=>{});return p}
async function test(name,fn){try{await fn();passed++;console.log('PASS '+name)}catch(e){failures++;console.error('FAIL '+name+': '+e.stack)}}
;(async()=>{
await test('Explicit discard sends loaded binding and backs up before deleting',async()=>{const p=page();let body;wbApi.discardWorkingState=async(id,b)=>{assert.equal(p.backups.length,1);body=b;return{code:200}};wbApi.getVersionStatus=async()=>({code:200,data:{branchId:2,headCommitId:'head',version:4,content:'{"v":"head"}',loadedRevision:'head-hash',workingState:null}});await p.discardCurrentDraft();assert.equal(body.expectedDraftVersion,7);assert.equal(body.loadedRevision,'loaded');assert.equal(body.confirmed,true);assert(p.backups[0].force);assert.equal(p.snapshot.v,'head');assert.equal(p.draftVersion,-1)})
await test('Discard cancelled by changed confirmation content',async()=>{const p=page(),d=defer();p.$confirm=()=>d.promise;let calls=0;wbApi.discardWorkingState=async()=>{calls++};const job=p.discardCurrentDraft();p.snapshot={v:'new'};d.resolve();await job;assert.equal(calls,0);assert.equal(p.snapshot.v,'new')})
await test('Discard conflict keeps current content and baseline',async()=>{const p=page();wbApi.discardWorkingState=async()=>({code:409,msg:'newer draft'});await p.discardCurrentDraft();assert.equal(p.snapshot.v,'local');assert.equal(p.draftVersion,7);assert.equal(p.backups.length,1);assert(p.errors.includes('newer draft'))})
await test('Late discard readback preserves new edits',async()=>{const p=page(),d=defer();wbApi.discardWorkingState=async()=>({code:200});wbApi.getVersionStatus=()=>d.promise;const job=p.discardCurrentDraft();for(let i=0;i<10;i++)await Promise.resolve();p.snapshot={v:'new-edit'};d.resolve({code:200,data:{branchId:2,content:'{"v":"server"}'}});await job;assert.equal(p.snapshot.v,'new-edit');assert.equal(p.backups.at(-1).snapshot.v,'new-edit');assert.equal(p.draftVersion,7)})
await test('Discard reply cannot overwrite another workspace',async()=>{const p=page(),d=defer();wbApi.discardWorkingState=()=>d.promise;const job=p.discardCurrentDraft();for(let i=0;i<10;i++)await Promise.resolve();p.currentDocumentId=9;p.workspaceGeneration++;p.snapshot={v:'other'};d.resolve({code:200});await job;assert.equal(p.snapshot.v,'other')})
await test('Merge sends frozen source and target draft identities',async()=>{const p=page();let body;wbApi.fastForwardBranchToMain=async(id,branch,b)=>{body=b;assert.equal(branch,2);return{code:200,data:{branchId:3,version:9,headCommitId:'head',content:'{"v":"merged"}',contentHash:'merged-hash'}}};await p.fastForwardCurrentBranch();assert.equal(body.sourceHeadCommitId,'head');assert.equal(body.sourceVersion,4);assert.equal(body.mainVersion,8);assert.equal(body.expectedMainDraftVersion,-1);assert.equal(body.mainLoadedRevision,'main-hash');assert.equal(p.currentBranchId,3);assert.equal(p.snapshot.v,'merged');assert.equal(p.draftVersion,-1)})
await test('Merge cannot authorize another source during confirmation',async()=>{const p=page(),d=defer();p.$confirm=()=>d.promise;let calls=0;wbApi.fastForwardBranchToMain=async()=>{calls++};const job=p.fastForwardCurrentBranch();p.branchVersion++;d.resolve();await job;assert.equal(calls,0)})
await test('Missing main identity and pending source changes send no merge',async()=>{for(const condition of ['main','pending']){const p=page();if(condition==='main')p.branches[0].loadedRevision=null;else p.pendingChanges=[{changeId:'local'}];let calls=0;wbApi.fastForwardBranchToMain=async()=>{calls++};await p.fastForwardCurrentBranch();assert.equal(calls,0);assert(p.errors.length)}})
await test('Merge conflict leaves source snapshot and loaded version',async()=>{const p=page();wbApi.fastForwardBranchToMain=async()=>({code:409,msg:'main changed'});await p.fastForwardCurrentBranch();assert.equal(p.currentBranchId,2);assert.equal(p.snapshot.v,'local');assert.equal(p.loadedRevision,'loaded')})
await test('Late merge reply preserves newer local edits',async()=>{const p=page(),d=defer();wbApi.fastForwardBranchToMain=()=>d.promise;const job=p.fastForwardCurrentBranch();for(let i=0;i<10;i++)await Promise.resolve();p.snapshot={v:'new-edit'};d.resolve({code:200,data:{branchId:3,content:'{"v":"server"}'}});await job;assert.equal(p.snapshot.v,'new-edit');assert.equal(p.currentBranchId,2);assert.equal(p.backups.at(-1).snapshot.v,'new-edit')})
console.log(JSON.stringify({passed,failures,realBrowser:false,realDatabase:false}));process.exitCode=failures?1:0
})().catch(e=>{console.error(e);process.exitCode=1})
