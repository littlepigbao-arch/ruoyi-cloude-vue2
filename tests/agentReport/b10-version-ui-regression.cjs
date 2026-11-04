// Exercise real commit/restore page methods with in-memory API boundaries.
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
const settle=async()=>{for(let i=0;i<20;i++)await Promise.resolve()}
function page(){const c=load(path.join(dir,'index.vue')).default,p=c.data();for(const[k,v]of Object.entries(c.methods))p[k]=v.bind(p);p.errors=[];Object.assign(p,{currentDocumentId:1,currentBranchId:2,branchVersion:4,headCommitId:'head',draftVersion:7,loadedRevision:'loaded',currentDocumentVersion:11,workspaceGeneration:1,pendingChanges:[{source:'manual',changeId:'local'}],pendingActions:[],commitMessage:'save',getSnapshot:()=>({v:'local'}),flushManualChanges(){},openVersionDialog:async()=>{},$confirm:()=>Promise.resolve(),$modal:{msgError:e=>p.errors.push(e),msgWarning:e=>p.errors.push(e),msgSuccess(){}}});p.runPersistTask=(_title,fn)=>{p.job=Promise.resolve().then(()=>fn(()=>{}));return p.job};return p}
async function test(name,fn){try{await fn();passed++;console.log('PASS '+name)}catch(e){failures++;console.error('FAIL '+name+': '+e.stack)}}
;(async()=>{
await test('Real commit method sends the loaded draft binding with its manual snapshot',async()=>{const p=page();let sent;wbApi.createCommit=async(_,body)=>{sent=body;return{code:200,data:{commitId:'new',branchVersion:5,workbookVersion:12,afterHash:'saved'}}};p.commitPendingChanges();await p.job;await settle();assert.equal(sent.expectedDraftVersion,7);assert.equal(sent.loadedRevision,'loaded');assert.equal(sent.baseVersion,4);assert.equal(sent.content,JSON.stringify({v:'local'}));assert.equal(p.loadedRevision,'saved');assert.equal(p.pendingChanges.length,0)})
await test('Stale commit rejection keeps local edits and loaded baseline',async()=>{const p=page();wbApi.createCommit=async()=>({code:409,msg:'draft changed'});p.commitPendingChanges();await p.job.catch(()=>{});await settle();assert.equal(p.pendingChanges[0].changeId,'local');assert.equal(p.loadedRevision,'loaded');assert.equal(p.draftVersion,7);assert(p.errors.some(e=>e.includes('draft changed')))})
await test('Missing loaded version sends no commit and preserves edits',async()=>{const p=page();p.loadedRevision=null;let calls=0;wbApi.createCommit=async()=>{calls++};p.commitPendingChanges();await p.job.catch(()=>{});await settle();assert.equal(calls,0);assert.equal(p.pendingChanges.length,1);assert(p.errors.length)})
await test('Commit reply cannot change a newly selected workspace',async()=>{const p=page(),d=defer();wbApi.createCommit=()=>d.promise;p.commitPendingChanges();await settle();p.currentDocumentId=99;p.workspaceGeneration++;p.loadedRevision='new-workspace';p.pendingChanges=[{changeId:'new-local'}];d.resolve({code:200,data:{commitId:'old',afterHash:'old-saved',branchVersion:5}});await p.job.catch(()=>{});await settle();assert.equal(p.loadedRevision,'new-workspace');assert.equal(p.pendingChanges[0].changeId,'new-local')})
await test('Real restore sends loaded binding and does not reload after conflict',async()=>{const p=page();p.pendingChanges=[];let sent,reads=0;wbApi.restoreVersion=async(_book,_target,body)=>{sent=body;return{code:409,msg:'dirty draft'}};wbApi.getVersionStatus=async()=>{reads++;return{code:200}};p.restoreToVersion({commitId:'old',commitNo:1});await settle();assert.equal(sent.expectedDraftVersion,7);assert.equal(sent.loadedRevision,'loaded');assert.equal(sent.branchId,2);assert.equal(reads,0);assert.equal(p.loadedRevision,'loaded');assert(p.errors.some(e=>e.includes('dirty draft')))})
await test('Restore confirmation cannot authorize another workspace',async()=>{const p=page(),d=defer();p.pendingChanges=[];p.$confirm=()=>d.promise;let calls=0;wbApi.restoreVersion=async()=>{calls++};p.restoreToVersion({commitId:'old',commitNo:1});p.currentDocumentId=99;p.workspaceGeneration++;d.resolve();await settle();assert.equal(calls,0)})
await test('Missing restore baseline sends no destructive request',async()=>{const p=page();p.pendingChanges=[];p.draftVersion=null;let calls=0;wbApi.restoreVersion=async()=>{calls++};p.restoreToVersion({commitId:'old',commitNo:1});await settle();assert.equal(calls,0);assert(p.errors.length)})
console.log(JSON.stringify({passed,failures,realBrowser:false,realDatabase:false}));process.exitCode=failures?1:0
})().catch(e=>{console.error(e);process.exitCode=1})
