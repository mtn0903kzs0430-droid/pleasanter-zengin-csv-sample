const fs=require('node:fs'),vm=require('node:vm'),assert=require('node:assert/strict'),path=require('node:path');
let harness=fs.readFileSync(path.join(__dirname,'mock-browser.txt'),'utf8');

harness=harness.replace("if(s==='#RakutenZenginPanel')","if(s==='#RakutenZenginOpen')");
harness=harness.replace("origin:'http://localhost'","origin:'http://localhost',reload(){alerts.push('reload');}");
harness=harness.replace('const elements=[],store=', 'const alerts=[];const elements=[],store=');
harness=harness.replace('confirm:()=>true,','alert:t=>alerts.push(t),confirm:()=>true,');
harness=harness.replace('on(n,f){','off(n){delete this.handlers[n];return this;}on(n,f){');
harness=harness.replace('$p:{siteId:', '$p:{events:{on_grid_load(){}},siteId:');
harness=harness.replace("const panel=new E('main');", "let present=false, mountCount=0, previousCount=0;const panel=new E('main');panel.append=function(e){if(e.html.includes('RakutenZenginOpen')){present=true;mountCount++;}return this;};");
harness=harness.replace("if(s==='#RakutenZenginOpen')return{length:0};", "if(s==='#RakutenZenginOpen')return{length:present?1:0};if(s==='#MainCommands')return panel;");
harness=harness.replace('on_grid_load(){}','on_grid_load(){previousCount++;}');
harness=harness.replace('name:opts.suggestedName','name:options.rename?"変更した名前.csv":opts.suggestedName');
harness=harness.replace('return{rows,store,updates,writes,click,message:','return{rows,store,updates,writes,click,alerts,message:');
harness=harness.replace('click,alerts,message:', 'click,alerts,gridRefresh:()=>{present=false;elements.find(e=>e.html.includes(\'RakutenZenginOpen\')).handlers={};ctx.$p.events.on_grid_load();},repeatEvent:()=>ctx.$p.events.on_grid_load(),rerun:()=>vm.runInNewContext(source,ctx),counts:()=>({mountCount,previousCount,dialogs:elements.filter(e=>e.html.startsWith(\'<dialog\')).length}),message:');
const {records,dummy}=JSON.parse(fs.readFileSync(path.join(__dirname,'fixture.json')));
const source=fs.readFileSync(path.join(__dirname,'../src/一覧CSV出力.js'),'utf8').replace(/const CONFIG = [\s\S]*?;/,'const CONFIG = '+JSON.stringify(dummy)+';');
const box={require,console,Buffer,records,dummy,source,fs,assert,vm,Blob,structuredClone};vm.createContext(box);vm.runInContext(harness+';globalThis.setup=setup;',box);
(async()=>{
 let lifecycle=box.setup();assert.equal(lifecycle.counts().mountCount,1);lifecycle.gridRefresh();assert.equal(lifecycle.counts().mountCount,2);await lifecycle.click('RakutenZenginOpen');lifecycle.repeatEvent();assert.equal(lifecycle.counts().mountCount,2);lifecycle.rerun();assert.equal(lifecycle.counts().dialogs,2);assert.equal(lifecycle.counts().previousCount,2);console.log('PASS grid refresh remount / click handler restored / no duplicate dialogs or buttons / previous event preserved');
 let a=box.setup({rename:true});await a.click('>CSV出力<');assert.equal(a.writes.length,0);await a.click('>OK<');assert.equal(a.updates.length,3);assert.deepEqual(a.writes[0],fs.readFileSync(path.join(__dirname,'expected-sample.bin')));assert.ok(a.rows.every(r=>r.DescriptionA==='変更した名前.csv'));assert.ok(a.alerts.some(t=>t.includes('処理完了')));console.log('PASS rename / exact CSV bytes / update filename / completion');
 for(const opt of [{cancel:true},{writeFail:true}]){a=box.setup(opt);await a.click('>CSV出力<');await a.click('>OK<');assert.equal(a.updates.length,0);assert.equal(a.store.size,0);assert.ok(a.alerts.length);}console.log('PASS cancelled and failed save do not update');
 a=box.setup();a.rows[0].NumHash.NumA=0;await a.click('>CSV出力<');assert.ok(a.alerts.some(t=>t.includes('金額')));assert.equal(a.writes.length,0);console.log('PASS invalid amount popup');
 a=box.setup({partial:true});await a.click('>CSV出力<');await a.click('>OK<');assert.equal(a.store.size,1);assert.ok(a.alerts.some(t=>t.includes('更新未完了')));await a.click('保存済みCSVの更新を再開');assert.equal(a.updates.length,3);assert.equal(a.writes.length,1);assert.equal(a.store.size,0);console.log('PASS retry without duplicate CSV');
 a=box.setup();await a.click('>CSV出力<');a.rows[0].NumHash.NumA++;await a.click('>OK<');assert.equal(a.writes.length,0);assert.equal(a.updates.length,0);console.log('PASS changed target blocks save');
 a=box.setup({fallback:true});await a.click('>CSV出力<');assert.equal(a.writes.length,0);assert.ok(a.alerts.some(t=>t.includes('保存先指定機能')));console.log('PASS unsupported environment popup');
 a=box.setup();await a.click('>CSV出力<');await a.click('キャンセル');assert.equal(a.writes.length,0);console.log('PASS confirmation cancel');
})().catch(e=>{console.error(e);process.exitCode=1;});
