const fs = require('node:fs'), path = require('node:path'), assert = require('node:assert/strict'), vm = require('node:vm');
const root = path.join(__dirname, '..');
const read = p => fs.readFileSync(path.join(root,p),'utf8').replace(/^\uFEFF/,'');
const p = JSON.parse(read('site-package/全銀CSV出力_v0.1.0.json'));
assert.equal(p.Sites.length,3); assert.deepEqual(p.Data,[]); assert.deepEqual(p.Permissions,[]);
for (const values of Object.values(p.PermissionIdList)) assert.deepEqual(values,[]);
for (const key of ['IncludeSitePermission','IncludeRecordPermission','IncludeColumnPermission','IncludeNotifications','IncludeReminders']) assert.equal(p.HeaderInfo[key],false);
const ids = new Set(p.Sites.map(s=>s.SiteId)); assert.equal(ids.size,3);
for(const c of p.HeaderInfo.Convertors) { assert.ok(ids.has(c.SiteId)); if(c.Order) for(const id of JSON.parse(c.Order)) assert.ok(ids.has(id)); }
const payment = p.Sites.find(s=>s.Title==='支払管理'), master=p.Sites.find(s=>s.Title==='マスタ支払先');
assert.equal(payment.SiteSettings.Links[0].SiteId,master.SiteId);
const linkcol=payment.SiteSettings.Columns.find(c=>c.ColumnName==='ClassA');
assert.equal(JSON.parse(linkcol.ChoicesText)[0].SiteId,master.SiteId);
assert.equal(payment.SiteSettings.Scripts[0].Body.trim(),read('src/一覧CSV出力.js').trim());
new vm.Script(read('src/一覧CSV出力.js'));
for(const s of payment.SiteSettings.ServerScripts) {const name='src/server/'+String(s.Id).padStart(2,'0')+'.js';assert.equal(s.Body.trim(),read(name).trim());new vm.Script(read(name));}
for(const c of master.SiteSettings.Columns) for(const key of ['ClientRegexValidation','ServerRegexValidation']) if(c[key]) new RegExp(c[key]);
assert.equal(payment.SiteSettings.Columns.find(c=>c.ColumnName==='ClassI').EditorReadOnly,true);
console.log('PASS package structure / empty data and permissions / links and order / source consistency / JS syntax / regex syntax');
