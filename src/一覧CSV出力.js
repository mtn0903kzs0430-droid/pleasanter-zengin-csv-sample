/* 楽天銀行・法人総合振込: 全銀CSV。固定長120バイト形式とは別形式。 */
(function (root) {
    'use strict';
    const columns = ['ResultId', 'SiteId', 'UpdatedTime', 'Title', 'Status',
        'ClassA', 'ClassB', 'ClassC', 'ClassD', 'ClassE', 'ClassF',
        'NumA', 'DateA', 'DateB', 'DescriptionA'];
    // DefaultはAPI利用者のタイムゾーンに変換済みの日付を返す。
    const view = () => ({ ApiDataType: 'Default', Incomplete: false,
        ColumnSorterHash: { ResultId: 'asc' } });
    function normalize(r) {
        const flat = { ...r, ...r.ClassHash, ...r.NumHash, ...r.DateHash, ...r.DescriptionHash };
        flat.Title = r.ItemTitle || r.Title;
        return flat;
    }
    function requireValue(ok, message) { if (!ok) throw new Error(message); }
    function digits(value, length, name, exact = false) {
        const s = String(value ?? '');
        requireValue(/^[0-9]+$/.test(s) && s.length <= length && (!exact || s.length === length),
            name + ': 半角数字' + (exact ? length : '1～' + length) + '桁で指定してください。');
        return s.padStart(length, '0');
    }
    function kana(value, max, name) {
        const s = String(value ?? '');
        requireValue(/^[0-9A-Zｱ-ﾝﾞﾟ()./ -]+$/.test(s) && s.trim().length > 0 && s.length <= max,
            name + ': 使用可能な半角文字で1～' + max + '文字。長音は半角ハイフンを使用してください。');
        return s;
    }
    function day(value) {
        const s = String(value ?? '').slice(0, 10);
        requireValue(/^\d{4}-\d{2}-\d{2}$/.test(s), '振込指定日が不正です。');
        const d = new Date(s + 'T00:00:00Z');
        requireValue(!isNaN(d) && d.toISOString().slice(0, 10) === s && s >= '2000-01-01',
            '振込指定日が未入力または不正です。');
        return s;
    }
    function eligible(r) { return Number(r.Status) === 100; }
    function fingerprint(r) {
        return JSON.stringify(columns.filter(c => c !== 'Title').map(c => r[c] ?? null));
    }
    function build(records, cfg) {
        requireValue(records.length > 0 && records.length <= 999999, '出力対象を1～999999件選択してください。');
        const ids = new Set();
        const date = day(records[0].DateA);
        const rows = [['1', '21', '0', digits(cfg.companyCode, 10, '委託者コード', true),
            kana(cfg.companyName, 40, '委託者名'), date.slice(5, 7) + date.slice(8, 10),
            '0036', '', digits(cfg.branch, 3, '依頼人支店番号', true), '', '1',
            digits(cfg.account, 7, '依頼人口座番号'), '']];
        let total = 0;
        for (const r of records) {
            try {
            const id = digits(r.ResultId, 20, 'レコードID').replace(/^0+/, '');
            requireValue(id !== '' && !ids.has(id), 'レコードIDが0または重複しています。');
            ids.add(id);
            requireValue(eligible(r), 'ID ' + id + ': 未出力ではありません。');
            requireValue(day(r.DateA) === date, '異なる振込指定日を同じファイルに出力できません。');
            requireValue(['1', '2', '4'].includes(String(r.ClassD)), 'ID ' + id + ': 預金種目は1・2・4のみです。');
            const amount = Number(r.NumA);
            requireValue(Number.isSafeInteger(amount) && amount > 0 && amount <= 9999999999,
                'ID ' + id + ': 金額は1～9999999999の整数にしてください。');
            total += amount;
            requireValue(total <= 999999999999, '合計金額が12桁を超えています。');
            rows.push(['2', digits(r.ClassB, 4, '銀行番号', true), '',
                digits(r.ClassC, 3, '支店番号', true), '', '', String(r.ClassD),
                digits(r.ClassE, 7, '口座番号'), kana(r.ClassF, 30, '受取人名'),
                digits(amount, 10, '金額'), '1', id, '', '', '']);
            } catch (e) { throw new Error('ID ' + r.ResultId + ': ' + e.message); }
        }
        rows.push(['8', digits(records.length, 6, '件数'), digits(total, 12, '合計金額'), '']);
        rows.push(['9', '']);
        // 未使用欄を空欄として残すため、行末のカンマも必要。
        return { rows, text: rows.map(r => r.join(',')).join('\r\n') + '\r\n', total, date };
    }
    function bytes(text) {
        // 許可済み半角文字のみ。ASCII + JIS X 0201カナ (Shift-JIS互換)。BOMなし。
        return Uint8Array.from(Array.from(text, c => {
            const n = c.charCodeAt(0);
            if (n <= 0x7f) return n;
            if (n >= 0xff61 && n <= 0xff9f) return n - 0xff61 + 0xa1;
            throw new Error('1バイトに変換できない文字です: ' + c);
        }));
    }
    function timestamp(d = new Date()) {
        const p = n => String(n).padStart(2, '0');
        return d.getFullYear() + '-' + p(d.getMonth()+1) + '-' + p(d.getDate()) + 'T' +
            p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds());
    }
    function filename(date, stamp) {
        return '楽天銀行_総合振込_振込日' + day(date).replace(/-/g, '') + '_作成' + stamp.slice(0, 10).replace(/-/g, '') + '_' +
            stamp.slice(11).replace(/:/g, '') + '.csv';
    }
    function updateData(stamp, name) {
        return { ApiVersion: 1.1, Status: 900,
            DateHash: { DateB: stamp }, DescriptionHash: { DescriptionA: name } };
    }
    root.Zengin = { columns, view, normalize, build, bytes, day, eligible, fingerprint,
        timestamp, filename, updateData, requireValue };
    if (typeof module !== 'undefined') module.exports = root.Zengin;
})(typeof globalThis !== 'undefined' ? globalThis : this);

/* 上のZengin本体と合わせて「一覧」スクリプトに登録。 */

(function () {
    'use strict';
    const CONFIG = { companyCode: '1234567890', companyName: 'ﾃｽﾄ', branch: '201', account: '1234567' };
    const Z = window.Zengin, siteId = Number($p.siteId());
    const storageKey = 'rakuten-zengin-status-pending-v3-' + location.origin + '-' + siteId;
    if (window.RakutenZenginMount) { window.RakutenZenginMount(); return; }
    let busy = false, prepared = null;
    const $open = $('<button id="RakutenZenginOpen" type="button" style="border-radius:0;background:#FF9800;color:#212121;font-weight:bold;padding:8px 16px;border:1px solid #E68900">全銀CSV出力</button>');
    function openDialog() { prepared=null; refreshRecovery(); $dialog[0].showModal(); }
    function mountButton() {
        if (!$('#RakutenZenginOpen').length) {
            // 一覧更新で削除されるとjQueryのイベントも消えるため再設定する。
            $open.off('click').on('click',openDialog);
            $('#MainCommands').append($open);
        }
    }
    window.RakutenZenginMount=mountButton;
    const previousGridLoad=$p.events.on_grid_load;
    $p.events.on_grid_load=function() {
        try { if(typeof previousGridLoad==='function') return previousGridLoad.apply(this,arguments); }
        finally { mountButton(); }
    };
    mountButton();
    const $dialog = $('<dialog style="width:600px;max-width:90vw;box-sizing:border-box;padding:32px;font-size:18px;border:1px solid #999;border-radius:8px"></dialog>').appendTo('body');
    const $date = $('<input type="date" aria-label="振込指定日" style="font-size:20px;padding:10px;margin:16px 12px">');
    const $export = $('<button type="button" style="font-size:18px;padding:10px 20px">CSV出力</button>');
    const $close = $('<button type="button" style="font-size:18px;padding:10px 20px;margin-left:12px">閉じる</button>');
    const $retry = $('<button type="button">保存済みCSVの更新を再開</button>');
    const $cancelPending = $('<button type="button">未保存の処理を取り消す</button>');
    const $confirmSaved = $('<button type="button">前回のCSV保存を確認したので更新する</button>');
    $dialog.append($('<h2>全銀CSV出力</h2>'), '振込指定日', $date, $('<p>指定日の「未出力」をすべて出力します。<br>一覧の絞り込み・行選択は使用しません。</p>'), $export, $close, $('<hr>'), $retry, $confirmSaved, $cancelPending);
    const $confirm = $('<dialog style="width:440px;max-width:90vw;padding:28px;font-size:18px;border-radius:8px"></dialog>').appendTo('body');
    const $summary = $('<p style="white-space:pre-wrap"></p>');
    const $ok = $('<button type="button" style="font-size:18px;padding:10px 24px">OK</button>');
    const $back = $('<button type="button" style="font-size:18px;padding:10px 24px;margin-left:12px">キャンセル</button>');
    $confirm.append($('<h2>出力内容の確認</h2>'), $summary, $ok, $back);
    function lock(value) { busy=value; [$open,$date,$export,$close,$retry,$confirmSaved,$cancelPending,$ok,$back].forEach(e=>e.prop('disabled',value)); }
    function refreshRecovery() { const p=pending(); $retry.toggle(!!p && p.phase==='saved'); $confirmSaved.toggle(!!p && p.phase!=='saved'); $cancelPending.toggle(!!p && p.phase!=='saved'); }
    function error(e, saved=false) { window.alert((saved ? 'CSV保存済み・更新未完了です。' : '処理を中止しました。CSVは保存されていません。')+'\n'+e.message); }
    function pending() { const s = localStorage.getItem(storageKey); return s ? JSON.parse(s) : null; }
    function persist(p) { localStorage.setItem(storageKey, JSON.stringify(p)); }
    function api(method, id, data) {
        return new Promise((resolve, reject) => {
            $p[method]({ id, data, done: result => {
                if (!result || Number(result.StatusCode) !== 200) reject(new Error('ID ' + id + ': APIエラー (' + String(result?.StatusCode) + ')'));
                else resolve(result.Response);
            }, fail: xhr => reject(new Error('ID ' + id + ': 通信エラー (' + xhr.status + ')')) });
        });
    }
    async function getRecord(id) {
        const response = await api('apiGet', id, { ApiVersion: 1.1, View: Z.view() });
        Z.requireValue(response && response.Data?.length === 1, 'ID ' + id + ': 取得できません。');
        const r = Z.normalize(response.Data[0]);
        Z.requireValue(Number(r.SiteId) === siteId && String(r.ResultId) === String(id), '対象サイトまたはIDが一致しません。');
        return r;
    }
    async function getTargets(date) {
        const records = [], seen = new Set();
        let offset = 0, total = null;
        const view = Z.view();
        view.ColumnFilterHash = {Status:'[100]', DateA: JSON.stringify([date + ' 00:00:00,' + date + ' 23:59:59'])};
        for (;;) {
            const response = await api('apiGet', siteId, {ApiVersion:1.1, Offset:offset, View:view});
            Z.requireValue(response && Array.isArray(response.Data) && Number.isInteger(response.TotalCount), '取得結果が不正です。');
            if (total === null) total = response.TotalCount;
            Z.requireValue(total === response.TotalCount, '取得中に対象件数が変わりました。再確認してください。');
            for (const raw of response.Data) {
                const r = Z.normalize(raw), id = String(r.ResultId);
                Z.requireValue(Number(r.SiteId) === siteId && Z.eligible(r) && Z.day(r.DateA) === date, 'ID ' + id + ': 取得条件が一致しません。');
                Z.requireValue(!seen.has(id), '取得中にIDが重複しました。再確認してください。');
                seen.add(id); records.push(r);
            }
            offset += response.Data.length;
            Z.requireValue(offset <= total, '取得件数が一致しません。再確認してください。');
            if (offset === total) break;
            Z.requireValue(response.Data.length > 0, '取得が途中で停止しました。');
        }
        return records.sort((a,b)=>Number(a.ResultId)-Number(b.ResultId));
    }
    async function recheck(records, date) {
        const current = await getTargets(date);
        Z.requireValue(current.length === records.length && current.every((r,i)=>Z.fingerprint(r) === Z.fingerprint(records[i])),
            '対象または内容が変わりました。「CSV出力」からやり直してください。');
    }

    async function complete(p) {
        Z.requireValue(p.phase==='saved','CSVの保存確認が済んでいません。');
        const failed=[];
        for (const item of p.items) {
            if(item.done) continue;
            try {
                const current=await getRecord(item.record.ResultId);
                if (!(Number(current.Status)===900 && current.DescriptionA===p.name && String(current.DateB).slice(0,19)===p.stamp)) {
                    Z.requireValue(Z.fingerprint(current)===Z.fingerprint(item.record),'取得後に変更されています。');
                    await api('apiUpdate',current.ResultId,Z.updateData(p.stamp,p.name));
                    const result=await getRecord(current.ResultId);
                    Z.requireValue(Number(result.Status)===900 && result.DescriptionA===p.name && String(result.DateB).slice(0,19)===p.stamp,'更新後の値を確認できません。');
                }
                item.done=true; item.error='';
            } catch(e) { item.error=e.message; failed.push(item.record.ResultId+': '+e.message); }
            persist(p);
        }
        prepared=null;
        if(failed.length) window.alert('CSV保存済み・更新未完了\n'+p.name+'\n未完了 '+failed.length+'件\n'+failed.join('\n')+'\nCSVを再出力せず「保存済みCSVの更新を再開」を押してください。');
        else { localStorage.removeItem(storageKey); window.alert('処理完了\n'+p.items.length+'件を出力済にしました。\n'+p.name); $dialog[0].close(); location.reload(); }
    }
    $close.on('click',function(){if(!busy)$dialog[0].close();});
    $dialog.on('cancel',function(e){if(busy)e.preventDefault();});
    $confirm.on('cancel',function(e){if(busy)e.preventDefault();else prepared=null;});
    $back.on('click',function(){prepared=null;$confirm[0].close();});
    $export.on('click',async function(){
        if(busy)return;lock(true);
        try {
            Z.requireValue(typeof window.showSaveFilePicker==='function','保存先指定機能が使えるEdge／Chromeと、対応する接続環境で開いてください。');
            Z.requireValue(!pending(),'前回の保存・更新が未完了です。復旧用ボタンで確認してください。');
            const date=Z.day($date.val()), records=await getTargets(date);
            Z.requireValue(records.length>0,'指定日に未出力のレコードはありません。');
            const csv=Z.build(records,CONFIG);prepared={records,csv};
            $summary.text('振込指定日 '+csv.date+'\n件数 '+records.length+'件\n合計金額 '+csv.total.toLocaleString()+'円\n\nCSVを保存しますか？');$confirm[0].showModal();
        }catch(e){error(e);}finally{lock(false);refreshRecovery();}
    });
    $ok.on('click',async function(){
        if(busy||!prepared)return;lock(true);$confirm[0].close();let p;
        try {
            const stamp=Z.timestamp();
            const handle=await window.showSaveFilePicker({suggestedName:Z.filename(prepared.csv.date,stamp),types:[{description:'全銀CSV',accept:{'text/csv':['.csv']}}]});
            await recheck(prepared.records,prepared.csv.date);
            const bytes=Z.bytes(prepared.csv.text);
            p={name:handle.name,stamp,phase:'awaitingSave',items:prepared.records.map(record=>({record,done:false}))};persist(p);
            const stream=await handle.createWritable();
            try {await stream.write(bytes);await stream.close();}catch(e){try{await stream.abort();}catch(_){}throw e;}
            p.phase='saved';persist(p);await complete(p);
        }catch(e){if(p?.phase==='awaitingSave')localStorage.removeItem(storageKey);if(e.name==='AbortError')window.alert('保存を中止しました。出力済には変更していません。');else error(e,p?.phase==='saved');}
        finally{lock(false);refreshRecovery();}
    });
    $retry.on('click',async function(){if(busy)return;lock(true);try{const p=pending();Z.requireValue(p&&p.phase==='saved','保存済みの更新情報がありません。');await complete(p);}catch(e){error(e,true);}finally{lock(false);refreshRecovery();}});
    $confirmSaved.on('click',async function(){if(busy)return;if(!window.confirm('前回のCSVが正しく保存されていることを確認しましたか？'))return;lock(true);try{const p=pending();Z.requireValue(p&&p.phase!=='saved','未保存の処理情報がありません。');p.phase='saved';persist(p);await complete(p);}catch(e){error(e,true);}finally{lock(false);refreshRecovery();}});
    $cancelPending.on('click',function(){const p=pending();if(p&&p.phase!=='saved'&&window.confirm('前回のCSVが保存されていないことを確認しましたか？')){localStorage.removeItem(storageKey);prepared=null;refreshRecovery();}});
})();
