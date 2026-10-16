// サーバスクリプト：更新前。画面・APIの双方で検査する。
(function () {
    const status = Number(model.Status);
    const previous = Number(saved.Status);
    if (status !== 100 && status !== 900) {
        context.Error('状況は「未出力」「出力済」のいずれかにしてください。');
        return;
    }
    let fields = ['Body','ClassA','ClassB','ClassC','ClassD','ClassE','ClassF','ClassH','ClassI','ClassJ','NumA','DateA'];
    // Titleは既存の計算式から生成されるため比較対象外。
    if (previous === 900) fields = fields.concat(['DateB','DescriptionA']);
    if (previous === 900 || status === 900) {
        const changed = fields.filter(function (name) { return String(model[name]) !== String(saved[name]); });
        if (changed.length) {
            context.Error('出力済の内容は変更できません。先に状況だけを「未出力」に戻して保存し、その後に編集してください。変更項目: ' + changed.join(', '));
        }
    }
})();
