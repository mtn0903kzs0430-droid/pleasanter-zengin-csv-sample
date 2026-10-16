// サーバスクリプト：画面表示の前、行表示の前
(function () {
    const fields = ['Title','Body','ClassA','ClassB','ClassC','ClassD','ClassE','ClassF','ClassH','ClassI','ClassJ','NumA','DateA'];
    if (Number(model.Status) === 900) {
        fields.forEach(function (name) { if (columns[name]) columns[name].ReadOnly = true; });
    }
    // 状況を未出力に戻す操作は残す。出力情報はスクリプトが記録する。
    columns.Status.ReadOnly = false;
    columns.DateB.ReadOnly = true;
    columns.DescriptionA.ReadOnly = true;
})();
