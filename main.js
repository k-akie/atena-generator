/* **************************************
 * 設定
 * **************************************/
const KEY_SLIDE_ID = 'slideId';
const KEY_SHEET_ID = 'spreadSheetId';
const slideId = PropertiesService.getScriptProperties().getProperty(KEY_SLIDE_ID);
const spreadSheetId = PropertiesService.getScriptProperties().getProperty(KEY_SHEET_ID);

const spreadSheetName = '年賀状_住所録';
const sheetName = '宛名';

/* **************************************
 * メニュー
 * **************************************/
function onOpen() {
    SlidesApp.getUi()
        .createMenu('宛名作成')
        .addItem('宛名スライド作成', '宛名スライド作成')
        .addItem('印字範囲: 表示', '印字範囲の表示')
        .addItem('印字範囲: 非表示', '印字範囲の非表示')
        .addItem('住所データリンクを表示', '住所データリンクを表示')
        .addSubMenu(
            SlidesApp.getUi().createMenu('初期設定')
                .addItem('テンプレスライド追加', 'テンプレスライド追加')
                .addItem('スライドクリア', 'スライドクリア')
                .addItem('住所録スプシ作成', '住所録スプシ作成')
        )
        .addToUi();
}

function 宛名スライド作成() {
    // 紐づいたスライドIDをプロパティに設定
    const presentation = SlidesApp.getActivePresentation();
    PropertiesService.getScriptProperties().setProperty(KEY_SLIDE_ID, presentation.getId());

    const slideId = PropertiesService.getScriptProperties().getProperty(KEY_SLIDE_ID);
    _スライドをリセットする(slideId);

    const destinations = _宛名住所の取得(spreadSheetId);
    if (destinations.length == 0) {
        _アラートDialog('宛名スライド作成', '宛名が0件です。');
        return;
    }

    _宛名スライドを生成する(slideId, destinations);
    _作成完了Dialog();
}

function 印字範囲の表示() {
    _印字範囲(true);
}

function 印字範囲の非表示() {
    _印字範囲(false);
}

function 住所録スプシ作成() {
    if (PropertiesService.getScriptProperties().getProperty(KEY_SHEET_ID)) {
        const url = SpreadsheetApp.openById(spreadSheetId).getUrl();
        _住所録リンクDialog('住所録スプシ作成', '<p>住所録スプシは作成済みです</p>', url);
        return;
    }

    const spreadSheet = _住所録スプシを作成する();
    PropertiesService.getScriptProperties().setProperty(KEY_SHEET_ID, spreadSheet.getId());

    _住所録リンクDialog('住所録スプシ作成', '<p>住所録スプシを作成しました</p>', spreadSheet.getUrl());
}

function 住所データリンクを表示() {
    const url = SpreadsheetApp.openById(spreadSheetId).getUrl();
    _住所録リンクDialog('住所データ', '', url);
}

function テンプレスライド追加() {
    _テンプレ生成();
    _テンプレ作成完了Dialog();
}

function スライドクリア() {
    _スライドをリセットする(slideId);
}

/* **************************************
 * ダイアログ表示
 * **************************************/
function _アラートDialog(title, message) {
    const ui = SlidesApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
}

function _住所録リンクDialog(title, message, url) {
    var html = HtmlService
        .createHtmlOutput(`
      ${message}
      <p>次のリンクから開いてください</p>
      <a href="${url}" target="_blank">住所データのスプレッドシート</a>
    `)
        .setWidth(280)
        .setHeight(180)
    SlidesApp.getUi().showModelessDialog(html, title);
}

function _作成完了Dialog() {
    var html = HtmlService
        .createHtmlOutput(`
      <p>宛名スライドの作成が完了しました</p>
      <ol>
        <li>メニュー「ファイル > 印刷プレビュー」を開く</li>
        <li>「スキップしたスライドを含める」の選択を解除</li>
        <li>印刷崩れがないかを確認し、問題があればプレビューを閉じて微調整してください</li>
        <li>調整が終わったら「PDF形式でダウンロード」します</li>
      </ol>
    `)
        .setWidth(450)
        .setHeight(220)
    SlidesApp.getUi().showModelessDialog(html, '宛名スライド作成');
}

function _テンプレ作成完了Dialog() {
    var html = HtmlService
        .createHtmlOutput(`
      <p>テンプレスライドを先頭に追加しました</p>
      <p>2枚目以降は生成時に削除されます</p>
    `)
        .setWidth(450)
        .setHeight(220)
    SlidesApp.getUi().showModelessDialog(html, 'テンプレスライド作成');
}

/* **************************************
 * 実処理
 * **************************************/
function _スライドをリセットする(slideId) {
    const slides = SlidesApp.openById(slideId);

    // 1枚目(テンプレ)以外を削除
    while (slides.getSlides().length > 1) {
        slides.getSlides()[1].remove();
    }

    slides.saveAndClose();
}

function _宛名スライドを生成する(slideId, destinations) {
    const slides = SlidesApp.openById(slideId);

    // 1枚目(テンプレ)の設定
    const templateSlide = slides.getSlides()[0];
    templateSlide.setSkipped(false);

    for (const destination of destinations) {
        const slide = slides.appendSlide(templateSlide);

        if (slide.getImages().length > 0) {
            slide.getImages()[0].remove();
        }

        const pitch = 16;
        const shapes = slide.getShapes();
        for (const shape of shapes) {
            // 末尾に改行がついちゃうので除外しておく
            const text = shape.getText().asString().replace(/\r?\n/g, "");
            // 背景を無色透明に
            shape.getFill().setTransparent();

            // 宛名の位置微調整
            const isShortToGivenName = (destination.givenName1.length < 3) && (destination.givenName2.length < 3);
            const isShortToFamilyName = (destination.familyName.length < 3);
            const isSingleTo = destination.givenName2 == '';
            // 差出人の位置微調整
            const isShortFromFamilyName = (destination.fromFamilyName.length < 3);
            const isSingleFrom = destination.fromName2 == '';

            switch (text) {
                // /////////////////////////////
                // 宛名
                // /////////////////////////////
                case "宛名郵便番号":
                    shape.getText().setText(destination.postNumber);
                    break;
                case "宛名住所１":
                    shape.getText().setText(destination.address1);
                    break;
                case "宛名住所２":
                    shape.getText().setText(destination.address2);
                    break;
                case "宛名苗字":
                    shape.getText().setText(destination.familyName);
                    if (isShortToFamilyName) shape.setTop(shape.getTop() + pitch);
                    if (isSingleTo) shape.setLeft(shape.getLeft() - pitch);
                    break;
                case "宛名名１":
                    shape.getText().setText(destination.givenName1);
                    if (isSingleTo) shape.setLeft(shape.getLeft() - pitch);
                    break;
                case "敬称１":
                    shape.getText().setText(destination.keisho1);
                    if (isShortToGivenName) shape.setTop(shape.getTop() - pitch * 2);
                    if (isSingleTo) shape.setLeft(shape.getLeft() - pitch);
                    break;
                case "宛名名２":
                    shape.getText().setText(destination.givenName2);
                    if (isSingleTo) shape.remove();
                    break;
                case "敬称２":
                    shape.getText().setText(destination.keisho2);
                    if (isShortToGivenName) shape.setTop(shape.getTop() - pitch * 2);
                    if (isSingleTo) shape.remove();
                    break;

                // /////////////////////////////
                // 差出人
                // /////////////////////////////
                case "差出人郵便番号":
                    shape.getText().setText(destination.fromPostNumber);
                    break;
                case "差出人住所１":
                    shape.getText().setText(destination.fromAddress1);
                    break;
                case "差出人住所２":
                    shape.getText().setText(destination.fromAddress2);
                    break;
                case "差出人苗字":
                    shape.getText().setText(destination.fromFamilyName);
                    if (isShortFromFamilyName) shape.setTop(shape.getTop() + pitch)
                    break;
                case "差出人名１":
                    shape.getText().setText(destination.fromName1);
                    break;
                case "差出人名２":
                    shape.getText().setText(destination.fromName2);
                    if (isSingleFrom) shape.remove();
                    break;
                default:
            }
        }
    }

    // テンプレスライドはスキップに指定する
    slides.getSlides()[0].setSkipped(true);

    slides.saveAndClose();
}

function _宛名住所の取得(spreadSheetId) {
    const spread = SpreadsheetApp.openById(spreadSheetId);
    const sheet = spread.getSheetByName(sheetName);
    // タイトル行を除いたデータ範囲
    const range = sheet.getRange(3, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = range.getValues();

    var dataList = [];
    for (const row of values) {
        var i = 0;
        const data = {
            isSend: row[i++],
            // 宛名
            postNumber: _郵便番号を印刷用表記にする(row[i++]),
            address1: _住所を縦書き用表記にする(row[i++]),
            address2: _住所を縦書き用表記にする(row[i++]),
            familyName: row[i++],
            givenName1: row[i++],
            keisho1: row[i++],
            givenName2: row[i++],
            keisho2: row[i++],
            // 差出人
            fromPostNumber: _郵便番号を印刷用表記にする(row[i++]),
            fromAddress1: _住所を縦書き用表記にする(row[i++]),
            fromAddress2: _住所を縦書き用表記にする(row[i++]),
            fromFamilyName: row[i++],
            fromName1: row[i++],
            fromName2: row[i++],
        }

        if (data.address1 == '') {
            // 住所１がなければデータなしとみなす
            continue;
        }

        if (data.isSend) {
            dataList.push(data);
        }
    }

    return dataList;
}

function _住所を縦書き用表記にする(str) {
    const dic = {
        0: '〇',
        1: '一',
        2: '二',
        3: '三',
        4: '四',
        5: '五',
        6: '六',
        7: '七',
        8: '八',
        9: '九',
        '-': '・',
        '－': '・',
        '―': '・',
    };

    // TODO 全角18文字以内じゃなかったら警告出したい

    return [...str].map((c) => dic[c] || c).join('');
}

function _郵便番号を印刷用表記にする(postNumber) {
    // 半角に揃える
    const halfWidth = `${postNumber}`.replace(/[０-９]/g, function (s) {
        return String.fromCharCode(s.charCodeAt(0) - 0xFEE0);
    });

    // ハイフンを消す
    const noHyphen = halfWidth.replace('-', '');

    // 均等割り付けができないので、半角スペース2個で調整する
    return noHyphen.split('').join('  ');
}

function _住所録スプシを作成する() {
    const spreadSheet = SpreadsheetApp.create(spreadSheetName);

    // シート
    const sheet = spreadSheet.getSheets()[0];
    sheet.setName(sheetName);

    // /////////////////////////////
    // タイトル設定
    // /////////////////////////////
    // 1行目
    sheet.getRange(1, 2).setValue('宛名');
    sheet.getRange(1, 10).setValue('差出人');

    // 2行目
    const toTitles = ['郵便番号', '住所１', '住所２', '苗字', '名前１', '敬称１', '名前２', '敬称２',];
    const fromTitles = ['郵便番号', '住所１', '住所２', '苗字', '名前１', '名前２',];
    const addressTitles = ['送る', ...toTitles, ...fromTitles, 'メモ',];
    for (const i in addressTitles) {
        sheet.getRange(2, Number(i) + 1).setValue(addressTitles[i]);
    }

    // /////////////////////////////
    // 本文設定
    // /////////////////////////////
    // 3～32行目(30件分)
    const dataSize = 30;
    sheet.getRange(3, 1, dataSize).insertCheckboxes();

    // /////////////////////////////
    // 見た目設定
    // /////////////////////////////
    const backgroundColor = '#d9d9d9';

    // タイトル
    sheet.getRange(1, 1, 2, addressTitles.length)
        .setBackground(backgroundColor)
        .setFontWeight('bold')
        .setBorder(false, false, true, false, false, false); // 下線

    // 縦線
    sheet.getRange(1, 2, 2 + dataSize, toTitles.length)
        .setBorder(false, true, false, true, null, null);
    sheet.getRange(1, 2 + toTitles.length, 2 + dataSize, fromTitles.length)
        .setBorder(false, true, false, true, null, null);

    // 横幅
    const widths = [60, 60, 234, 215, 60, 60, 47, 60, 47, 60, 247, 47, 34, 47, 47, 100,];
    for (const i in widths) {
        sheet.setColumnWidth(Number(i) + 1, widths[i]);
    }

    // 本文
    const formulaRange = sheet.getRange(4, 1 + toTitles.length + 1, dataSize - 1, 4);
    formulaRange.setFormula('=if($N4&$O4="","",J$3)');
    formulaRange.setBackground(backgroundColor);

    return spreadSheet;
}

function _テンプレ生成() {
    const shapePositions = [
        { name: '宛名郵便番号', left: 118.95157172736221, top: 28.893308624507874, width: 151.51180620268573, height: 22.062992486428087, fontSize: 16 },
        { name: '宛名住所１', left: 240.50787401574803, top: 61.72244094488189, width: 14.456692913385826, height: 286.5590551181102, fontSize: 13 },
        { name: '宛名住所２', left: 221.9209226377953, top: 147.83267598425198, width: 14.456692913385826, height: 200.45669291338584, fontSize: 13 },
        { name: '宛名苗字', left: 158.10466791338584, top: 73.50984251968504, width: 28.724409448818896, height: 66.89763779527559, fontSize: 27 },
        { name: '宛名名１', left: 158.10466791338584, top: 177.73031496062993, width: 28.724409448818896, height: 99.63779527559055, fontSize: 27 },
        { name: '宛名名２', left: 123.81136082677166, top: 177.73031496062993, width: 28.724409448818896, height: 99.63779527559055, fontSize: 27 },
        { name: '敬称１', left: 158.10466791338584, top: 276.253937007874, width: 28.724409448818896, height: 99.63779527559055, fontSize: 27 },
        { name: '敬称２', left: 123.81136082677166, top: 276.253937007874, width: 28.724409448818896, height: 99.63779527559055, fontSize: 27 },
        { name: '差出人郵便番号', left: 10.588890255905511, top: 342.95472440944883, width: 92.57480384796624, height: 12.354330401720963, fontSize: 9 },
        { name: '差出人住所１', left: 58.840389763779534, top: 156.32874015748033, width: 10.464566929133865, height: 175.98425196850394, fontSize: 8 },
        { name: '差出人住所２', left: 46.20062598425197, top: 168.32874015748033, width: 10.464566929133857, height: 175.98425196850394, fontSize: 8 },
        { name: '差出人苗字', left: 25.137956692913388, top: 214.8464566929134, width: 17.92913385826772, height: 37.818897637795274, fontSize: 15 },
        { name: '差出人名１', left: 25.137956692913388, top: 273.34645669291336, width: 17.92913385826772, height: 92.36220472440945, fontSize: 15 },
        { name: '差出人名２', left: 7.208824803149606, top: 273.34645669291336, width: 17.929133858267715, height: 92.36220472440945, fontSize: 15 },
    ];

    // FIXME ページ設定でスライドサイズを変える

    const slides = SlidesApp.openById(slideId);
    const slide = slides.appendSlide();
    slide.setSkipped(true);
    slide.move(0);

    for (const position of shapePositions) {
        const shape = slide.insertShape(
            SlidesApp.ShapeType.TEXT_BOX,
            position.left, position.top, position.width, position.height
        );
        shape.getText().setText(position.name);
        shape.getText().getTextStyle().setFontFamilyAndWeight('Klee One', 700)
        shape.getText().getTextStyle().setFontSize(position.fontSize);

        if (position.name.includes('住所２')) {
            shape.setContentAlignment(SlidesApp.ContentAlignment.BOTTOM);
        } else {
            shape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        }
        shape.getFill().setSolidFill('#CCCCCC', 0.50);
    }
}

function _印字範囲(isShow) {
    const slides = SlidesApp.openById(slideId);

    if (!isShow) {
        for (const slide of slides.getSlides()) {
            const shapes = slide.getShapes();
            for (const shape of shapes) {
                // 末尾に改行がついちゃうので除外しておく
                const text = shape.getText().asString().replace(/\r?\n/g, "");
                if (text == '印字範囲(郵便番号以外)') {
                    shape.remove();
                }
            }
        }

        return;
    }

    const position = { name: '印字範囲\r\n(郵便番号以外)', left: 10.208661417322833, top: 61.72244094488189, width: 260.70078740157481, height: 282.1653543307087, fontSize: 10 };
    for (const slide of slides.getSlides()) {
        const shape = slide.insertShape(
            SlidesApp.ShapeType.TEXT_BOX,
            position.left, position.top, position.width, position.height
        );
        shape.getText().setText(position.name);
        shape.getText().getTextStyle().setFontSize(position.fontSize);
        shape.getText().getTextStyle().setForegroundColor('#999999');
        shape.setContentAlignment(SlidesApp.ContentAlignment.TOP);
        shape.getFill().setSolidFill('#FFF2CC');

        shape.sendToBack(); // 最背面
    }
}






