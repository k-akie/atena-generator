/*
 * 開発用コード
 */

// スライド
function _部品の情報を出力する() {
    const slides = SlidesApp.openById(slideId);
    const slide = slides.getSlides()[0]; // 0番から
    const shapes = slide.getShapes();
    for (const shape of shapes) {
        const text = shape.getText().asString().replace(/\r?\n/g, "");
        const fontSize = shape.getText().getTextStyle().getFontSize();
        const fontFamily = shape.getText().getTextStyle().getFontFamily();
        const fontWeight = shape.getText().getTextStyle().getFontWeight();
        const baselineOffset = shape.getText().getTextStyle().getBaselineOffset();
        const top = shape.getTop();
        const left = shape.getLeft();
        const height = shape.getHeight();
        const width = shape.getWidth();
        const bottom = top + height;
        const fillColor = shape.getFill().getSolidFill().getColor().asRgbColor().asHexString();
        const fillAlpha = shape.getFill().getSolidFill().getAlpha();
        const alignment = shape.getContentAlignment().toString();

        // console.log(text, left, top, width, height, bottom);
        // console.log(text, fillColor, fillAlpha, alignment);
        // console.log(text, fontSize, fontFamily, fontWeight, baselineOffset);
        console.log(`{name: '${text}', left: ${left}, top: ${top}, width: ${width}, height: ${height}, fontSize: ${fontSize}},`)
    }
}

// スプシ
function _セル情報を出力する() {
    const spreadSheet = SpreadsheetApp.openById(spreadSheetId);
    const sheet = spreadSheet.getSheets()[0];

    let widths = [];
    for (let i = 1; i <= 16; i++) {
        const value = sheet.getRange(2, i).getValue();
        const width = sheet.getColumnWidth(i);
        widths[i - 1] = width;
    }
    console.log(widths);
}
