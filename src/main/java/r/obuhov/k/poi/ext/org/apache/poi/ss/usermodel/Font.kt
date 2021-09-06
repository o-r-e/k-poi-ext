package r.obuhov.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.Font


fun Font.clone(source: Font) {
    this.fontName = source.fontName
    this.fontHeight = source.fontHeight
    this.italic = source.italic
    this.strikeout = source.strikeout
    this.color = source.color
    this.typeOffset = source.typeOffset
    this.underline = source.underline
    this.charSet = source.charSet
    this.bold = source.bold
}
