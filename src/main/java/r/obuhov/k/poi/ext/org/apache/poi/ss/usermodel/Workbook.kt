package r.obuhov.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook


// region Font
val Workbook.defaultFont: Font
    get() = this.getFontAt(0)

inline fun Workbook.defaultFont(init: (font: Font) -> Unit): Font = this.defaultFont.apply(init)

fun Workbook.cloneFont(source: Font = this.defaultFont): Font = this.createFont().apply { this.clone(source) }

inline fun Workbook.cloneFont(source: Font = this.defaultFont, init: (font: Font) -> Unit): Font = this.cloneFont(source).apply(init)
// endregion


// region Cell style
val Workbook.defaultCellStyle: CellStyle
    get() = this.getCellStyleAt(0)

inline fun Workbook.defaultCellStyle(init: (cellStyle: CellStyle) -> Unit): CellStyle = this.defaultCellStyle.apply(init)

fun Workbook.cloneCellStyle(source: CellStyle = this.defaultCellStyle): CellStyle = this.createCellStyle().apply { this.cloneStyleFrom(source) }

inline fun Workbook.cloneCellStyle(source: CellStyle = this.defaultCellStyle, init: (cellStyle: CellStyle) -> Unit): CellStyle = this.cloneCellStyle(source).apply(init)
// endregion


// region Sheet
inline fun Workbook.createSheet(name: String? = null, init: (sheet: Sheet) -> Unit): Sheet {
    val result = if (name == null) {
        this.createSheet()
    } else {
        this.createSheet(name)
    }
    result.apply(init)
    return result
}
// endregion
