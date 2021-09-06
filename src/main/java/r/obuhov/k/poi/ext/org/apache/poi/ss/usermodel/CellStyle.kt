@file:Suppress("unused")

package r.obuhov.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.IndexedColors


// region Border
fun CellStyle.border(
    style: BorderStyle = BorderStyle.THIN,
    color: IndexedColors = IndexedColors.AUTOMATIC,
    left: Boolean = true,
    top: Boolean = true,
    right: Boolean = true,
    bottom: Boolean = true
) {
    if (left) {
        this.borderLeft = style
        this.leftBorderColor = color.index
    }
    if (top) {
        this.borderTop = style
        this.topBorderColor = color.index
    }
    if (right) {
        this.borderRight = style
        this.rightBorderColor = color.index
    }
    if (bottom) {
        this.borderBottom = style
        this.bottomBorderColor = color.index
    }
}

fun CellStyle.borderLeft(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, top = false, right = false, bottom = false)
}

fun CellStyle.borderTop(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, right = false, bottom = false)
}

fun CellStyle.borderRight(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, top = false, bottom = false)
}

fun CellStyle.borderBottom(style: BorderStyle = BorderStyle.THIN, color: IndexedColors = IndexedColors.AUTOMATIC) {
    this.border(style, color, left = false, top = false, right = false)
}
// endregion
