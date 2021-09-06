/*
 *    Copyright 2021 Obuhov R.
 *
 *    Licensed under the Apache License, Version 2.0 (the "License");
 *    you may not use this file except in compliance with the License.
 *    You may obtain a copy of the License at
 *
 *        http://www.apache.org/licenses/LICENSE-2.0
 *
 *    Unless required by applicable law or agreed to in writing, software
 *    distributed under the License is distributed on an "AS IS" BASIS,
 *    WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *    See the License for the specific language governing permissions and
 *    limitations under the License.
 */

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
