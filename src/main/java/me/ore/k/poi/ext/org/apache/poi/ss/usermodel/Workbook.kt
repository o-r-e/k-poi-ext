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

@file:Suppress("unused")

package me.ore.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.Font
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook


// region Font
/**
 * Default workbook font (font with index `0`)
 */
val Workbook.defaultFont: Font
    get() = this.getFontAt(0)

/**
 * Returns default workbook font (font with index `0`)
 *
 * Finally, applies [init] to the default font
 *
 * @param init The block to be applied to the default font
 *
 * @return Default workbook font
 */
inline fun Workbook.defaultFont(init: (font: Font) -> Unit): Font = this.defaultFont.apply(init)

/**
 * Creates a new font for the workbook and copies the settings from [source] into it
 *
 * @param source Font settings source
 *
 * @return New font
 */
fun Workbook.cloneFont(source: Font = this.defaultFont): Font = this.createFont().apply { this.clone(source) }

/**
 * Creates a new font for the workbook and copies the settings from [source] into it
 *
 * Finally, applies [init] to the new font
 *
 * @param source Font settings source
 * @param init The block to be applied to the new font
 *
 * @return New font
 */
inline fun Workbook.cloneFont(source: Font = this.defaultFont, init: (font: Font) -> Unit): Font = this.cloneFont(source).apply(init)
// endregion


// region Cell style
/**
 * Default workbook cell style (cell style with index `0`)
 */
val Workbook.defaultCellStyle: CellStyle
    get() = this.getCellStyleAt(0)

/**
 * Returns default workbook cell style (cell style with index `0`)
 *
 * Finally, applies [init] to the default cell style
 *
 * @param init The block to be applied to the default cell style
 *
 * @return Default workbook cell style
 */
inline fun Workbook.defaultCellStyle(init: (cellStyle: CellStyle) -> Unit): CellStyle = this.defaultCellStyle.apply(init)

/**
 * Creates a new cell style for the workbook and copies the settings from [source] into it
 *
 * @param source Cell style settings source
 *
 * @return New cell style
 */
fun Workbook.cloneCellStyle(source: CellStyle = this.defaultCellStyle): CellStyle = this.createCellStyle().apply { this.cloneStyleFrom(source) }

/**
 * Creates a new cell style for the workbook and copies the settings from [source] into it
 *
 * Finally, applies [init] to the new cell style
 *
 * @param source Cell style settings source
 * @param init The block to be applied to the new cell style
 *
 * @return New cell style
 */
inline fun Workbook.cloneCellStyle(source: CellStyle = this.defaultCellStyle, init: (cellStyle: CellStyle) -> Unit): CellStyle = this.cloneCellStyle(source).apply(init)
// endregion


// region Sheet
/**
 * Creates a sheet in the workbook with the passed [name] and applies [init] to the created sheet
 *
 * @param name New sheet name
 * @param init Block to be applied to the new sheet
 *
 * @return New sheet
 */
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
