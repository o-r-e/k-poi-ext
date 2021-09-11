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

import org.apache.poi.ss.util.CellAddress
import org.apache.poi.util.Units
import me.ore.k.poi.ext.CellInit
import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellReference
import java.util.*
import kotlin.math.roundToInt


// region Rows
/**
 * Searches for a row with the specified title. Similar to [getRow][Sheet.getRow] (which is used internally), but first row is "1"
 *
 * @param rowTitle Title of the required row
 *
 * @return Found row or `null`
 *
 * @throws NumberFormatException If [rowTitle] is not an integer
 */
fun Sheet.getRow(rowTitle: String): Row? = this.getRow(rowTitle.toInt() - 1)

/**
 * Searches for a row with the specified index
 *
 * If a row is found, apply [init] to it
 *
 * @param rowIndex The index of the desired row
 * @param init Block to be applied to the found row
 *
 * @return Found row or `null`
 */
inline fun Sheet.getRow(rowIndex: Int, init: (row: Row) -> Unit): Row? = this.getRow(rowIndex)?.apply(init)

/**
 * Creates a row with the specified title. Similar to [createRow][Sheet.createRow] (which is used internally), but the first row is "1"
 *
 * @param rowTitle Title of the required row
 *
 * @return Created row
 *
 * @throws NumberFormatException If [rowTitle] is not an integer
 */
fun Sheet.createRow(rowTitle: String): Row = this.createRow(rowTitle.toInt() - 1)

/**
 * Searches for or creates (if not found) the row at the specified index; index of first row is `0`
 *
 * @param rowIndex The index of the desired row
 *
 * @return Found or created row
 */
fun Sheet.row(rowIndex: Int): Row = this.getRow(rowIndex) ?: this.createRow(rowIndex)

/**
 * Searches for or creates (if not found) the row with the specified title; same as [row][Sheet.row], but title of first row is "1"
 *
 * @param rowTitle The title of the desired row; title of first row is "1"
 *
 * @return Found or created row
 *
 * @throws NumberFormatException If [rowTitle] is not an integer
 */
fun Sheet.row(rowTitle: String): Row = this.row(rowTitle.toInt() - 1)
// endregion


// region Columns
/**
 * Set the width (in units of 1/256th of a character width) of the column
 *
 * Full description - [Sheet.setColumnWidth][Sheet.setColumnWidth]
 *
 * @param columnTitle Title of column ("A", "B", etc.)
 * @param width Column width
 */
fun Sheet.setColumnWidth(columnTitle: String, width: Int) {
    val columnIndex = CellReference.convertColStringToIndex(columnTitle)
    this.setColumnWidth(columnIndex, width)
}

/**
 * Set the width (in pixels) of the column
 *
 * The column width will be set approximately, i.e. may be slightly different from the passed value [width]
 *
 * @param columnIndex Column index; first column index - `0`
 * @param width Column width in pixels
 */
fun Sheet.setColumnWidthInPixels(columnIndex: Int, width: Int) {
    val widthInUnits = ((Units.pixelToEMU(width).toDouble() / Units.EMU_PER_CHARACTER) * 256).roundToInt()
    this.setColumnWidth(columnIndex, widthInUnits)
}

/**
 * Set the width (in pixels) of the column
 *
 * The column width will be set approximately, i.e. may be slightly different from the passed value [width]
 *
 * @param columnTitle Column title ("A", "B", etc.)
 * @param width Column width in pixels
 */
fun Sheet.setColumnWidthInPixels(columnTitle: String, width: Int) {
    val columnIndex = CellReference.convertColStringToIndex(columnTitle)
    this.setColumnWidthInPixels(columnIndex, width)
}

/**
 * Sets the default cell style for the column
 *
 * @param columnTitle Column title ("A", "B", etc.)
 * @param style Default cell style
 */
fun Sheet.setDefaultColumnStyle(columnTitle: String, style: CellStyle) {
    val columnIndex = CellReference.convertColStringToIndex(columnTitle)
    this.setDefaultColumnStyle(columnIndex, style)
}
// endregion


// region Get or create empty cells
/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int): Cell = this.row(rowIndex).cell(columnIndex)

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex).apply(init)


/**
 * Finds or creates a cell that matches [address]
 *
 * If there is no row corresponding to [address], then creates such a row
 *
 * @param address Cell address
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress): Cell = this.cell(address.row, address.column)

/**
 * Finds or creates a cell that matches [address]
 *
 * If there is no row corresponding to [address], then creates such a row
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, init: (cell: Cell) -> Unit): Cell = this.cell(address).apply(init)


/**
 * Finds or creates a cell that matches [address]
 *
 * If there is no row corresponding to [address], then creates such a row
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String): Cell = this.cell(CellAddress(address))

/**
 * Finds or creates a cell that matches [address]
 *
 * If there is no row corresponding to [address], then creates such a row
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, init: (cell: Cell) -> Unit): Cell = this.cell(address).apply(init)
// endregion


// region Get or create filled cells
/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Number): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: Number): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: Number): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Date): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: Date): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: Date): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Calendar): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: Calendar): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: Calendar): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: RichTextString): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: RichTextString): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: RichTextString): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: String): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: String): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: String): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Boolean): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [rowIndex] and [columnIndex]
 *
 * If there is no row corresponding to [rowIndex], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param rowIndex Row index; first row index - `0`
 * @param columnIndex Column index; first column index - `0`
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: CellAddress, value: Boolean): Cell = this.cell(address) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: CellAddress, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 *
 * @return Found or created cell
 */
fun Sheet.cell(address: String, value: Boolean): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }

/**
 * Finds or creates a cell for the specified [address]
 *
 * If there is no row corresponding to [address], then a new row with that index will be created
 *
 * Sets the found or created cell to [value]
 *
 * Finally, applies [init] to the found or created cell
 *
 * @param address Cell address such as "B4" or "H13" (as in MS Excel)
 * @param value The value to be set in the found or created cell
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Sheet.cell(address: String, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)
// endregion


// region Cell init
/**
 * Creates a cell initializer with the specified [afterCellInit] and [beforeCellInit] cell handlers; by default these handlers do nothing
 *
 * @param afterCellInit Cell handler; is applied to the cell after performing any actions with it using the [CellInit] [CellInit] methods
 * @param beforeCellInit Cell handler; applies to the cell before performing any actions with it using the [CellInit] [CellInit] methods
 *
 * @return Cell initializer
 */
fun Sheet.cellInit(
    afterCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    beforeCellInit: (cell: Cell) -> Unit = CellInit.NO_OP
): CellInit = CellInit(this, afterCellInit, beforeCellInit)


/**
 * Creates a cell initializer with the specified [afterCellInit] and [beforeCellInit] cell handlers; by default these handlers do nothing
 *
 * Finally, applies [init] to the created cell initializer
 *
 * @param afterCellInit Cell handler; is applied to the cell after performing any actions with it using the [CellInit] [CellInit] methods
 * @param beforeCellInit Cell handler; applies to the cell before performing any actions with it using the [CellInit] [CellInit] methods
 * @param init The block to be applied to the created cell initializer
 *
 * @return Cell initializer
 */
inline fun Sheet.cells(
    noinline afterCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    noinline beforeCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    init: (cells: CellInit) -> Unit
): CellInit = this.cellInit(afterCellInit, beforeCellInit).apply(init)
// endregion
