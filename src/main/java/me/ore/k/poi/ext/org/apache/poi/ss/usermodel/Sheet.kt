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

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.util.CellAddress
import org.apache.poi.util.Units
import me.ore.k.poi.ext.CellInit
import java.util.*
import kotlin.math.roundToInt


// region Rows
fun Sheet.getRow(rowTitle: String): Row? = this.getRow(CellAddress("A$rowTitle").row)

inline fun Sheet.getRow(rowIndex: Int, init: (row: Row) -> Unit): Row = this.getRow(rowIndex).apply(init)

fun Sheet.createRow(rowTitle: String): Row = this.createRow(CellAddress("A$rowTitle").row)

fun Sheet.row(rowIndex: Int): Row = this.getRow(rowIndex) ?: this.createRow(rowIndex)

fun Sheet.row(rowTitle: String): Row = this.row(CellAddress("A$rowTitle").row)
// endregion


// region Columns
fun Sheet.setColumnWidth(columnTitle: String, width: Int) {
    val address = CellAddress("${columnTitle}1")
    this.setColumnWidth(address.column, width)
}

fun Sheet.setColumnWidthInPixels(columnIndex: Int, width: Int) {
    val widthInUnits = ((Units.pixelToEMU(width).toDouble() / Units.EMU_PER_CHARACTER) * 256).roundToInt()
    this.setColumnWidth(columnIndex, widthInUnits)
}

fun Sheet.setColumnWidthInPixels(columnTitle: String, width: Int) {
    val address = CellAddress("${columnTitle}1")
    this.setColumnWidthInPixels(address.column, width)
}
// endregion


// region Get or create empty cells, fill cells
fun Sheet.cell(rowIndex: Int, columnIndex: Int): Cell = this.row(rowIndex).cell(columnIndex)
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex).apply(init)

fun Sheet.cell(address: CellAddress): Cell = this.cell(address.row, address.column)
inline fun Sheet.cell(address: CellAddress, init: (cell: Cell) -> Unit): Cell = this.cell(address).apply(init)

fun Sheet.cell(address: String): Cell = this.cell(CellAddress(address))
inline fun Sheet.cell(address: String, init: (cell: Cell) -> Unit): Cell = this.cell(address).apply(init)
// endregion


// region Get or create filled cells
fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Number): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: Number): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: Number): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Date): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: Date): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: Date): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Calendar): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: Calendar): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: Calendar): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: RichTextString): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: RichTextString): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: RichTextString): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: String): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: String): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: String): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)


fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Boolean): Cell = this.cell(rowIndex, columnIndex) { it.setCellValue(value) }
inline fun Sheet.cell(rowIndex: Int, columnIndex: Int, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, value).apply(init)

fun Sheet.cell(address: CellAddress, value: Boolean): Cell = this.cell(address) { it.setCellValue(value) }
inline fun Sheet.cell(address: CellAddress, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)

fun Sheet.cell(address: String, value: Boolean): Cell = this.cell(CellAddress(address)) { it.setCellValue(value) }
inline fun Sheet.cell(address: String, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(address, value).apply(init)
// endregion


// region Cell init
fun Sheet.cellInit(
    afterCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    beforeCellInit: (cell: Cell) -> Unit = CellInit.NO_OP
): CellInit = CellInit(this, afterCellInit, beforeCellInit)

inline fun Sheet.cells(
    noinline afterCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    noinline beforeCellInit: (cell: Cell) -> Unit = CellInit.NO_OP,
    init: (cells: CellInit) -> Unit
): CellInit = this.cellInit(afterCellInit, beforeCellInit).apply(init)
// endregion
