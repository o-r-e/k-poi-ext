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

package me.ore.k.poi.ext

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellAddress
import java.util.*


class CellInit (
    val sheet: Sheet,
    val afterCellInit: (cell: Cell) -> Unit = NO_OP,
    val beforeCellInit: (cell: Cell) -> Unit = NO_OP
) {
    companion object {
        val NO_OP = { _: Cell -> }
    }


    // region Init cell
    inline fun initCell(rowIndex: Int, columnIndex: Int, type: CellType? = null, init: (cell: Cell) -> Unit): Cell {
        val row = this.sheet.getRow(rowIndex)
            ?: this.sheet.createRow(rowIndex)

        val result = row.getCell(columnIndex) ?: run {
            if (type == null) {
                row.createCell(columnIndex)
            } else {
                row.createCell(columnIndex, type)
            }
        }
        result.apply(this.beforeCellInit)
        result.apply(init)
        result.apply(this.afterCellInit)
        return result
    }

    inline fun initCell(address: CellAddress, type: CellType? = null, init: (cell: Cell) -> Unit): Cell =
        this.initCell(address.row, address.column, type, init)

    inline fun initCell(address: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell =
        this.initCell(CellAddress(address), type, init)
    // endregion


    // region Empty cell
    fun cell(rowIndex: Int, columnIndex: Int, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {}
    inline fun cell(rowIndex: Int, columnIndex: Int, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type, init)

    fun cell(address: CellAddress, type: CellType? = null): Cell = this.initCell(address, type) {}
    inline fun cell(address: CellAddress, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type, init)

    fun cell(address: String, type: CellType? = null): Cell = this.initCell(address, type) {}
    inline fun cell(address: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type, init)
    // endregion


    // region Create and fill cell
    fun cell(rowIndex: Int, columnIndex: Int, value: Number, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value.toDouble())
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    fun cell(rowIndex: Int, columnIndex: Int, value: Date, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(rowIndex: Int, columnIndex: Int, value: Calendar, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(rowIndex: Int, columnIndex: Int, value: RichTextString, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(rowIndex: Int, columnIndex: Int, value: String, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(rowIndex: Int, columnIndex: Int, value: Boolean, type: CellType? = null): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }
    fun cell(rowIndex: Int, columnIndex: Int, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }


    fun cell(address: CellAddress, value: Number, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value.toDouble())
    }
    fun cell(address: CellAddress, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    fun cell(address: CellAddress, value: Date, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: CellAddress, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: CellAddress, value: Calendar, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: CellAddress, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: CellAddress, value: RichTextString, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: CellAddress, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: CellAddress, value: String, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: CellAddress, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: CellAddress, value: Boolean, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: CellAddress, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }


    fun cell(address: String, value: Number, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value.toDouble())
    }
    fun cell(address: String, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    fun cell(address: String, value: Date, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: String, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: String, value: Calendar, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: String, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: String, value: RichTextString, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: String, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: String, value: String, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: String, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    fun cell(address: String, value: Boolean, type: CellType? = null): Cell = this.initCell(address, type) {
        it.setCellValue(value)
    }
    fun cell(address: String, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.initCell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }
    // endregion
}
