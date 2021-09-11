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

@file:Suppress("unused", "MemberVisibilityCanBePrivate")

package me.ore.k.poi.ext

import org.apache.poi.ss.usermodel.*
import org.apache.poi.ss.util.CellAddress
import java.util.*


/**
 * Cell initializer
 *
 * Has several methods for working with cells.
 * In such methods, [beforeCellInit] is first applied to the found or created created cell, then any "initializing" block (if supported by the method)
 *   and at the end - [afterCellInit]
 */
class CellInit (
    /**
     * The sheet in which cells will be created and / or processed
     */
    val sheet: Sheet,

    /**
     * Final cell handler. The default is [NO_OP]
     */
    val afterCellInit: (cell: Cell) -> Unit = NO_OP,

    /**
     * Initial cell handler. The default is [NO_OP]
     */
    val beforeCellInit: (cell: Cell) -> Unit = NO_OP
) {
    companion object {
        /**
         * An "empty" cell handler that does nothing
         */
        val NO_OP = { _: Cell -> }
    }


    // region Find or create empty cell
    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Applied to the found or created cell: [beforeCellInit], then [init] and at the end - [afterCellInit]
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init Block to be applied to the found or created cell
     *
     * @return Found or created cell
     */
    inline fun cell(rowIndex: Int, columnIndex: Int, type: CellType? = null, init: (cell: Cell) -> Unit = {}): Cell {
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

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Applied to the found or created cell: [beforeCellInit], then [init] and at the end - [afterCellInit]
     *
     * @param address Cell address
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init Block to be applied to the found or created cell
     *
     * @return Found or created cell
     */
    inline fun cell(address: CellAddress, type: CellType? = null, init: (cell: Cell) -> Unit = {}): Cell =
        this.cell(address.row, address.column, type, init)

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Applied to the found or created cell: [beforeCellInit], then [init] and at the end - [afterCellInit]
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Office)
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init Block to be applied to the found or created cell
     *
     * @return Found or created cell
     */
    inline fun cell(address: String, type: CellType? = null, init: (cell: Cell) -> Unit = {}): Cell =
        this.cell(CellAddress(address), type, init)
    // endregion


    // region Fill found or new cell
    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Number, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value.toDouble())
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Date, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Calendar, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: RichTextString, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: String, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Boolean, type: CellType? = null): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [rowIndex] and [columnIndex]. If there is no row in the sheet that matches [rowIndex], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param rowIndex Row index; first Row index - `0`
     * @param columnIndex Column index; first column index - `0`
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(rowIndex: Int, columnIndex: Int, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(rowIndex, columnIndex, type) {
        it.setCellValue(value)
        it.apply(init)
    }


    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Number, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value.toDouble())
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Date, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Calendar, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: RichTextString, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: String, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Boolean, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: CellAddress, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }


    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Number, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value.toDouble())
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Number, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value.toDouble())
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Date, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Date, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Calendar, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Calendar, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: RichTextString, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: RichTextString, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: String, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: String, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Boolean, type: CellType? = null): Cell = this.cell(address, type) {
        it.setCellValue(value)
    }

    /**
     * Finds or creates a cell that matches [address]. If there is no row in the sheet that matches [address], then creates such a row.
     *
     * Execution order:
     * * Applies [beforeCellInit] to found or created cell
     * * Sets [value] to a cell
     * * Applies [init] to the cell
     * * Applies [afterCellInit] to the cell
     *
     * @param address Cell address such as "B4" or "H13" (as in MS Excel)
     * @param value Cell value
     * @param type Cell type; set only in a new cell and only if not equal to `null`
     * @param init The block to be applied to the cell
     *
     * @return Found or created cell
     */
    fun cell(address: String, value: Boolean, type: CellType? = null, init: (cell: Cell) -> Unit): Cell = this.cell(address, type) {
        it.setCellValue(value)
        it.apply(init)
    }
    // endregion
}
