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
import org.apache.poi.ss.util.CellRangeAddress
import org.apache.poi.ss.util.CellReference
import java.util.*
import kotlin.collections.ArrayList


// region Height
/**
 * Calculates a suitable line height based on the cells in the current line.
 * The result is the maximum result of calling [Cell.calculateSuitableHeight] [Cell.calculateSuitableHeight] for each cell of the current row
 *
 * @param ignoreMergedRegions If it is `true`, the calculation will not take into account the merged regions
 *
 * @return Suitable height for current row
 */
fun Row.calculateSuitableHeight(ignoreMergedRegions: Boolean = false): Double {
    var result = this.heightInPoints.toDouble()

    for (cellIndex in this.firstCellNum .. this.lastCellNum) {
        this.getCell(cellIndex)?.let { cell ->
            val region = if (ignoreMergedRegions) {
                null
            } else {
                this.sheet.mergedRegions.firstOrNull {
                    it.containsRow(this.rowNum) && it.containsColumn(cellIndex)
                }
            }

            var cellHeight = if (region == null) {
                cell.calculateSuitableHeight(true)
            } else {
                cell.calculateSuitableHeight(false)
            }

            if (region != null) {
                for (mergedRowIndex in (region.firstRow + 1) .. region.lastRow) {
                    val mergedRowHeight = this.sheet.getRow(mergedRowIndex)?.heightInPoints
                        ?: this.sheet.defaultRowHeightInPoints

                    cellHeight -= mergedRowHeight
                }
            }

            if (result < cellHeight) {
                result = cellHeight
            }
        }
    }

    return result
}

/**
 * Changes the height of the current row to fit the contents of the "tallest" cell (see [Cell.calculateSuitableHeight] [Cell.calculateSuitableHeight])
 *
 * If the "tallest" cell is located on two or more lines (included in the merged region), there are two options for changing the height of the current line:
 *
 * * [stretchMerged] = `true` - the height of all rows to which such a cell belongs is proportionally changed
 * * [stretchMerged] = `false` - only the current row height changes
 *
 * @param stretchMerged Indicates whether to change the height of other lines
 */
fun Row.stretchToMaxContent(stretchMerged: Boolean = true) {
    val mergedRegions = this.sheet.mergedRegions.filter { it.containsRow(this.rowNum) }

    if (mergedRegions.isEmpty()) {
        val rowHeight = this.calculateSuitableHeight(true).toFloat()
        if (this.heightInPoints < rowHeight) {
            this.heightInPoints = rowHeight
        }
    } else {
        val affectedRegions = mergedRegions.filter { it.firstRow == this.rowNum }

        val cellToRegionPairs = ArrayList<Pair<Cell, CellRangeAddress?>>()
        for (columnIndex in this.firstCellNum .. this.lastCellNum) {
            this.getCell(columnIndex)?.let { cell ->
                val region = affectedRegions.firstOrNull { it.firstColumn == columnIndex }
                cellToRegionPairs.add(cell to region)
            }
        }

        var height = this.heightInPoints.toDouble()
        cellToRegionPairs.forEach { pair ->
            if (pair.second == null) {
                val cellHeight = pair.first.calculateSuitableHeight(true)
                if (height < cellHeight) {
                    height = cellHeight
                }
            }
        }

        val affectedRows = ArrayList<Row>()
        var affectedRowHeightSum = 0.0
        cellToRegionPairs.forEach { pair ->
            pair.second?.let { region ->
                val cellHeight = pair.first.calculateSuitableHeight(region)
                if (height < cellHeight) {
                    val currentAffectedRows: MutableList<Row> = ArrayList()
                    var heightSum = 0.0
                    for (affectedRowIndex in region.firstRow .. region.lastRow) {
                        val affectedRow = this.sheet.getRow(affectedRowIndex)
                            ?: this.sheet.createRow(affectedRowIndex)
                        currentAffectedRows.add(affectedRow)
                        heightSum += affectedRow.heightInPoints
                    }

                    if (heightSum < cellHeight) {
                        height = cellHeight
                        affectedRowHeightSum = heightSum
                        affectedRows.clear()
                        affectedRows.addAll(currentAffectedRows)
                    }
                }
            }
        }

        if (affectedRows.isEmpty()) {
            if (this.heightInPoints < height) {
                this.heightInPoints = height.toFloat()
            }
        } else if (!stretchMerged) {
            height = height - affectedRowHeightSum + this.heightInPoints
            if (this.heightInPoints < height) {
                this.heightInPoints = height.toFloat()
            }
        } else {
            val heightMultiplier = (height / affectedRowHeightSum).toFloat()
            if (heightMultiplier > 1) {
                affectedRows.forEach { row ->
                    row.heightInPoints = row.heightInPoints * heightMultiplier
                }
            }
        }
    }
}
// endregion


// region Get or create cells
/**
 * If there is a cell in the current row and the specified column, applies [init] to it and returns it
 *
 * @param columnIndex The index of the column in which the cell is searched
 * @param init Block applied to the found cell
 *
 * @return Found cell or `null`
 */
inline fun Row.getCell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell? = this.getCell(columnIndex)?.apply(init)

/**
 * If there is a cell in the current row and the specified column, returns it
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the cell is searched
 *
 * @return Found cell or `null`
 */
fun Row.getCell(columnTitle: String): Cell? = this.getCell(CellReference.convertColStringToIndex(columnTitle))

/**
 * If there is a cell in the current row and the specified column, applies [init] to it and returns it
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the cell is searched
 * @param init Block applied to the found cell
 *
 * @return Found cell or `null`
 */
inline fun Row.getCell(columnTitle: String, init: (cell: Cell) -> Unit): Cell? = this.getCell(columnTitle)?.apply(init)


/**
 * Creates a cell in the current row in the specified column and applies [init] to it
 *
 * @param columnIndex The index of the column in which to create the cell
 * @param init The block to be applied to the cell
 *
 * @return Created cell
 */
inline fun Row.createCell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.createCell(columnIndex).apply(init)

/**
 * Creates a cell in the current row in the specified column
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which to create the cell
 *
 * @return Created cell
 */
fun Row.createCell(columnTitle: String): Cell = this.createCell(CellReference.convertColStringToIndex(columnTitle))

/**
 * Creates a cell in the current row in the specified column and applies [init] to it
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which to create the cell
 * @param init The block to be applied to the cell
 *
 * @return Created cell
 */
inline fun Row.createCell(columnTitle: String, init: (cell: Cell) -> Unit): Cell = this.createCell(columnTitle).apply(init)


/**
 * Searches for an existing cell in the current row in the specified column; if there is no such cell, then creates a new one
 *
 * @param columnIndex The index of the column in which the desired cell should be
 *
 * @return Found or created cell; will never be `null`
 */
fun Row.cell(columnIndex: Int): Cell = this.getCell(columnIndex) ?: this.createCell(columnIndex)

/**
 * Searches for an existing cell in the current row in the specified column; if there is no such cell, then creates a new one.
 *
 * After applies to cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell; will never be `null`
 */
inline fun Row.cell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex).apply(init)

/**
 * Searches for an existing cell in the current row in the specified column; if there is no such cell, then creates a new one
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 *
 * @return Found or created cell; will never be `null`
 */
fun Row.cell(columnTitle: String): Cell = this.getCell(columnTitle) ?: this.createCell(columnTitle)

/**
 * Searches for an existing cell in the current row in the specified column; if there is no such cell, then creates a new one.
 *
 * After applies to cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell; will never be `null`
 */
inline fun Row.cell(columnTitle: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle).apply(init)
// endregion


// region Fill cells
/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: Number): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: Number): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: Date): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: Date): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: Calendar): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: Calendar): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: RichTextString): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: RichTextString): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: String): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: String): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnIndex: Int, value: Boolean): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnIndex The index of the column in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnIndex: Int, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 *
 * @return Found or created cell
 */
fun Row.cell(columnTitle: String, value: Boolean): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }

/**
 * Finds or creates a cell in the current row in the specified column and assigns [value] to that cell
 *
 * After applies to this cell [init]
 *
 * @param columnTitle The title of the column ("A", "B", etc.) in which the desired cell should be
 * @param value Cell value
 * @param init Block to be applied to the found or created cell
 *
 * @return Found or created cell
 */
inline fun Row.cell(columnTitle: String, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)
// endregion
