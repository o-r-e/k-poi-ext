@file:Suppress("unused")

package r.obuhov.k.poi.ext.org.apache.poi.ss.usermodel

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.RichTextString
import org.apache.poi.ss.usermodel.Row
import org.apache.poi.ss.util.CellAddress
import org.apache.poi.ss.util.CellRangeAddress
import java.util.*
import kotlin.collections.ArrayList


// region Height
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

fun Row.stretchToMaxContent(stretchMergedProportionally: Boolean = true) {
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
        } else if (!stretchMergedProportionally) {
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
inline fun Row.getCell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell? = this.getCell(columnIndex)?.apply(init)

fun Row.getCell(columnTitle: String): Cell? = this.getCell(CellAddress("${columnTitle}${this.rowNum + 1}").column)
inline fun Row.getCell(columnTitle: String, init: (cell: Cell) -> Unit): Cell? = this.getCell(columnTitle)?.apply(init)


inline fun Row.createCell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.createCell(columnIndex).apply(init)

fun Row.createCell(columnTitle: String): Cell = this.createCell(CellAddress("${columnTitle}${this.rowNum + 1}").column)
inline fun Row.createCell(columnTitle: String, init: (cell: Cell) -> Unit): Cell = this.createCell(columnTitle).apply(init)

fun Row.cell(columnIndex: Int): Cell = this.getCell(columnIndex) ?: this.createCell(columnIndex)
inline fun Row.cell(columnIndex: Int, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex).apply(init)

fun Row.cell(columnTitle: String): Cell = this.getCell(columnTitle) ?: this.createCell(columnTitle)
inline fun Row.cell(columnTitle: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle).apply(init)
// endregion


// region Fill cells
fun Row.cell(columnIndex: Int, value: Number): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: Number): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: Number, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


fun Row.cell(columnIndex: Int, value: Date): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: Date): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: Date, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


fun Row.cell(columnIndex: Int, value: Calendar): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: Calendar): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: Calendar, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


fun Row.cell(columnIndex: Int, value: RichTextString): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: RichTextString): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: RichTextString, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


fun Row.cell(columnIndex: Int, value: String): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: String): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: String, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)


fun Row.cell(columnIndex: Int, value: Boolean): Cell = this.cell(columnIndex).apply { this.setCellValue(value) }
inline fun Row.cell(columnIndex: Int, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(columnIndex, value).apply(init)

fun Row.cell(columnTitle: String, value: Boolean): Cell = this.cell(columnTitle).apply { this.setCellValue(value) }
inline fun Row.cell(columnTitle: String, value: Boolean, init: (cell: Cell) -> Unit): Cell = this.cell(columnTitle, value).apply(init)
// endregion
