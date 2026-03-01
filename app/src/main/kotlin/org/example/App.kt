package org.example

import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.WorkbookFactory
import java.io.File

class App {
    val greeting: String
        get() {
            return "Hello World!"
        }
}



// ─── Grade Logic ─────────────────────────────────────────────────────────────

fun getGrade(score: Double): String = when {
    score >= 90.0 -> "A"
    score >= 80.0 -> "B"
    score >= 70.0 -> "C"
    score >= 60.0 -> "D"
    else          -> "F"
}

fun processStudent(name: String, score: Double?) {
    if (score == null) {
        println("No score for $name")
    } else {
        println("$name scored $score : Grade ${getGrade(score)}")
    }
}

// ─── Excel Reader (Apache POI) ────────────────────────────────────────────────

data class Student(val name: String, val score: Double?)

fun readExcel(path: String): List<Student> {
    val workbook = WorkbookFactory.create(File(path))
    val sheet = workbook.getSheetAt(0)
    val students = mutableListOf<Student>()

    // Find header row indexes
    val headerRow = sheet.getRow(0) ?: return emptyList()
    var nameCol = -1
    var scoreCol = -1
    for (cell in headerRow) {
        when (cell.stringCellValue.trim().lowercase()) {
            "name"  -> nameCol  = cell.columnIndex
            "score" -> scoreCol = cell.columnIndex
        }
    }

    if (nameCol == -1 || scoreCol == -1) {
        error("Excel file must have 'Name' and 'Score' columns in the first row.")
    }

    for (row in sheet.drop(1)) {
        val nameCell  = row.getCell(nameCol)  ?: continue
        val scoreCell = row.getCell(scoreCol)

        val name = nameCell.stringCellValue.trim()
        if (name.isEmpty()) continue

        val score: Double? = when {
            scoreCell == null                        -> null
            scoreCell.cellType == CellType.BLANK     -> null
            scoreCell.cellType == CellType.NUMERIC   -> scoreCell.numericCellValue.toDouble()
            scoreCell.cellType == CellType.STRING    -> scoreCell.stringCellValue.trim().toDoubleOrNull()
            else                                     -> null
        }

        students.add(Student(name, score))
    }

    workbook.close()
    return students
}

// ─── Main ─────────────────────────────────────────────────────────────────────

fun main(args: Array<String>) {
    val excelPath = if (args.isNotEmpty()) args[0] else "students.xlsx"

    if (!File(excelPath).exists()) {
        println("Error: File \"$excelPath\" not found.")
        return
    }

    println("=== Student Grade Calculator (Kotlin) ===\n")

    val students = readExcel(excelPath)
    students.forEach { processStudent(it.name, it.score) }
}