# Student Grade Calculator — Kotlin

A command-line application built with Kotlin that reads student names and scores from an Excel file and prints their corresponding letter grades to the console.

> **Course** — Introduction to Android Development
**Lecturer** — Eng. MOUNE
**Exercise 1**

---

## Features

- Reads student data from a `.xlsx` Excel file using **Apache POI**
- Automatically detects `Name` and `Score` column positions from the header row
- Calculates and displays a letter grade for each student
- Handles missing or blank scores gracefully

## Grading Scale

| Grade | Score Range     |
|-------|-----------------|
| A     | 90 and above    |
| B     | 80 – 89         |
| C     | 70 – 79         |
| D     | 60 – 69         |
| F     | Below 60        |

---

## Requirements

- JDK 8 or higher
- Kotlin
- Apache POI (for Excel file parsing)

## Input File Format

The Excel file must contain a header row with at least two columns named `Name` and `Score` (case-insensitive). Column order does not matter.

**Example (`students.xlsx`):**

| Name    | Score |
|---------|-------|
| Alice   | 95    |
| Bob     | 72    |
| Charlie | 58    |
| Diana   |       |

## Usage

Run the application from the command line, optionally passing the path to the Excel file:

```bash
# Uses default file: students.xlsx in the working directory
kotlin MainKt

# Or specify a custom file path
kotlin MainKt path/to/your/students.xlsx
```

## Sample Output

```
=== Student Grade Calculator (Kotlin) ===

Alice scored 95.0 : Grade A
Bob scored 72.0 : Grade C
Charlie scored 58.0 : Grade F
No score for Diana
```

## Project Structure

```
src/
└── main/
    └── kotlin/
        └── org/example/
            └── App.kt        # Main application logic
```

### Key Functions

| Function | Description |
|----------|-------------|
| `getGrade(score)` | Returns the letter grade for a given numeric score |
| `processStudent(name, score)` | Prints the grade result or a missing score notice |
| `readExcel(path)` | Parses the Excel file and returns a list of `Student` objects |
| `main(args)` | Entry point; accepts an optional file path argument |