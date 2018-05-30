package demo1

import jdk.nashorn.internal.objects.NativeString
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream
import java.io.FileOutputStream
import java.text.SimpleDateFormat
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter


const val filePath = "C:/Users/tellezga/Dropbox/Kotlin/jug/"
const val fileExtension = ".xlsx"

const val demoFile1Name = "DemoFile"
const val demo1FilePath = filePath + demoFile1Name + fileExtension

const val demoFile2Name = "ExtractDataFromHere"
const val demo2FilePath = filePath + demoFile2Name + fileExtension

val formatter = DataFormatter()


val map: MutableMap<String, DeviceType> = HashMap<String, DeviceType>()

fun main(args: Array<String>) {
    val workbook = openXLSXFile(demo1FilePath)
    //processStuff(workbook)
    loadDataToMemory(openXLSXFile(demo2FilePath))
    var updatedWorkbook = operateWithData(workbook)
    updatedWorkbook = useDatesAndCellFormat(updatedWorkbook)
    writeFile(updatedWorkbook)
}


fun openXLSXFile(fileLocation: String): XSSFWorkbook {
    val file = FileInputStream(File(fileLocation))
    //Todo: Inline return val
    return XSSFWorkbook(file)
}

fun processStuff(workbook: Workbook) {

    val sheet = workbook.getSheetAt(0)
    val evaluator = workbook.creationHelper.createFormulaEvaluator()

    for (row in sheet) {
        for (cell in row) {
            val cellValue = evaluator.evaluateInCell(cell)
            print(getStringCellValue(cellValue))
        }
        print("\n")
    }
}


fun getStringCellValue(cell: Cell): String =
//Todo: "when" as expression
        when (cell.cellTypeEnum) {
            CellType.STRING -> cell.stringCellValue + "\t"
            CellType.NUMERIC -> cell.numericCellValue.toString() + "\t"
            CellType.BOOLEAN -> "" + cell.booleanCellValue + "\t"
            else -> "UNDEFINED "
        }
//Todo: Property access syntax

fun loadDataToMemory(workbook: Workbook) {
    val kaiserDeviceInfoSheet = workbook.getSheetAt(0)
    val formatter = DataFormatter()

    for (rowNum in 1..kaiserDeviceInfoSheet.lastRowNum) {
        var cell = kaiserDeviceInfoSheet.getRow(rowNum).getCell(4)

        val partNumber = formatter.formatCellValue(cell)
        val deviceTypeId = formatter.formatCellValue(kaiserDeviceInfoSheet.getRow(rowNum).getCell(1))
        val deviceTypeName = formatter.formatCellValue(cell.row.getCell(2))

        if (!map.containsKey(partNumber)) {

            val deviceType: DeviceType = DeviceType(
                    partNumber = partNumber,
                    deviceTypeId = deviceTypeId.toInt(),
                    deviceTypeName = deviceTypeName.toUpperCase())  //Extension function

            map[partNumber] = deviceType

            //TODO: Printing redundant fields: Use String templates
            //print("part: $partNumber, typeId: $deviceTypeId, deviceType: $deviceType")
            //println()
        }
    }
    //Todo: Print using sortedBy
    println("Print deviceTypes by Id")
    (map.values.sortedBy {
        it.deviceTypeId
    }).forEach {
        println(it)
    }


}

fun operateWithData(workbook: Workbook): Workbook {

    val sheet = workbook.getSheetAt(0)

    for (rowNum in 1..sheet.lastRowNum) {
        val row = sheet.getRow(rowNum) ?: break;
        //Get 4th column: Part number
        val cell = row.getCell(3) ?: break
        val partNumber = formatter.formatCellValue(cell)


        //Get cell to write deviceTypeID, at column 7th (G)
        var escribemeCell = row.getCell(6)
        if (escribemeCell == null) {
            escribemeCell = row.createCell(6)
        }
        escribemeCell.setCellType(CellType.STRING);
        //Todo: Null safety
        val deviceTypeId = map[partNumber]!!.deviceTypeId
        //Numeric Operations
        escribemeCell.setCellValue(deviceTypeId.toString())


        //Get cell to write deviceTypeName, at 8th column  (H)
        val deviceTypeNameCell = setStringValueOnCell(
                value = map[partNumber]?.deviceTypeName.toString(),
                cell = row.getCell(7),
                row = row,
                index = 7)

        println("partNumber: $partNumber, deviceTypeId: $deviceTypeId, deviceTypeName: ${deviceTypeNameCell.stringCellValue}")

        val manufacturerCell = row.getCell(1)
        manufacturerCell.setCellType(CellType.STRING);
        val deviceNameCell = row.getCell(2)
        deviceNameCell.setCellType(CellType.STRING);

        //Also, Concat A + B, with a space and UPPER CASE, at 10th column (J)
        val concatenatedValue= NativeString.toUpperCase(manufacturerCell) + " " + deviceNameCell.stringCellValue.toUpperCase() + ""
        setStringValueOnCell(
                row = row,
                index = 9,
                cell = row.getCell(9),
                value = concatenatedValue)


        //Does it match the cross referenced. 11th column  (K)
        val doNamesMatch = concatenatedValue.equals(map[partNumber]?.deviceTypeName, true)

        val asdf = setStringValueOnCell(
                cell = row.getCell(10),
                index = 10,
                row = row,

                value = if (doNamesMatch) "" else "nop"
        )
    }
    return workbook
}

fun setStringValueOnCell
        (row: Row, index: Int, cell: Cell, value: String = "")
        =
        (if (cell == null) row.createCell(index) else cell).apply {
            setCellType(CellType.STRING)
            setCellValue(value)
        }


fun useDatesAndCellFormat(workbook: Workbook): Workbook {
    val sheet = workbook.getSheetAt(0)

    val sdf = SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss'Z'")
    val deadline = sdf.parse("2016-05-13T20:43:44Z")


    for (rowNum in 1..sheet.lastRowNum) {
        val row = sheet.getRow(rowNum) ?: break;
        //Get 13th column: Collection Date
        val cell = row.getCell(12) ?: break
        val cellDate = formatter.formatCellValue(cell)
        println("date: $cellDate")

        val collectionDate = sdf.parse(cellDate)


        when {
            collectionDate > deadline -> println("collectionDate is after deadline")

            collectionDate < deadline -> println("collectionDate is before deadline")
            collectionDate == deadline -> println("collectionDate is equal to deadline")
            else -> System.out.println("How to get here?")
        }


        println("compareTo: " + collectionDate.compareTo(deadline))

        //Todo: "when" as statement, again

        when {  //Todo: Simple date comparison syntax
            collectionDate > deadline ->
                setStyle(cell, workbook, IndexedColors.RED.getIndex())
            collectionDate < deadline -> setStyle(cell, workbook, IndexedColors.LIGHT_GREEN.getIndex())
            collectionDate == deadline -> setStyle(cell, workbook, IndexedColors.YELLOW.getIndex())
            else -> setStyle(cell, workbook, IndexedColors.GREY_50_PERCENT.getIndex())
        }
    }
    return workbook
}

fun setStyle(cell: Cell, workbook: Workbook, colorIndex: Short): Unit {

    cell.run {
        cellStyle = workbook.createCellStyle().also {
            it.setFillPattern(FillPatternType.SOLID_FOREGROUND)
            it.fillForegroundColor = colorIndex
        }
    }
}

fun writeFile(workbook: Workbook) {
    try {

        val now = LocalDateTime.now()

        val dateFormatter = DateTimeFormatter.ofPattern("yyyy-MM-dd HH-mm-ss-SSS")
        val formattedDate = now.format(dateFormatter)

        println("Current Date and Time is: $formattedDate")

        val fileOut = FileOutputStream(filePath + demoFile1Name + "output" + "_" + formattedDate + fileExtension)
        workbook.write(fileOut)
    } catch (e: Exception) {
        println("Exception: " + e.message)
    }

}


