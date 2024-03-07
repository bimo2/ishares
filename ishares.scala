import java.io.File
import java.io.FileWriter
import java.sql.Date
import java.text.SimpleDateFormat
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import scala.jdk.CollectionConverters._

val schemaSQL = """
DROP USER IF EXISTS device;
DROP DATABASE IF EXISTS ishares;
CREATE DATABASE ishares;

CREATE TABLE etfs (
  id VARCHAR(10) PRIMARY KEY,
  issuer VARCHAR(255),
  name VARCHAR(255),
  report_date DATE
);

CREATE TABLE assets (
  id VARCHAR(10) PRIMARY KEY,
  name VARCHAR(255),
  sector VARCHAR(255),
  class VARCHAR(255),
  location VARCHAR(255),
  exchange VARCHAR(255)
);

CREATE TABLE quotes (
  etf_id VARCHAR(10) NOT NULL,
  date DATE NOT NULL,
  nav_per_share DECIMAL(18, 2),
  shares DECIMAL(18, 4),
  PRIMARY KEY (etf_id, date),
  FOREIGN KEY (etf_id) REFERENCES etfs(id)
);

CREATE TABLE holdings (
  etf_id VARCHAR(10) NOT NULL,
  asset_id VARCHAR(10) NOT NULL,
  market_value DECIMAL(18, 2),
  weight FLOAT,
  notional_value DECIMAL(18, 2),
  shares DECIMAL(18, 4),
  PRIMARY KEY (etf_id, asset_id),
  FOREIGN KEY (etf_id) REFERENCES etfs(id),
  FOREIGN KEY (asset_id) REFERENCES assets(id)
);

CREATE TABLE dividends (
  etf_id VARCHAR(10) NOT NULL,
  record_date DATE NOT NULL,
  ex_date DATE,
  payable_date DATE,
  value DECIMAL(18, 2) NOT NULL,
  PRIMARY KEY (etf_id, record_date),
  FOREIGN KEY (etf_id) REFERENCES etfs(id)
);

CREATE USER 'device'@'%' IDENTIFIED BY 'CL0UD5Q1';
GRANT ALL PRIVILEGES ON ishares.* TO 'device'@'%';
FLUSH PRIVILEGES;
""".stripMargin.trim

def holdingsSQL(sheet: Sheet): String = {
  val xlsxDate = new SimpleDateFormat("dd-MMM-yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")
  val date = sqlDate.format(xlsxDate.parse(sheet.getRow(0).getCell(0).getStringCellValue()))
  val etfName = sheet.getRow(1).getCell(0).getStringCellValue()

  // TODO: return ETF data
  println(s"$etfName ($date)")

  val etfSQL = s"INSERT INTO etfs (id, issuer, name, report_date)\nVALUES\n  ('ETF', 'BlackRock, Inc.', '$etfName', '$date');"
  val assets = new StringBuilder
  val holdings = new StringBuilder

  for (index <- 8 to sheet.getLastRowNum) {
    val row = sheet.getRow(index)
    val ticker = row.getCell(0).getStringCellValue()
    val name = row.getCell(1).getStringCellValue()
    val sector = row.getCell(2).getStringCellValue()
    val assetClass = row.getCell(3).getStringCellValue()
    val marketValue = BigDecimal(row.getCell(4).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
    val weight = BigDecimal(row.getCell(5).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
    val notionalValue = BigDecimal(row.getCell(6).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
    val shares = BigDecimal(row.getCell(7).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
    val location = row.getCell(9).getStringCellValue()
    val exchange = row.getCell(10).getStringCellValue()

    assets.append(s"\n  ('$ticker', '$name', '$sector', '$assetClass', '$location', '$exchange'),")
    holdings.append(s"\n  ('ETF', '$ticker', $marketValue, $weight, $notionalValue, $shares),")
  }

  assets.update(assets.length - 1, ';')
  holdings.update(holdings.length - 1, ';');

  val assetsSQL = s"\n\nINSERT INTO assets (id, name, sector, class, location, exchange)\nVALUES${assets.toString()}"
  val holdingsSQL = s"\n\nINSERT INTO holdings (etf_id, asset_id, market_value, weight, notional_value, shares)\nVALUES${holdings.toString()}"

  etfSQL + assetsSQL + holdingsSQL
}

def historicalSQL(sheet: Sheet, etf: String): String = {
  val xlsxDate = new SimpleDateFormat("MMM dd, yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")
  val rows = sheet.iterator().asScala.drop(1)

  val data = rows.map { row =>
    val date = sqlDate.format(xlsxDate.parse(row.getCell(0).getStringCellValue()))
    val value = BigDecimal(row.getCell(1).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
    val shares = BigDecimal(row.getCell(3).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)

    s"\n  ('$etf', '$date', $value, $shares)"
  }

  s"INSERT INTO quotes (etf_id, date, nav_per_share, shares)\nVALUES${data.mkString(",")};"
}

def distributionsSQL(sheet: Sheet, etf: String): String = {
  val xlsxDate = new SimpleDateFormat("MMM dd, yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")
  val rows = sheet.iterator().asScala.drop(1)

  val data = rows.map { row =>
    val recordDate = sqlDate.format(xlsxDate.parse(row.getCell(0).getStringCellValue()))
    val exDate = sqlDate.format(xlsxDate.parse(row.getCell(1).getStringCellValue()))
    val payableDate = sqlDate.format(xlsxDate.parse(row.getCell(2).getStringCellValue()))
    val value = BigDecimal(row.getCell(3).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)

    s"\n  ('$etf', '$recordDate', '$exDate', '$payableDate', $value)"
  }

  s"INSERT INTO dividends (etf_id, record_date, ex_date, payable_date, value)\nVALUES${data.mkString(",")};"
}

def sql(file: File): Unit = {
  println(s"sql <- ${file.getName}")

  val path = s"${file.getName.stripSuffix(".xlsx")}.sql"
  val workbook = WorkbookFactory.create(file)

  for (index <- 0 until workbook.getNumberOfSheets) {
    val sheet = workbook.getSheetAt(index)

    val sql = sheet.getSheetName match {
      case "Holdings" => holdingsSQL(sheet)
      case "Historical" => historicalSQL(sheet, "ETF")
      case "Distributions" => distributionsSQL(sheet, "ETF")
      case _ => s"TODO"
    }

    sqlout(path, sql)
  }

  workbook.close()
}

def sqlout(path: String, statement: String): Unit = {
  val file = new File(s"sql/$path")
  val directory = file.getParentFile

  if (!directory.exists) directory.mkdirs()
  if (!file.exists) file.createNewFile()

  val writer = new FileWriter(file, true)

  try {
    writer.write(statement)
  } finally {
    writer.close()
  }
}

@main def script(args: String*): Unit = {
  sqlout("schema.sql", schemaSQL)

  val path = new File("xlsx")

  if (!path.isDirectory) return

  val files = path.listFiles.filter(_.isFile)

  files.foreach(sql)
}
