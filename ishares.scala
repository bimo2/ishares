import java.io.File
import java.io.FileWriter
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
  asset_class VARCHAR(255),
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

def column(sheet: Sheet, head: Int, name: String): Option[Int] = {
  val header = sheet.getRow(head)

  if (header != null) {
    val iterator = header.cellIterator()
    var index = 0

    while (iterator.hasNext) {
      val cell = iterator.next()
      val string = cell.getStringCellValue()

      if (string == name) return Some(index)

      index += 1
    }
  }

  None
}

def holdingsSQL(sheet: Sheet, etf: String): String = {
  val xlsxDate = new SimpleDateFormat("dd-MMM-yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")
  val reportDate = sqlDate.format(xlsxDate.parse(sheet.getRow(0).getCell(0).getStringCellValue()))
  val etfSQL = s"\nUPDATE etfs SET report_date = '$reportDate' WHERE id = '$etf';"

  val assets = new StringBuilder
  val holdings = new StringBuilder

  val indices = Map(
    "id" -> column(sheet, 7, "Ticker"),
    "name" -> column(sheet, 7, "Name"),
    "sector" -> column(sheet, 7, "Sector"),
    "asset_class" -> column(sheet, 7, "Asset Class"),
    "market_value" -> column(sheet, 7, "Market Value"),
    "weight" -> column(sheet, 7, "Weight (%)"),
    "notional_value" -> column(sheet, 7, "Notional Value"),
    "shares" -> column(sheet, 7, "Shares"),
    "location" -> column(sheet, 7, "Location"),
    "exchange" -> column(sheet, 7, "Exchange")
  )

  for (index <- 8 to sheet.getLastRowNum) {
    val row = sheet.getRow(index)

    val id = indices("id") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    val name = indices("name") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    val sector = indices("sector") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    val assetClass = indices("asset_class") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    val marketValue = indices("market_value") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    val weight = indices("weight") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    val notionalValue = indices("notional_value") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    val shares = indices("shares") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    val location = indices("location") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    val exchange = indices("exchange") match {
      case Some(index) => row.getCell(index).getStringCellValue()
      case _ => null
    }

    assets.append(s"\n  ('$id', '$name', '$sector', '$assetClass', '$location', '$exchange'),")
    holdings.append(s"\n  ('$etf', '$id', $marketValue, $weight, $notionalValue, $shares),")
  }

  assets.update(assets.length - 1, ';')
  holdings.update(holdings.length - 1, ';');

  val assetsSQL = s"\nINSERT INTO assets (id, name, sector, asset_class, location, exchange)\nVALUES${assets.toString()}"
  val holdingsSQL = s"\nINSERT INTO holdings (etf_id, asset_id, market_value, weight, notional_value, shares)\nVALUES${holdings.toString()}"

  List(etfSQL, assetsSQL, holdingsSQL).mkString("\n")
}

def historicalSQL(sheet: Sheet, etf: String): String = {
  val xlsxDate = new SimpleDateFormat("MMM dd, yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")

  val indices = Map(
    "date" -> Some(0),
    "nav_per_share" -> column(sheet, 0, "NAV per Share"),
    "shares" -> column(sheet, 0, "Shares Outstanding")
  )

  val data = sheet.iterator().asScala.drop(1).map { row =>
    val date = indices("date") match {
      case Some(index) => sqlDate.format(xlsxDate.parse(row.getCell(index).getStringCellValue()))
      case _ => null
    }

    val navPerShare = indices("nav_per_share") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    val shares = indices("shares") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    s"\n  ('$etf', '$date', $navPerShare, $shares)"
  }

  s"\nINSERT INTO quotes (etf_id, date, nav_per_share, shares)\nVALUES${data.mkString(",")};"
}

def distributionsSQL(sheet: Sheet, etf: String): String = {
  val xlsxDate = new SimpleDateFormat("MMM dd, yyyy")
  val sqlDate = new SimpleDateFormat("yyyy-MM-dd")

  val indices = Map(
    "record_date" -> column(sheet, 0, "Record Date"),
    "ex_date" -> column(sheet, 0, "Ex-Date"),
    "payable_date" -> column(sheet, 0, "Payable Date"),
    "value" -> column(sheet, 0, "Total Distribution")
  )

  val data = sheet.iterator().asScala.drop(1).map { row =>
    val recordDate = indices("record_date") match {
      case Some(index) => sqlDate.format(xlsxDate.parse(row.getCell(index).getStringCellValue()))
      case _ => null
    }

    val exDate = indices("ex_date") match {
      case Some(index) => sqlDate.format(xlsxDate.parse(row.getCell(index).getStringCellValue()))
      case _ => null
    }

    val payableDate = indices("payable_date") match {
      case Some(index) => sqlDate.format(xlsxDate.parse(row.getCell(index).getStringCellValue()))
      case _ => null
    }

    val value = indices("value") match {
      case Some(index) => BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
      case _ => null
    }

    s"\n  ('$etf', '$recordDate', '$exDate', '$payableDate', $value)"
  }

  s"\nINSERT INTO dividends (etf_id, record_date, ex_date, payable_date, value)\nVALUES${data.mkString(",")};"
}

def sql(file: File): Unit = {
  println(s"sql <- ${file.getName}")

  val path = s"${file.getName.stripSuffix(".xlsx")}.sql"
  val workbook = WorkbookFactory.create(file)
  val etf = workbook.getSheet("Performance").getRow(0).getCell(0).getStringCellValue()

  sqlout(path, s"INSERT INTO etfs (id, issuer, name)\nVALUES\n  ('ETF', '$etf', 'BlackRock, Inc.');")

  for (index <- 0 until workbook.getNumberOfSheets) {
    val sheet = workbook.getSheetAt(index)

    val sql = sheet.getSheetName match {
      case "Holdings" => holdingsSQL(sheet, "ETF")
      case "Historical" => historicalSQL(sheet, "ETF")
      case "Distributions" => distributionsSQL(sheet, "ETF")
      case _ => ""
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
  if (statement.isEmpty) return

  val writer = new FileWriter(file, true)

  try {
    writer.write(s"$statement\n")
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
