import java.io.File
import java.io.FileWriter
import java.text.SimpleDateFormat
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.WorkbookFactory
import scala.jdk.CollectionConverters._

val iSharesETFs = Map(
  // https://www.ishares.com/us/products/333011/ishares-bitcoin-trust
  "iShares Bitcoin Trust" -> "IBIT",
  // https://www.ishares.com/us/products/326614/ishares-blockchain-and-tech-etf
  "iShares Blockchain and Tech ETF" -> "IBLC",
  // https://www.ishares.com/us/products/239726/ishares-core-sp-500-etf
  "iShares Core S&P 500 ETF" -> "IVV",
  // https://www.ishares.com/us/products/239737/ishares-global-100-etf
  "iShares Global 100 ETF" -> "IOO",
  // https://www.ishares.com/us/products/239742/ishares-global-financials-etf
  "iShares Global Financials ETF" -> "IXG",
  // https://www.ishares.com/us/products/239750/ishares-global-tech-etf
  "iShares Global Tech ETF" -> "IXN",
  // https://www.ishares.com/us/products/239561/ishares-gold-trust-fund
  "iShares Gold Trust" -> "IAU",
  // https://www.ishares.com/us/products/239725/ishares-sp-500-growth-etf
  "iShares S&P 500 Growth ETF" -> "IVW",
  // https://www.ishares.com/us/products/239728/ishares-sp-500-value-etf
  "iShares S&P 500 Value ETF" -> "IVE",
  // https://www.ishares.com/us/products/239705/ishares-phlx-semiconductor-etf
  "iShares Semiconductor ETF" -> "SOXX",
  // https://www.ishares.com/us/products/292414/ishares-u-s-consumer-focused-etf
  "iShares U.S. Consumer Focused ETF" -> "IEDI",
  // https://www.ishares.com/us/products/239507/ishares-us-energy-etf
  "iShares U.S. Energy ETF" -> "IYE",
  // https://www.ishares.com/us/products/239508/ishares-us-financials-etf
  "iShares U.S. Financials ETF" -> "IYF",
  // https://www.ishares.com/us/products/239509/ishares-us-financial-services-etf
  "iShares U.S. Financial Services ETF" -> "IYG",
  // https://www.ishares.com/us/products/239511/ishares-us-healthcare-etf
  "iShares U.S. Healthcare ETF" -> "IYH",
  // https://www.ishares.com/us/products/239522/ishares-us-technology-etf
  "iShares U.S. Technology ETF" -> "IYW"
)

val schemaSQL = """
DROP USER IF EXISTS device;
DROP DATABASE IF EXISTS ishares;
CREATE DATABASE ishares;

CREATE TABLE ishares.etfs (
  id VARCHAR(10) PRIMARY KEY,
  issuer VARCHAR(255),
  name VARCHAR(255),
  report_date DATE
);

CREATE TABLE ishares.assets (
  id VARCHAR(10) PRIMARY KEY,
  name VARCHAR(255),
  sector VARCHAR(255),
  asset_class VARCHAR(255),
  location VARCHAR(255),
  exchange VARCHAR(255)
);

CREATE TABLE ishares.quotes (
  etf_id VARCHAR(10) NOT NULL,
  date DATE NOT NULL,
  nav_per_share DECIMAL(18, 2),
  shares DECIMAL(18, 4),
  PRIMARY KEY (etf_id, date),
  FOREIGN KEY (etf_id) REFERENCES etfs(id)
);

CREATE TABLE ishares.holdings (
  etf_id VARCHAR(10) NOT NULL,
  asset_id VARCHAR(10) NOT NULL,
  market_value DECIMAL(18, 2),
  weight FLOAT,
  notional_value DECIMAL(18, 2),
  shares DECIMAL(18, 4),
  INDEX (etf_id, asset_id),
  FOREIGN KEY (etf_id) REFERENCES etfs(id),
  FOREIGN KEY (asset_id) REFERENCES assets(id)
);

CREATE TABLE ishares.dividends (
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

def sqlnull(value: Any): String = {
  value match {
    case string: String => {
      string match {
        case "--" => "NULL"
        case _ => s"'$string'"
      }
    }
    case decimal: BigDecimal => s"$decimal"
    case _ => "NULL"
  }
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
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
      case _ => null
    }

    val weight = indices("weight") match {
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
      case _ => null
    }

    val notionalValue = indices("notional_value") match {
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
      case _ => null
    }

    val shares = indices("shares") match {
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
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

    assets.append(s"\n  ('$id', ${sqlnull(name)}, ${sqlnull(sector)}, ${sqlnull(assetClass)}, ${sqlnull(location)}, ${sqlnull(exchange)}),")
    holdings.append(s"\n  ('$etf', '$id', ${sqlnull(marketValue)}, ${sqlnull(weight)}, ${sqlnull(notionalValue)}, ${sqlnull(shares)}),")
  }

  assets.update(assets.length - 1, ';')
  holdings.update(holdings.length - 1, ';')

  val assetsSQL = s"\nINSERT IGNORE INTO assets (id, name, sector, asset_class, location, exchange)\nVALUES${assets.toString()}"
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
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(2, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
      case _ => null
    }

    val shares = indices("shares") match {
      case Some(index) => try BigDecimal(row.getCell(index).getNumericCellValue()).setScale(4, BigDecimal.RoundingMode.HALF_UP)
        catch case _: Throwable => null
      case _ => null
    }

    s"\n  ('$etf', '$date', ${sqlnull(navPerShare)}, ${sqlnull(shares)})"
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

    s"\n  ('$etf', '$recordDate', ${sqlnull(exDate)}, ${sqlnull(payableDate)}, $value)"
  }

  s"\nINSERT INTO dividends (etf_id, record_date, ex_date, payable_date, value)\nVALUES${data.mkString(",")};"
}

def sql(file: File): Unit = {
  println(s"sql <- ${file.getName}")

  val path = s"${file.getName.stripSuffix(".xlsx")}.sql"
  val workbook = WorkbookFactory.create(file)
  val etf = workbook.getSheet("Performance").getRow(0).getCell(0).getStringCellValue()
  val ticker = iSharesETFs(etf)

  sqlout(path, s"INSERT INTO etfs (id, issuer, name)\nVALUES\n  ('$ticker', ${sqlnull(etf)}, 'BlackRock, Inc.');")

  for (index <- 0 until workbook.getNumberOfSheets) {
    val sheet = workbook.getSheetAt(index)

    val sql = sheet.getSheetName match {
      case "Holdings" => holdingsSQL(sheet, ticker)
      case "Historical" => historicalSQL(sheet, ticker)
      case "Distributions" => distributionsSQL(sheet, ticker)
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
