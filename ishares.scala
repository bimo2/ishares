import java.io.{File, PrintWriter}
import org.apache.poi.ss.usermodel.{Sheet, WorkbookFactory}

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

def holdingsSQL(sheet: Sheet): Unit = {
  println(s"parse holdings")
}

def historicalSQL(sheet: Sheet): Unit = {
  println(s"parse historical")
}

def distributionsSQL(sheet: Sheet): Unit = {
  println(s"parse distributions")
}

def sql(xlsx: File): Unit = {
  println(s"generate sql from: ${xlsx.getName}")

  val workbook = WorkbookFactory.create(xlsx)

  for (index <- 0 until workbook.getNumberOfSheets) {
    val sheet = workbook.getSheetAt(index)

    sheet.getSheetName match {
      case "Holdings" => holdingsSQL(sheet)
      case "Historical" => historicalSQL(sheet)
      case "Distributions" => distributionsSQL(sheet)
      case _ => // do nothing
    }
  }

  workbook.close()
}

def write(path: String, content: String): Unit = {
  val file = new File(path)
  val directory = file.getParentFile

  if (!directory.exists()) directory.mkdirs()

  file.createNewFile()

  val writer = new PrintWriter(file)

  try {
    writer.println(content)
  } finally {
    writer.close()
  }
}

@main def script(args: String*): Unit = {
  write("sql/schema.sql", schemaSQL)

  val xlsxPath = new File("xlsx")

  if (!xlsxPath.isDirectory) return

  val files = xlsxPath.listFiles.filter(_.isFile)

  files.foreach(sql)
}
