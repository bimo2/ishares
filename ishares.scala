import java.io.File
import org.apache.poi.ss.usermodel.WorkbookFactory

def sql(xlsx: File): Unit = {
  println(s"generate sql from: ${xlsx.getName}")

  val workbook = WorkbookFactory.create(xlsx)

  for (index <- 0 until workbook.getNumberOfSheets) {
    val sheet = workbook.getSheetAt(index)

    println(s"sheet: ${sheet.getSheetName}")
  }

  workbook.close()
}

@main def script(args: String*): Unit = {
  val xlsx = new File("xlsx")

  if (!xlsx.isDirectory) return

  val files = xlsx.listFiles.filter(_.isFile)

  files.foreach(sql)
}
