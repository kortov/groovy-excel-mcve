import com.jameskleeh.excel.ExcelBuilder
import org.apache.poi.ss.usermodel.BorderStyle
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.awt.*

class ExcelBookFormula {
  private XSSFWorkbook workbook
  private static final String DARK_BLUE_COLOR = "#2f75b5"
  File file = new File('test.xlsx')

  private Map headerStyle = [
          border           : BorderStyle.MEDIUM,
          foregroundColor  : DARK_BLUE_COLOR,
          font             : [
                  bold : true,
                  color: Color.WHITE
          ],
          alignment        : "center",
          verticalAlignment: "center",
          wrapped          : true
  ]


  void write() {
    workbook = ExcelBuilder.build {
      Sheet sh = sheet("Sheet") {
        columns(height: 30F) {
          column("Имя", "name", headerStyle)
          skipCells(30)
          column("Имя2", "name2", headerStyle)

        }
        skipRows(1)
        row {
            formula {
              "${exactCell("name2", 2).anchor()}"
            }
        }
      }
    }
    workbook.write(new FileOutputStream(file))
  }
}
