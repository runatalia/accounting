import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.*;
//работает только если 13 колонка стиль текст
//25036 ячеек для проверки

public class ChangeColumn13 {

    private Sheet sheet;
    private int quantityLine;

    ChangeColumn13(Sheet sheet, int quantityLine) {
        this.sheet = sheet;
        this.quantityLine = quantityLine;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public int getQuantityPg() {
        return quantityLine;
    }

    public void setQuantityPg(int quantityPg) {
        this.quantityLine = quantityPg;
    }

    public ArrayList<Integer> changeErrorAnswer() {
        Iterator<Row> iterator = sheet.iterator();
        int count = 0;
        ArrayList<Integer> arrayIndexRow = new ArrayList<>();
        Row row;
        while (iterator.hasNext()) {   // проход по всем страницам в столбце 13,если в ячейки неверное значение,то заменяем его на пустое
            Row nextRow = iterator.next();
            Cell cell = nextRow.getCell(13);
            switch (cell.getCellType()) {
                case STRING:
                    if (cell.getStringCellValue().equals("1,000-")) {
                        cell.setCellValue("");
                        row = sheet.getRow(cell.getRowIndex() - 1);
                        Cell cellFix = row.getCell(13);
                        cellFix.setCellValue("0,000");
                        arrayIndexRow.add(cellFix.getRowIndex());  // так как в текущей неверное значение,в предыдущей ячейке тоже меняем значение
                    }
                    break;
                case BOOLEAN:
                    break;
                case NUMERIC:
                    break;
            }
            count++;
            if (count == quantityLine) break;
        }
        return arrayIndexRow;
    }

}
