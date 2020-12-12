import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;



public class DeleteEmptyLine {
    private Cell cell;
    private Row row;
    private Sheet sheet;
    private int quantityLine;


    DeleteEmptyLine(Sheet sheet, int quantityLine) {
        this.sheet = sheet;
        this.quantityLine = quantityLine;
    }
    public int getQuantityPg() {
        return quantityLine;
    }

    public void setQuantityPg(int quantityPg) {
        this.quantityLine = quantityPg;
    }
    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public Row getRow() {
        return row;
    }

    public void setRow(Row row) {
        this.row = row;
    }

    public Cell getCell() {
        return cell;
    }

    public void setCell(Cell cell) {
        this.cell = cell;
    }

    public void deleteLine() throws NullPointerException{
        System.out.println("Убираем лишние строки");

        for (int i = 7; i < quantityLine; i++) {
            row = sheet.getRow(i);

            if(row.getCell(13).getStringCellValue().equals("1,000")||row.getCell(13).getStringCellValue().equals("0,000")){

            }
              else{  removeRow(sheet, i);
                System.out.print(i + " ");
            }
        }
    }

    public static void removeRow(Sheet sheet, int rowIndex) {
        int lastRowNum = sheet.getLastRowNum();
        if (rowIndex >= 0 && rowIndex < lastRowNum) {
            sheet.shiftRows(rowIndex + 1, lastRowNum, -1);
        }
        if (rowIndex == lastRowNum) {
            Row removingRow = sheet.getRow(rowIndex);
            if (removingRow != null) {
                sheet.removeRow(removingRow);
            }
        }
    }

}
