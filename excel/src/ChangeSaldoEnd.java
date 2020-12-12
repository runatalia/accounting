import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;

public class ChangeSaldoEnd {
    private Cell cell;
    private Row row;
    private Sheet sheet;


    private int quantityLine;

    private ArrayList<String> copySaldoOut = new ArrayList<>();
    private ArrayList<Integer> errorColamn13;

    ChangeSaldoEnd(Sheet sheet, ArrayList<Integer> errorColamn13, int quantityLine) {
        this.errorColamn13 = new ArrayList<>(errorColamn13);
        this.sheet = sheet;
        this.quantityLine = quantityLine;
    }

    public Sheet getSheet() {
        return sheet;
    }

    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    public int getQuantityLine() {
        return quantityLine;
    }

    public void setQuantityLine(int quantityLine) {
        this.quantityLine = quantityLine;
    }

    public ArrayList<Integer> getErrorColamn13() {
        return errorColamn13;
    }

    public void setErrorColamn13(ArrayList<Integer> errorColamn13) {
        this.errorColamn13 = errorColamn13;
    }

    public ArrayList<String> getCopySaldoOut() {
        return copySaldoOut;
    }

    public void setCopySaldoOut(ArrayList<String> copySaldoOut) {
        this.copySaldoOut = copySaldoOut;
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

    public void saldoMinus1() {
        System.out.println("Исправляем выбыло при сальдо  -1");
        for (int i = 0; i < errorColamn13.size() - 1; i++) {
            copySaldoOut.clear();
            row = sheet.getRow(errorColamn13.get(i) + 1);
            if (row.getCell(5).getStringCellValue().equals("0,000") && row.getCell(9).getStringCellValue().equals("1,000")
                    && !row.getCell(10).getStringCellValue().equals("") && !row.getCell(11).getStringCellValue().equals("")) {
                for (int j = 9; j < 13; j++) {
                    cell = row.getCell(j);
                    copySaldoOut.add(cell.getStringCellValue());


                }
                row = sheet.getRow(errorColamn13.get(i));
                for (int j = 9, k = 0; j < 13; j++, k++) {
                    cell = row.getCell(j);
                    cell.setCellValue(copySaldoOut.get(k));
                }

            } else System.out.println(" просмотри строку " + (errorColamn13.get(i) + 1) + " она пропущенна");
        }
        System.out.println("Закончено исправление выбыло при сальдо  -1");
        System.out.println();
        System.out.println("--------------------------------------------------------------------------");

    }

    public void saldo0In() {
        System.out.println("Исправляем выбыло при сальдо  0");
        for (int i = 7; i < quantityLine; i++) {
            copySaldoOut.clear();
            row = sheet.getRow(i + 1);
            if (row.getCell(5).getStringCellValue().equals("0,000") && row.getCell(9).getStringCellValue().equals("1,000")
                    && !row.getCell(10).getStringCellValue().equals("") && !row.getCell(11).getStringCellValue().equals("")) {
                for (int j = 9; j < 13; j++) {
                    cell = row.getCell(j);
                    copySaldoOut.add(cell.getStringCellValue());
                    //    cell.setCellValue("");

                }
                row = sheet.getRow(i);
                if (row.getCell(5).getStringCellValue().equals("1,000") && (row.getCell(9).getStringCellValue().equals("0,000") || row.getCell(9).getStringCellValue().equals(""))
                        && row.getCell(10).getStringCellValue().equals("") && row.getCell(11).getStringCellValue().equals("")) {
                    for (int j = 9, k = 0; j < 13; j++, k++) {
                        cell = row.getCell(j);
                        cell.setCellValue(copySaldoOut.get(k));
                    }
                    i += 1;
                }


            }
        }
        System.out.println("Закончено исправление выбыло  при сальдо  0");
        System.out.println();
        System.out.println("--------------------------------------------------------------------------");
    }

    public void saldo0Out() {
        System.out.println("Исправляем прибыло при сальдо  0");
        for (int i = 7; i < quantityLine; i++) {
            copySaldoOut.clear();
            row = sheet.getRow(i + 1);
            if (row.getCell(5).getStringCellValue().equals("1,000") && (row.getCell(9).getStringCellValue().equals("") || row.getCell(9).getStringCellValue().equals("0,000"))
                    && row.getCell(10).getStringCellValue().equals("") && row.getCell(11).getStringCellValue().equals("")
                    && !row.getCell(6).getStringCellValue().equals("")) {
                for (int j = 5; j < 9; j++) {
                    cell = row.getCell(j);
                    copySaldoOut.add(cell.getStringCellValue());

                }
                row = sheet.getRow(i);
                if ((row.getCell(5).getStringCellValue().equals("0,000") || row.getCell(5).getStringCellValue().equals("")) && row.getCell(9).getStringCellValue().equals("1,000")
                        && !row.getCell(10).getStringCellValue().equals("") && !row.getCell(11).getStringCellValue().equals("")) {
                    for (int j = 5, k = 0; j < 9; j++, k++) {
                        cell = row.getCell(j);
                        cell.setCellValue(copySaldoOut.get(k));
                    }
                    i += 1;
                }


            }
        }
        System.out.println("Закончено исправление прибыло при сальдо  0");
        System.out.println();
        System.out.println("--------------------------------------------------------------------------");
    }
}
