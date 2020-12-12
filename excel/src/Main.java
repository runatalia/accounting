import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;


//25036 ячеек для проверки

public class Main {
    public static void main(String[] args) throws IOException {

        try (FileInputStream file = new FileInputStream(new File("ЧЭРЗЮ-УР2020.xlsx"));
             FileOutputStream outputStream = new FileOutputStream("out.xls");
             Workbook workbook = new XSSFWorkbook(file);) {
            int quantityLine = 4905;  //кол-в строк
            Sheet sheet = workbook.getSheetAt(0);
            ChangeColumn13 deleteErrorAnswer = new ChangeColumn13(sheet, quantityLine);
            ArrayList<Integer> errorColamn13 = new ArrayList<>(deleteErrorAnswer.changeErrorAnswer());// массив индексов,где задваивались значения

            ChangeSaldoEnd changeSaldoEnd = new ChangeSaldoEnd(sheet, errorColamn13, quantityLine);
            changeSaldoEnd.saldoMinus1();  //исправляем сальдо при -1
            changeSaldoEnd.saldo0In();      //исправляем остальное сальдо
            changeSaldoEnd.saldo0Out();


            DeleteEmptyLine deleteEmptyLine = new DeleteEmptyLine(sheet, quantityLine);
            try {
                deleteEmptyLine.deleteLine();
            } catch (NullPointerException e) {
                System.out.print("Строки закончились");
            }


            workbook.write(outputStream);

        } catch (IOException e) {
            System.out.print(e.getStackTrace());
        }

    }
}