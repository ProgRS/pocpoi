package github.progrs.pocpoi;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class ReadExcel {

    public static void main(String[] args) throws IOException {
        FileInputStream input = new FileInputStream(new File("workbook.xlsx"));

        XSSFWorkbook workbook = new XSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = (Iterator<Row>) sheet.iterator();

        while (iterator.hasNext()){
            Row currentRow =iterator.next();
            Iterator<Cell> cell = currentRow.iterator();
            while (cell.hasNext()){
                Cell currentCell = cell.next();

                if(currentCell.getCellType() == CellType.STRING){
                    System.out.println(currentCell.getStringCellValue());
                }

                if(currentCell.getCellType() == CellType.NUMERIC){
                    System.out.println(currentCell.getNumericCellValue() + "");
                }
            }
        }
    }

}
