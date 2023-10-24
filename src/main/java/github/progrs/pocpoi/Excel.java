package github.progrs.pocpoi;

import org.apache.poi.xssf.usermodel.*;

import java.io.*;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.springframework.boot.SpringApplication;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;
import java.util.TreeMap;

public class Excel {


    public static void main(String[] args) throws IOException {

        SpringApplication.run(PocpoiApplication.class, args);
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Sheet 1");

        XSSFRow row;


        Map<String, Object[]> infos = new TreeMap<String, Object[]>();
        //Header of my sheet
        infos.put("0", new Object[]{"Name", "Age", "Course"});
        //Datas of my sheet
        infos.put("1", new Object[]{"Luis", "36", "Apache POI - Java"});
        infos.put("2", new Object[]{"Fernando", "37", "Python Django"});
        infos.put("3", new Object[]{"Davi", "21", "Angular 14"});
        infos.put("4", new Object[]{"Vitor", "30", "Laravel"});

        for(Map.Entry<String, Object[]> entry : infos.entrySet()){

            String key = entry.getKey();
            Object[] data = entry.getValue();

            row = sheet.createRow(Integer.parseInt(key));

            int cellIndex = 0;
            for(Object obj: data){
                XSSFCell cell = row.createCell(cellIndex++);
                cell.setCellValue((String) obj);

            }

            System.out.println(entry.getKey());
            System.out.println(entry.getValue()[0]);
            System.out.println(entry.getValue()[1]);
            System.out.println(entry.getValue()[2]);

        }

        FileOutputStream out = new FileOutputStream( new File("workbook.xlsx"));

        workbook.write(out);

        out.close();

        System.out.println("Created");


    }

}
