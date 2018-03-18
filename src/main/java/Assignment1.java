package com.mycompany.assignment1try;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import org.apache.poi.ss.usermodel.DataFormatter;
import java.io.Writer;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    public static void read() throws IOException {

        Writer writer = null;
        boolean line = true;

        try {

            DataFormatter df = new DataFormatter();
            FileInputStream excelFile = new FileInputStream(new File("F:/folder uum/SEM5/RealTime/PracticumStudentSupervisorList.xlsx"));
            Workbook workbook = new XSSFWorkbook(excelFile);
            Sheet datatypeSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = datatypeSheet.iterator();
            File file = new File("/Users/Alif Haikal/cuba.wiki/Home.md");
            writer = new BufferedWriter(new FileWriter(file));

            while (iterator.hasNext()) {

                Row currentRow = iterator.next();
                Iterator<Cell> cellIterator = currentRow.iterator();

                while (cellIterator.hasNext()) {

                    Cell currentCell = cellIterator.next();
                    String val = df.formatCellValue(currentCell);

                    System.out.print(val + "|");

                    writer.write(val + "|");

                }
                System.out.println();
                writer.write("\n");
                if (line == true) {
                    writer.write("---|---|---|---|\n");
                    line = false;
                }

            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        try {
            if (writer != null) {
                writer.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void gitbash() throws IOException {
        String[] command = {"C:/Program Files/Git/git-bash.exe",};

        Runtime.getRuntime().exec(command);

    }

    public static void main(String[] args) throws IOException {

        read();
        gitbash();

    }
}
