package nl.malotaux.eric.stp;

import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws IOException {
        System.out.println("Hello World!");
        String pad = "/Users/ema21214/Downloads/Voorbeeld.xlsx";

        try (XSSFWorkbook workbook = new XSSFWorkbook((new FileInputStream(pad)))) {
            XSSFSheet first = workbook.getSheetAt(0);
            for (int row = 12; row < 28; row++) {
                XSSFRow currentRow = first.getRow(row);
                Service service = new Service(currentRow.getCell(0).getDateCellValue(),
                        currentRow.getCell(1).getStringCellValue(),
                        currentRow.getCell(2).getStringCellValue(),
                        currentRow.getCell(3).getStringCellValue());
                System.out.println(service);
            }
        }

    }

    static class Service {
        Date datum;
        String leader;
        String parent;
        String member;

        public Service(Date cell, String cell1, String cell2, String cell3) {
            this.datum = cell;
            this.leader = cell1;
            this.parent = cell2;
            this.member = cell3;
        }

        public Date getDatum() {
            return datum;
        }

        public String getLeader() {
            return leader;
        }

        public String getParent() {
            return parent;
        }

        public String getMember() {
            return member;
        }

        public String toString() {
            return ToStringBuilder.reflectionToString(this,  ToStringStyle.JSON_STYLE);
        }
    }
}
