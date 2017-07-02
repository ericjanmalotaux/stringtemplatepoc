package nl.malotaux.eric.stp;

import io.vavr.collection.List;
import org.apache.commons.lang3.builder.ToStringBuilder;
import org.apache.commons.lang3.builder.ToStringStyle;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.stringtemplate.v4.ST;
import org.stringtemplate.v4.STGroupFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.util.Date;

/**
 * Hello world!
 */
public class App {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        ST template = new STGroupFile("nl/malotaux/eric/stp/service.stg").getInstanceOf("services");
        template.add("services", readServices("nl/malotaux/eric/stp/Voorbeeld.xlsx"));
        Files.write(Paths.get("target/services.html"), template.render().getBytes());
    }

    private static java.util.List<Service> readServices(String pad) throws IOException, InvalidFormatException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(ClassLoader.getSystemResource(pad).openStream())) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            return List.range(12, 27).map(sheet::getRow).map(row -> new Service(
                    row.getCell(0).getDateCellValue(),
                    row.getCell(1).getStringCellValue(),
                    row.getCell(2).getStringCellValue(),
                    row.getCell(3).getStringCellValue())).toJavaList();
        }
    }

    static class Service {
        Date datum;
        String leader;
        String parent;
        String member;

        public Service(Date datum, String leader, String parent, String member) {
            this.datum = datum;
            this.leader = leader;
            this.parent = parent;
            this.member = member;
        }

        public String getDatum() {
            return DateFormat.getDateInstance(DateFormat.MEDIUM).format(datum);
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
            return ToStringBuilder.reflectionToString(this, ToStringStyle.JSON_STYLE);
        }
    }
}
