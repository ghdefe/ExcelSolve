package DB;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * db
 *
 * @author chunmiaoz
 */
public class DbAccountSolve {

    public static void main(String[] args) {
        final File file = new File("F:\\工作记录\\caiting\\文档管理\\数字财政项目 - TDSQL数据库账号清单.xlsx");
        DbAccountSolve dbAccountSolve = new DbAccountSolve();
        try {
            XSSFSheet sheet = dbAccountSolve.getSheet(file);
            dbAccountSolve.getLocation(sheet);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private XSSFSheet getSheet(File file) throws IOException {
        BufferedInputStream bufferedInputStream = new BufferedInputStream(new FileInputStream(file));
        final XSSFWorkbook sheets = new XSSFWorkbook(bufferedInputStream);
        return sheets.getSheetAt(0);
    }

    private void readAllLine(File file) {

        try (BufferedInputStream bufferedInputStream = new BufferedInputStream(new FileInputStream(file))) {

            final XSSFWorkbook sheets = new XSSFWorkbook(bufferedInputStream);
            final XSSFSheet sheetA = sheets.getSheetAt(0);
            final int lastRowNum = sheetA.getLastRowNum();
            for (int i = 0; i < lastRowNum; i++) {
                final XSSFRow row = sheetA.getRow(i);
                final XSSFCell cell = row.getCell(0);
                final String stringCellValue = cell.getStringCellValue();
                System.out.println("the row of " + i + " : " + stringCellValue);
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void getLocation(XSSFSheet sheet) {
        String[] locations = {""};
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 0; i < lastRowNum; i++) {
            XSSFRow row = sheet.getRow(i);
            XSSFCell cell = row.getCell(0);
            String str = cell.getStringCellValue();

            String reg = ".*潮州.*";
            Pattern compile = Pattern.compile(reg);
            Matcher matcher = compile.matcher(str);
            boolean b = matcher.find();
            String group = "";
            if (b) {
                group = matcher.group(0);
            }
            System.out.println(b + " : " + group);
        }

    }
}
