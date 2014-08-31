import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by chenlian on 14-8-29.
 */
public class Excel implements MicroSoftDocument {
    @Override
    public boolean toFile(String outFilePath) {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        HSSFRow row = null;
        HSSFCell cell = null;
        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        HSSFDataFormat dataFormat = wb.createDataFormat();
        HSSFFont font = wb.createFont();
        sheet.setDefaultColumnWidth(22);
        font.setFontName("微软雅黑");
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        style.setFont(font);
        for (short s = 0; s < 20; s++) {
            row = sheet.createRow(s);
            row.setHeightInPoints(18);

            if (s % 2 == 0) row.setHeightInPoints(22);
            for (int c = 0; c < 8; c++) {

                cell = row.createCell(c);


                if (s == 0) {
                    cell.setCellStyle(style);
                    cell.setCellValue("列的标题" + s + "-" + c);
                } else
                    cell.setCellValue("列的内容" + s + "-" + c);
            }
        }
        try {
            FileOutputStream outputStream = new FileOutputStream(outFilePath);
            wb.write(outputStream);
            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


        return false;
    }

    @Override
    public boolean openFile(String filePath) {
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
            for (int numS = 0; numS < workbook.getNumberOfSheets(); numS++) {
                HSSFSheet sheet = workbook.getSheetAt(numS);
                if (null == sheet) continue;
                for (int rowNum = 0; rowNum < sheet.getLastRowNum(); rowNum++) {
                    HSSFRow row = sheet.getRow(rowNum);
                    if (null == row) continue;
                    for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
                        HSSFCell cell = row.getCell(cellNum);

                        System.out.print(String.format("第%d行-第%d列的数据是%s", rowNum, cellNum, (cell.getStringCellValue())));

                    }
                    System.out.println();
                }

            }

            return true;
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }
}
