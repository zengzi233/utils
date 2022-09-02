package utils.pdf.converter;

import com.itextpdf.io.exceptions.IOException;
import com.itextpdf.kernel.events.PdfDocumentEvent;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.element.Cell;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Locale;

/**
 * Created with IntelliJ IDEA.
 *
 * @author Randy.Z
 * Date: 2022/09/02
 * Time: 8:25 PM
 */
public class Cell2PdfUtil {

    public static void excelToPdf(String inFilePath, String outFilePath) {
        PdfFont pdfFont = PdfFontFactory.createFont("STSong-Light", "UniGB-UCS2-H", true);
        try (PdfDocument pdf = new PdfDocument(new PdfWriter(new FileOutputStream(outFilePath)));
             Document document = new Document(pdf, PageSize.A4.rotate());) {
            String type = "xls".equals(getSuffix(inFilePath).trim()) ?
                    "org.apache.poi.hssf.usermodel.HSSFWorkbook" : "org.apache.poi.xssf.usermodel.XSSFWorkbook";
            // 文件输入流读取文件
            InputStream in = new FileInputStream(inFilePath);
            // 反射创建workbook
            Class workbookClass = Class.forName(type);
            org.apache.poi.ss.usermodel.Workbook workbook = (org.apache.poi.ss.usermodel.Workbook) workbookClass.getConstructor(InputStream.class).newInstance(in);

            Sheet sheet = workbook.getSheetAt(0);
            int column = sheet.getRow(0).getLastCellNum();
            int row = sheet.getPhysicalNumberOfRows();

            Table table = new Table(column - sheet.getRow(0).getFirstCellNum());

            String str = null;
            for (int i = sheet.getFirstRowNum(); i < row; i++) {
                for (int j = sheet.getRow(0).getFirstCellNum(); j < column; j++) {
                    //得到excel单元格的内容
                    org.apache.poi.ss.usermodel.Cell cell = sheet.getRow(i).getCell(j);
                    if (cell.getCellType() == CellType.NUMERIC) {
                        str = (int) cell.getNumericCellValue() + "";
                    } else {
                        str = cell.getStringCellValue();
                    }
                    //创建pdf单元格对象，并往pdf单元格对象赋值。
                    Cell cells = new Cell().setFont(pdfFont).add(new Paragraph(str));
                    //pdf单元格对象添加到table对象
                    table.addCell(cells);
                }
            }
            document.add(table);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static String getSuffix(String filePath) {
        int dotIndex = filePath.lastIndexOf(".");
        return filePath.substring(dotIndex + 1);
    }

}
