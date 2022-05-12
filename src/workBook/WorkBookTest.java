package workBook;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Calendar;
import java.util.Date;

public class WorkBookTest {
    public static void main(String[] args) {
        Workbook workBook = new HSSFWorkbook();//HSSFWorkbook--xls;XSSFWorkbook--xlsx
        CreationHelper creationHelper = workBook.getCreationHelper();
        Sheet sheet1 = workBook.createSheet("first sheet");
        Sheet sheet2 = workBook.createSheet("second sheet");
        String sheetName = WorkbookUtil.createSafeSheetName("[O'Briesn's sales*?]");//创建一个安全的名字，用''替代无效字符
        System.out.println(sheetName);
        Sheet sheet3 = workBook.createSheet(sheetName);
        //创建普通Excel表格数据
        Row row = sheet1.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue(1);
        row.createCell(1).setCellValue(1.3);
        row.createCell(2).setCellValue(creationHelper.createRichTextString("this is a string"));
        row.createCell(3).setCellValue(true);
        //创建时间类型数据
        row = sheet2.createRow(0);
        row.createCell(0).setCellValue(new Date());
        CellStyle cellStyle = workBook.createCellStyle();
        cellStyle.setDataFormat(
                creationHelper.createDataFormat().getFormat("yyyy/m/d")
        );
        cell = row.createCell(1);
        cell.setCellValue(new Date());
        cell.setCellStyle(cellStyle);
        cell = row.createCell(2);
        cell.setCellValue(Calendar.getInstance());
        cell.setCellStyle(cellStyle);

        try (OutputStream out = new FileOutputStream("D:\\lrjob\\red\\workbook.xls")) {
            workBook.write(out);
            System.out.println("导出excel文件成功！");
        }catch(Exception e){
            e.printStackTrace();
        }
    }
}
