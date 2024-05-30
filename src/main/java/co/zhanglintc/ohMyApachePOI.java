package co.zhanglintc;

import lombok.SneakyThrows;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;

// Press Shift twice to open the Search Everywhere dialog and type `show whitespaces`,
// then press Enter. You can now see whitespace characters in your code.
public class ohMyApachePOI {
    @SneakyThrows
    private static void writeExcel() {
        // 指定创建的excel文件名称
        BufferedOutputStream outputStream = new BufferedOutputStream(Files.newOutputStream(Paths.get("data.xls")));

        // 定义一个工作薄（所有要写入excel的数据，都将保存在workbook中）
        HSSFWorkbook workbook = new HSSFWorkbook();

        // 创建一个sheet
        HSSFSheet sheet = workbook.createSheet("my-sheet");

        CellStyle cellStyle = workbook.createCellStyle();
        DataFormat format = workbook.createDataFormat();
        cellStyle.setDataFormat(format.getFormat("@"));
        // cell.setCellStyle(cellStyle);

        for (int i = 0; i < 1000; i++) { // 假设最多1000行
            HSSFRow row1 = sheet.createRow(i);
            HSSFCell cell1 = row1.createCell(0); // 第一列
            cell1.setCellStyle(cellStyle);
        }

        // 开始写入数据流程，2大步：1、定位到单元格，2、写入数据；定位单元格，需要通过行、列配合指定。
        // step1: 先选择第几行（0表示第一行），下面表示在第6行
        HSSFRow row = sheet.createRow(0);
        // step2：选择第几列（0表示第一列），注意，这里的第几列，是在上面指定的row基础上，也就是第6行，第3列
        HSSFCell cell = row.createCell(1);
        // step3：设置单元格的数据（写入数据）
        cell.setCellValue("312");

        // 执行写入操作
        workbook.write(outputStream);
        workbook.close();
        outputStream.flush();
        outputStream.close();
    }

    @SneakyThrows
    private static void readExcel() {
        // 指定excel文件，创建缓存输入流
        BufferedInputStream inputStream = new BufferedInputStream(Files.newInputStream(Paths.get("data.xls")));

        // 直接传入输入流即可，此时excel就已经解析了
        HSSFWorkbook workbook = new HSSFWorkbook(inputStream);

        // 选择要处理的sheet名称
        HSSFSheet sheet = workbook.getSheet("my-sheet");
        HSSFCell head = sheet.getRow(0).getCell(1);
        System.out.println(head.getCellType());
    }

    public static void main(String[] args) {
        ohMyApachePOI.writeExcel();
        ohMyApachePOI.readExcel();
    }
}