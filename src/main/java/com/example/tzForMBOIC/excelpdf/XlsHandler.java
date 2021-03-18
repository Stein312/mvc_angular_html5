package com.example.tzForMBOIC.excelpdf;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.servlet.view.document.AbstractExcelView;
import sun.rmi.runtime.Log;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;

public class XlsHandler extends AbstractExcelView {

/*
* Класс для генерации нового excel документа на основе загруженного.
* Используется HssfWorkbook так как XssfWorkbook выдает ClassNotPathException
* */
    @Override
    protected void buildExcelDocument(
            Map<String, Object> map,
            HSSFWorkbook hssfWorkbook,
            HttpServletRequest httpServletRequest,
            HttpServletResponse httpServletResponse) throws Exception {

        POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(SinglFileXLS.getInstance()));
        //FileInputStream file = new FileInputStream(new File("C:\\Users\\bovae\\Downloads\\apache-tomcat-8.5.64\\Лист Microsoft Excel (2).xlsx"));
        //Генерируеца новый файл
        HSSFWorkbook workbook=new HSSFWorkbook(fs);

        HSSFSheet outSheet=hssfWorkbook.createSheet("Лист 1");
        HSSFSheet sheet=workbook.getSheetAt(0);
        Iterator iterRow=sheet.rowIterator();
        httpServletResponse.setHeader("Content-Disposition", "attachment; filename=excelDocument.xls");

        int value;
        //Проходит по файлу и заносит те которые подходят в новый
        while (iterRow.hasNext()){
            Row row= (Row) iterRow.next();
            Row outRow=outSheet.createRow(row.getRowNum());
            Iterator iterCell=row.cellIterator();
            while (iterCell.hasNext()){
                Cell cell= (Cell) iterCell.next();
                //value=Integer.parseInt(cell.toString());
                value= (int) cell.getNumericCellValue();
                if(value%2==0)
                    outRow.createCell(cell.getColumnIndex()).setCellValue(value);
            }
        }

    }
}
