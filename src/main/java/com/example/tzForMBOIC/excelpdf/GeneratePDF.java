package com.example.tzForMBOIC.excelpdf;


import com.lowagie.text.Document;
import com.lowagie.text.pdf.PdfPTable;
import com.lowagie.text.pdf.PdfWriter;
import com.spire.xls.Workbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.web.servlet.view.document.AbstractPdfView;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.Map;

/*
        Класс для генерации excel файла с рандомными числами.
 */
public class GeneratePDF extends AbstractPdfView {


        /**
         * Метод для генерации pdf документа для возможности просмотра изменений в excel файле.
         *
         *
         */



        @Override
        protected void buildPdfDocument(Map<String, Object> map,
                                        Document document,
                                        PdfWriter pdfWriter,
                                        HttpServletRequest httpServletRequest,
                                        HttpServletResponse httpServletResponse) throws Exception {
                HSSFWorkbook workbook=new HSSFWorkbook(new FileInputStream(SinglFileXLS.getInstance()));
                PdfPTable table=new PdfPTable(workbook.getSheetAt(0).getRow(0).getLastCellNum());

                HSSFSheet sheet=workbook.getSheetAt(0);
                Iterator iterRow=sheet.rowIterator();
                Integer value;
                while (iterRow.hasNext()){
                       Row row= (Row) iterRow.next();

                        Iterator iterCell=row.cellIterator();
                        while (iterCell.hasNext()){
                                Cell cell= (Cell) iterCell.next();
                                //value=Integer.parseInt(cell.toString());
                                value= (int) cell.getNumericCellValue();
                                if(value%2==0)
                                        table.addCell(value.toString());
                                else table.addCell(" ");
                        }
                }
                document.add(table);


        }
}
