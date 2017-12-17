package com.PTM;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;

import java.io.*;
import java.util.Iterator;
import java.util.List;

public class ReadWord {

    //Read and Write to word sample
    @Test
    public void readDocx() throws IOException {

//        XWPFDocument document = new XWPFDocument();
//
//        //Write the Document in file system
//        FileOutputStream out = new FileOutputStream(new File("createdocument.docx"));
//
//        //create table
//        XWPFTable table = document.createTable();
//
//        //create first row
//        XWPFTableRow tableRowOne = table.getRow(0);
//        tableRowOne.getCell(0).setText("col one, row one, header");
//        tableRowOne.addNewTableCell().setText("col two, row one, header");
//        tableRowOne.addNewTableCell().setText("col three, row one, header");
//
//        //create second row
//        XWPFTableRow tableRowTwo = table.createRow();
//        tableRowTwo.getCell(0).setText("col one, row two");
//        tableRowTwo.getCell(1).setText("col two, row two");
//        tableRowTwo.getCell(2).setText("col three, row two");
//
//        document.write(out);
//        out.close();
//        System.out.println("createdocument.docx written successully");

        //To read the whole doc
//        XWPFDocument docx = new XWPFDocument(new FileInputStream("createdocument.docx"));
//
//        //using XWPFWordExtractor Class
//        XWPFWordExtractor we = new XWPFWordExtractor(docx);
//        System.out.println(we.getText());

    }

    //Read table cell
    @Test
    public void readDocxTable() throws IOException {

        FileInputStream fis = new FileInputStream("TC602 Vol 22.docx");
//        XWPFDocument xdoc=new XWPFDocument(new FileInputStream("createdocument.docx"));
//        Iterator<IBodyElement> bodyElementIterator = xdoc.getBodyElementsIterator();
//
//        while(bodyElementIterator.hasNext()) {
//
//            IBodyElement element = bodyElementIterator.next();
//
//            if("TABLE".equalsIgnoreCase(element.getElementType().name())) {
//                List<XWPFTable> tableList =  element.getBody().getTables();
//                for (XWPFTable table: tableList){
//                    System.out.println("Total Number of Rows of Table:"+table.getNumberOfRows());
//                    System.out.println(table.getText());
//                    String cell = table.getRow(1).getCell(0).getText();
//                    System.out.println(cell);
//                }
//            }
//        }
        XWPFDocument doc = new XWPFDocument(fis);
        List<XWPFTable>  tables = doc.getTables();
        for ( XWPFTable table : tables )
        {
            for ( XWPFTableRow row : table.getRows() )
            {
                System.out.println(row.getCell(1).getText());
            }
        }
    }
}
