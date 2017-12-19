package com.PTM;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
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
    public void readDocxTable() throws IOException, InvalidFormatException {

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
//===========================================================================================================//read each cell per row
//        XWPFDocument doc = new XWPFDocument(fis);
//        List<XWPFTable>  tables = doc.getTables();
//
//        for ( XWPFTable table : tables )
//        {
//            for ( XWPFTableRow row : table.getRows() )
//            {
////                System.out.println(row.getCell(1).getText());
//                System.out.println("A row");
//
//                for (XWPFTableCell cell : row.getTableCells()) {
//                    System.out.println("=============================== ");
//                    System.out.println("cell data is: "+cell.getText());
////                    String sFieldValue = cell.getText();
////                    if (sFieldValue.matches("Whatever you want to match with the string") || sFieldValue.matches("Approved")) {
////                        System.out.println("The match as per the Document is True");
////                    }
//                }
//                System.out.println(" ");
//            }
//        }
//===========================================================================================================
//        XWPFDocument doc = new XWPFDocument(fis);
//        List<XWPFParagraph> paragraphs = doc.getParagraphs();
//
//        for (XWPFParagraph p : paragraphs) {
//            if (p.getText().contains("Variable data")) {
//                System.out.println("=============================== ");
////                System.out.println(p.);
//
//            }
//            else break;
//        }
//===========================================================================================================
//        XWPFDocument doc = new XWPFDocument(fis);
//        Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
//        while (iter.hasNext()) {
//            IBodyElement elem = iter.next();
//            if (elem instanceof XWPFParagraph) {
//                if (((XWPFParagraph)elem).getText().contains("Variable data"));
//                {
////                    if (elem instanceof XWPFTable){
//                        System.out.println(((XWPFTable) elem).getText());
//
////                    }
//                }
//            }
//            else if (elem instanceof XWPFTable) {
////                System.out.println(((XWPFTable) elem).getText());
//            }


        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
        Iterator<IBodyElement> bodyElementIterator = doc.getBodyElementsIterator();
        while (bodyElementIterator.hasNext()) {
            IBodyElement element = bodyElementIterator.next();
//System.out.println("++++++++++++ "+element.getElementType().name());
            if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                List <XWPFTable> tableList = element.getBody().getTables();
                //System.out.println("++++++++++++ "+element.getElementType().name());
//                for (XWPFTable table : tableList) {
//
//                    System.out.println("Total Number of Rows of Table:" + table.getNumberOfRows());
//                    for (int i = 0; i < table.getRows().size(); i++) {
//
//                        for (int j = 0; j < table.getRow(i).getTableCells().size(); j++) {
//                            System.out.println(table.getRow(i).getCell(j).getText());
//                        }
//                    }
//                }
            }
        }
        }
//===========================================================================================================
        //HWPFDocument
//        HWPFDocument doc = new HWPFDocument(new FileInputStream("TC602 Vol 22.docx"));
//        System.out.println(doc.getText());
//
//        System.out.println("Process Completed Successfully");
    }

//    @Test
//    public void searchText() throws IOException {
//
//        FileInputStream input_document = new FileInputStream(new File("test_document.doc"));
//        /* Create Word Extractor object to extract content of word document*/
//        WordExtractor my_word = new WordExtractor(input_document);
//
//    }

