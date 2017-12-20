package com.PTM;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;

public class ReadWord_copy {

    //Read table cell
    @Test
    public void readDocxTable() throws IOException, InvalidFormatException {

        FileInputStream fis = new FileInputStream("TC602 Vol 22.docx");
        XWPFDocument doc = new XWPFDocument(OPCPackage.open(fis));
        Iterator<IBodyElement> bodyElementIterator = doc.getBodyElementsIterator();
        while (bodyElementIterator.hasNext()) {

            IBodyElement element = bodyElementIterator.next();

            if ("TABLE".equalsIgnoreCase(element.getElementType().name())) {
                List<XWPFTable> tables = element.getBody().getTables();

                for (XWPFTable table : tables) {
                    for (XWPFTableRow row : table.getRows()) {

                        if (row.getCell(0).getText().contains("Version") ||
                                row.getCell(0).getText().contains("Name") ||
                                row.getCell(0).getText().contains("Number")) {
                            break;
                        }
                        System.out.println(row.getCell(0).getText());
                    }
                }
            }
        }
    }
}



