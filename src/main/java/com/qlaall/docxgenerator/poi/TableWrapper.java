package com.qlaall.docxgenerator.poi;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.xmlbeans.XmlCursor;

public class TableWrapper {
    private XWPFTable table;
    private Integer rowsCount;
    private Integer columnCount;

    public TableWrapper(XWPFTable table) {
        this.table = table;
        updateInfo();
    }
    public static TableWrapper insertAfter(XWPFParagraph paragraph){
        XmlCursor xmlCursor = paragraph.getCTP().newCursor();
        xmlCursor.toNextSibling();
        XWPFTable xwpfTable = paragraph.getDocument().insertNewTbl(xmlCursor);
        return new TableWrapper(xwpfTable);
    }
    public static TableWrapper insertBefore(XWPFParagraph paragraph){
        XmlCursor xmlCursor = paragraph.getCTP().newCursor();
        xmlCursor.toPrevSibling();
        XWPFTable xwpfTable = paragraph.getDocument().insertNewTbl(xmlCursor);
        return new TableWrapper(xwpfTable);
    }
    private void updateInfo(){
        this.rowsCount=this.table.getRows().size();
        this.columnCount=this.table.getRow(this.rowsCount-1).getTableCells().size();
    }
}
