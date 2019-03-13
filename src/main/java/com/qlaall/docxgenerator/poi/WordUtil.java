package com.qlaall.docxgenerator.poi;


import com.qlaall.docxgenerator.poi.styles.POIColor;
import com.qlaall.docxgenerator.poi.styles.Style;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;

import java.io.IOException;
import java.io.InputStream;
import java.math.BigInteger;

/**
 * @author qilei
 * @since 2018-03-20 09:56
 */
public class WordUtil {

    public static XWPFParagraph insertParagraphAfter(XWPFParagraph currentParagraph) {
        XmlCursor xmlCursor = currentParagraph.getCTP().newCursor();
        xmlCursor.toNextSibling();
        POIXMLDocumentPart part = currentParagraph.getDocument().getPart();
        XWPFParagraph xwpfParagraph = currentParagraph.getDocument().insertNewParagraph(xmlCursor);
        return xwpfParagraph;
    }


    public static XWPFTable insertTableAfter(XWPFParagraph currentParagraph) {
        XmlCursor xmlCursor = currentParagraph.getCTP().newCursor();
        xmlCursor.toNextSibling();
        XWPFTable xwpfTable = currentParagraph.getDocument().insertNewTbl(xmlCursor);
        return xwpfTable;
    }

    public static XWPFParagraph insertParagraphAfter(XWPFDocument docx,XWPFTable xwpfTable) {
        XmlCursor xmlCursor = xwpfTable.getCTTbl().newCursor();
        xmlCursor.toNextSibling();
        XWPFParagraph xwpfParagraph = docx.insertNewParagraph(xmlCursor);
        return xwpfParagraph;
    }

    public static XWPFTable insertTableAfter(XWPFDocument docx, XWPFTable xwpfTable) {
        XmlCursor xmlCursor = xwpfTable.getCTTbl().newCursor();
        xmlCursor.toNextSibling();
        return docx.insertNewTbl(xmlCursor);
    }

    /**
     * 移除paragraph中所有的run，不保留run
     *
     * @param paragraph
     */
    public static void clearMe(XWPFParagraph paragraph) {
        for (int i = paragraph.getRuns().size(); i > 0; i--) {
            paragraph.removeRun(0);
        }
    }

}
