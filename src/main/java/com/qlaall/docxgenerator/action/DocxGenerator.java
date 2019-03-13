package com.qlaall.docxgenerator.action;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;
import java.util.function.Consumer;

/**
 * @author: qlaall
 * @Date:2018/9/25
 * @Time:20:56
 */
public class DocxGenerator {
    private static final Logger LOGGER = LoggerFactory.getLogger(DocxGenerator.class);

    /**
     * 正式生成时，首先对段落进行填充，也就是整段只有标记字符，匹配到后，由paragraphHandler进行处理
     * 再进行表格的填充
     * 表格填充完毕后，再进行文字的处理
     * @param paragraphHandlerMap
     * @param tableCellHandlerMap
     * @param document
     * @return
     * @throws IOException
     */
    public static byte[] fillDocx(Map<String, BiConsumer<XWPFDocument, XWPFParagraph>> paragraphHandlerMap,
                           Map<String, Consumer<XWPFTableCell>> tableCellHandlerMap,
                           XWPFDocument document) throws IOException {
        XWPFDocument docx = document;
        /**
         * 初始化游标位置为0
         * init cursor position
         */
        int paragraphCursor = 0;
        /**
         * 继续处理标志
         * continue flag
         */
        boolean proccessing = true;
        List<XWPFParagraph> paragraphs = docx.getParagraphs();
        while (proccessing) {
            for (int i = paragraphCursor; i < paragraphs.size(); i++) {
                XWPFParagraph paragraph = paragraphs.get(i);
                if (paragraphHandlerMap.get(paragraph.getText()) != null) {
                    LOGGER.debug("dealing paragraph with {}", paragraph.getText());
                    paragraphCursor = paragraphs.indexOf(paragraph);
                    paragraphHandlerMap.get(paragraph.getText()).accept(docx, paragraph);
                    paragraphCursor++;
                    break;
                }
                if (i == paragraphs.size() - 1) {
                    proccessing = false;
                }
            }
        }
        final List<XWPFTable> tables = docx.getTables();
        for (XWPFTable table : tables) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    if (tableCellHandlerMap.get(cell.getText()) != null) {
                        LOGGER.debug("dealing tableCell with {}", cell.getText());
                        tableCellHandlerMap.get(cell.getText()).accept(cell);
                    }
                }
            }
        }
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        docx.enforceUpdateFields();
        docx.write(byteArrayOutputStream);
        LOGGER.debug("DocxGenerator finish");
        return byteArrayOutputStream.toByteArray();
    }
}
