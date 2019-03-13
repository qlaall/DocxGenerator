package com.qlaall.docxgenerator.action;

import org.apache.poi.xwpf.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.function.BiConsumer;

/**
 * @author: qlaall
 * @Date:2018/9/25
 * @Time:20:56
 */
public class DocxGenerator<T> {
    private static final Logger LOGGER = LoggerFactory.getLogger(DocxGenerator.class);
    /**
     * Generate时使用的模板文件
     */
    private XWPFDocument templateDocument;
    public DocxGenerator(XWPFDocument templateDocument) {
        this.templateDocument = templateDocument;
    }

    /**
     *
     * @param paragraphHandleMap    段落handle，文档优先处理段落
     * @param tableCellHandleMap    单元格内容handle，优先级仅次于段落
     * @param dataModel 数据模型
     * @return
     * @throws IOException
     */
    public byte[] fillDocx(Map<String, BiConsumer<XWPFParagraph,T>> paragraphHandleMap,
                           Map<String, BiConsumer<XWPFTableCell,T>> tableCellHandleMap,
                           T dataModel) throws IOException {
        XWPFDocument docx = this.templateDocument;
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
                if (paragraphHandleMap.get(paragraph.getText()) != null) {
                    LOGGER.debug("dealing paragraph with {}", paragraph.getText());
                    paragraphCursor = paragraphs.indexOf(paragraph);
                    paragraphHandleMap.get(paragraph.getText()).accept(paragraph,dataModel);
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
                    if (tableCellHandleMap.get(cell.getText()) != null) {
                        LOGGER.debug("dealing tableCell with {}", cell.getText());
                        tableCellHandleMap.get(cell.getText()).accept(cell,dataModel);
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
