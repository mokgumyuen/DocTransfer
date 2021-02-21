package com.mokgumyuen.word4j.demo;


import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.io.*;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

public class DocTransferDemo {

    /**
     * 输入目录
     */
    private final static String INPUT_PATH = "D:\\wordTransfer";

    /**
     * 输出目录
     */
    private final static String OUTPUT_PATH = "D:\\wordTransferOutput";

    /**
     * 错误日志
     */
    private final static String LOG_PATH = "D:\\wordTransferError";

    public static void main(String[] args) {

        File file = new File(INPUT_PATH);
        if (!file.exists()) {
            throw new RuntimeException("文件夹不存在");
        }
        File newFilePath = new File(OUTPUT_PATH);
        if (!newFilePath.exists()) {
            newFilePath.mkdirs();
        }
        File errorPath = new File(LOG_PATH);
        if (!errorPath.exists()) {
            errorPath.mkdirs();
        }
        try (OutputStream errOs = new FileOutputStream(new File(LOG_PATH + "\\error.log"))) {
            File[] files = file.listFiles();
            for (File f : files) {
                String path = f.getAbsolutePath();
                String name = f.getName();
                System.out.println(name);
                String s;
                try {
                    s = deal2007(path);
                } catch (Exception e) {
                    write(errOs, e);
                    try {
                        s = deal2003(path);
                    } catch (Exception exception) {
                        write(errOs, exception);
                        continue;
                    }
                }
                int endIndex = name.lastIndexOf('.');
                String newName = name.substring(0, endIndex) + ".csv";
                try (OutputStream fos = new FileOutputStream(new File(OUTPUT_PATH + "\\" + newName))) {
                    fos.write(s.getBytes());
                    fos.flush();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            errOs.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static String deal2003(String path) throws Exception {
        try (InputStream is = new FileInputStream(path)) {
            HWPFDocument hwpf = new HWPFDocument(is);
            //遍历range范围内的table。
            TableIterator tableIterator = new TableIterator(hwpf.getRange());
            List<List<String>> tables = new ArrayList<>();
            StringBuffer buffer = new StringBuffer();
            while (tableIterator.hasNext()) {
                Table table = tableIterator.next();
                int rowNum = table.numRows();
                List<String> rows = new ArrayList<>();
                for (int j = 0; j < rowNum; j++) {
                    TableRow row = table.getRow(j);
                    int cellNum = row.numCells();
                    List<String> cells = new ArrayList<>();
                    for (int k = 0; k < cellNum; k++) {
                        TableCell cell = row.getCell(k);
                        //输出单元格的文本
                        String trim = cell.text().trim()
                                .replaceAll("\\r", " ")
                                .replaceAll("\\t", " ");
                        cells.add(trim);
                    }
                    rows.add(String.join(",", cells));
                }
                tables.add(rows);
            }
            tables.forEach(x -> x.forEach(y -> {
                buffer.append(y);
                buffer.append("\n");
            }));
            return buffer.toString();
        } catch (FileNotFoundException e) {
            throw new Exception("path:" + path, e);
        } catch (IOException e) {
            throw new Exception("path:" + path, e);
        }
    }

    public static String deal2007(String path) throws Exception {
        try (InputStream is = new FileInputStream(path)) {
            XWPFDocument ex = new XWPFDocument(is);
            List<XWPFTable> tables = ex.getTables();
            StringBuffer buffer = new StringBuffer();
            List<List<String>> textTables = new ArrayList<>();
            for (XWPFTable table : tables) {
                // 获取表格的行
                List<XWPFTableRow> rows = table.getRows();
                List<String> textRows = new ArrayList<>();
                for (XWPFTableRow row : rows) {
                    // 获取表格的每个单元格
                    List<XWPFTableCell> tableCells = row.getTableCells();
                    List<String> cells = new ArrayList<>();
                    for (XWPFTableCell cell : tableCells) {
                        // 获取单元格的内容
                        String text = cell.getText()
                                .replaceAll("\\r", " ")
                                .replaceAll("\\t", " ");
                        cells.add(text);
                    }
                    textRows.add(String.join(",", cells));
                }
                textTables.add(textRows);
            }
            textTables.forEach(x -> x.forEach(y -> {
                buffer.append(y);
                buffer.append("\n");
            }));
            return buffer.toString();
        } catch (FileNotFoundException e) {
            throw new Exception("path:" + path, e);
        } catch (IOException e) {
            throw new Exception("path:" + path, e);
        }

    }

    private static void write(OutputStream errOs, Exception e) {
        try {
            errOs.write((LocalDateTime.now()
                    + " message:" + e.getMessage() +
                    " cause:" + e.getCause().toString()
                    + "\n").getBytes());
            errOs.flush();
        } catch (IOException ioException) {
            ioException.printStackTrace();
        }
    }
}
