import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by prajena on 4/13/17.
 */
public class DiffExcel {
    private final String leftPath;
    private HSSFWorkbook leftwb;
    private HSSFWorkbook rightwb;
    private HSSFWorkbook resultwb;
    private String[] sheetNames;

    public DiffExcel(String leftPath, String rightPath) {
        this.leftPath = leftPath;
        try {
            leftwb = new HSSFWorkbook(new FileInputStream(new File(leftPath)));
            rightwb = new HSSFWorkbook(new FileInputStream(new File(rightPath)));
            resultwb = new HSSFWorkbook();
        } catch (IOException e) {
            System.out.println("Exception while Reading file : " + e);
        }
    }

    public DiffExcel(String leftPath, String rightPath, String... sheetNames) {
        this(leftPath, rightPath);
        this.sheetNames = sheetNames;
    }

    public void doDiff() throws Exception {
        // If sheet name given then compare those only
        if (sheetNames != null) {
            for (String sheetName : sheetNames) {
                HSSFSheet lSheet = leftwb.getSheet(sheetName);
                HSSFSheet rSheet = rightwb.getSheet(sheetName);
                validateSheet(lSheet, rSheet, sheetName);
                resultwb.createSheet(sheetName);
                diffSheet(lSheet, rSheet);
            }
        } else {// else compare for all sheets
            int leftSheetCount = leftwb.getNumberOfSheets();
            int rightSheetCount = rightwb.getNumberOfSheets();
            int lSC = 0;
            outer:
            while (lSC < leftSheetCount) {
                HSSFSheet lSheet = leftwb.getSheetAt(lSC);
                int rSC = 0;
                while (rSC < rightSheetCount) {
                    HSSFSheet rSheet = rightwb.getSheetAt(rSC);
                    if (lSheet.getSheetName().equalsIgnoreCase(rSheet.getSheetName())) {
                        resultwb.createSheet(lSheet.getSheetName());
                        diffSheet(lSheet, rSheet);
                        lSC++;
                        continue outer;
                    }
                    rSC++;
                }
                // If inner loop completed
                if (rSC == rightSheetCount) {
                    resultwb.createSheet(lSheet.getSheetName());
                    diffSheet(lSheet, null);
                }
                lSC++;
            }
        }
        writeReport();
    }

    private void writeReport() throws IOException {
        FileOutputStream fileOutputStream = new FileOutputStream(new File(leftPath + "_report.xls"));
        resultwb.write(fileOutputStream);
        fileOutputStream.close();
        System.out.println("Diff Completed");
    }

    private void diffSheet(HSSFSheet lSheet, HSSFSheet rSheet) {
        int maxRC = Math.max(lSheet == null ? 0 : lSheet.getPhysicalNumberOfRows(), rSheet == null ? 0 : rSheet.getPhysicalNumberOfRows());
        for (int rc = 0; rc < maxRC; rc++) {
            String resultSheetName = (lSheet == null ? rSheet.getSheetName() : lSheet.getSheetName());
            HSSFRow row = resultwb.getSheet(resultSheetName).createRow(rc);
            generateRowDiff(lSheet == null ? null : lSheet.getRow(rc), rSheet == null ? null : rSheet.getRow(rc), row);
        }
    }

    private String getCellValue(HSSFRow row, int cellIndex) {
        if (row == null)
            return "";
        HSSFCell cell = row.getCell(cellIndex);
        if (cell == null)
            return "";
        int cellType = cell.getCellType();
        String val = "";
        switch (cellType) {
            case Cell.CELL_TYPE_NUMERIC:
                val = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                val = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA:
                val = cell.getCellFormula().toString();
                break;
            default:
                val = cell.getStringCellValue();
        }
        return val;
    }

    private void generateRowDiff(HSSFRow lRow, HSSFRow rRow, HSSFRow row) {
        int maxCC = Math.max((lRow == null ? 0 : lRow.getLastCellNum()), (rRow == null ? 0 : rRow.getLastCellNum()));
        for (int cc = 0; cc < maxCC; cc++) {
            String lVal = getCellValue(lRow, cc);
            String rVal = getCellValue(rRow, cc);
            HSSFCell cell = row.createCell(cc);

            HSSFCellStyle cellStyle = getResultCellStyle(lRow, rRow, cc);
            if (cellStyle != null) {
                HSSFCellStyle cellStyle1 = resultwb.createCellStyle();
                cellStyle1.cloneStyleFrom(cellStyle);
                cell.setCellStyle(cellStyle1);
            }
            if (!lVal.equalsIgnoreCase(rVal)) {

                // Background Yellow for diff cell
                HSSFCellStyle cellStyle2 = resultwb.createCellStyle();
                cellStyle2.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                cellStyle2.setFillForegroundColor(HSSFColor.LIGHT_YELLOW.index);
                cell.setCellStyle(cellStyle2);
                cell.setCellValue(lVal);

                HSSFFont font = resultwb.createFont();
                font.setStrikeout(true);
                font.setItalic(true);
                font.setColor(HSSFColor.BLUE.index);

                HSSFFont font1 = resultwb.createFont();
                font1.setColor(HSSFColor.RED.index);

                HSSFRichTextString richTextString = new HSSFRichTextString(lVal + " " + rVal);
                richTextString.applyFont(0, lVal.length(), font);
                richTextString.applyFont(lVal.length(), richTextString.length(), font1);

                cell.setCellValue(richTextString);
            } else {//if (row.getRowNum() == 0)
                cell.setCellValue(lVal);
            }
        }
    }

    private HSSFCellStyle getResultCellStyle(HSSFRow lRow, HSSFRow rRow, int cc) {
        return (lRow != null && lRow.getCell(cc) != null ? lRow.getCell(cc).getCellStyle() : (rRow != null && rRow.getCell(cc) != null ? rRow.getCell(cc).getCellStyle() : null));
    }

    private void copyRow(HSSFRow sRow, HSSFRow tRow) {

    }

    private void validateSheet(HSSFSheet lSheet, HSSFSheet rSheet, String sheetName) {
        if (lSheet == null && rSheet == null) {
            System.out.println("Invalid sheet name : " + sheetName);
            throw new IllegalArgumentException("Invalid sheet name : " + sheetName);
        } else if (lSheet == null) {
            System.out.println(sheetName + " sheet not found in left file");
            throw new IllegalArgumentException(sheetName + " sheet not found in left file");
        } else if (rSheet == null) {
            System.out.println(sheetName + " sheet not found in left file");
            throw new IllegalArgumentException(sheetName + " sheet not found in left file");
        }
    }

}
