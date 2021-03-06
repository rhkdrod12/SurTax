/**
 *
 */
package excel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import utils.DataVo;
import utils.Path;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;


/**
 * @author Kim
 *
 */
public class ExcelIO {
    
    
    
    static public List<List<String>> ReadExcel(String path, int sheetIdx, int statRow, int startCell, int cellCount) {
        
        if (path != null && path != "") {
            File file = new File(path);
            if (file.exists() && file.isFile()) {
                try {
                    FileInputStream fileInput = new FileInputStream(file);
                    XSSFWorkbook workbook = new XSSFWorkbook(fileInput);
                    
                    int rowFirstIdx = statRow;
                    int cellFirstIdx = startCell;
                    
                    List<List<String>> list = new ArrayList<List<String>>();
                    
                    XSSFSheet sheet = workbook.getSheetAt(sheetIdx);
                    
                    int rowLastIdx = sheet.getLastRowNum() + 1;
                    for (int rowIdx = rowFirstIdx; rowIdx < rowLastIdx; rowIdx++) {
                        XSSFRow row = sheet.getRow(rowIdx);
                        if (row != null) {
                            List<String> cellList = new ArrayList<String>();
                            int cellLastIdx = cellCount; // row.getPhysicalNumberOfCells();
                            for (int cellIdx = cellFirstIdx; cellIdx < cellLastIdx; cellIdx++) {
                                XSSFCell cell = row.getCell(cellIdx);
                                if (cell == null) {
                                    cellList.add("");
                                } else {
                                    if (!cell.getCellType().equals(CellType.STRING)) {
                                        cell.setCellType(CellType.STRING);
                                    }
                                    cellList.add(cell.getStringCellValue());
                                }
                            }
                            list.add(cellList);
                        }
                    }
                    
                    return list;
                    
                } catch (IOException e) {
                    System.out.println("????????? ????????? ??????????????????. ????????? ??????????????????.");
                }
            } else {
                System.out.println("????????? ???????????????. ????????? ??????????????????.");
            }
        } else {
            System.out.println("????????? ???????????????. ?????? ??????????????????");
        }
        
        return null;
    }
    
    private static String basePath = Path.BASE_PATH;
    static XSSFWorkbook workbook;
    static int startRowIdx = 31;
    static int statColumnIdx = 2;
    static private String[] columnNames = {"????????????", "????????????", "????????????", "??????"};
    
    static public void WriteExcel2(Map<String, DataVo> data) {
        
        String path = basePath + "/??????/????????? ?????? ??????Test.xlsx";
        File file = new File(path);
        try {
            FileInputStream fileInput = new FileInputStream(file);
            workbook = new XSSFWorkbook(fileInput);
            XSSFSheet sheet = workbook.getSheetAt(0);
            sheet.setForceFormulaRecalculation(true);
            
            int rowIdx = startRowIdx;
            int cellIdx = statColumnIdx;
    
            // XSSFCellStyle cellStyle = sheet.getWorkbook().createCellStyle();
            // XSSFRow row = sheet.createRow(rowIdx);
            // XSSFCell cell = row.createCell(cellIdx);
            // cell.setCellValue("??????");
            
            for (Map.Entry<String, DataVo> val : data.entrySet()) {
                
                makeHeaderFormExcel(val.getKey(), sheet, workbook.createCellStyle(), rowIdx, cellIdx);
                rowIdx += 2;
                
                DataVo vo = val.getValue();
                
                List<String> list = new ArrayList<>(vo.getData().keySet());
                Collections.sort(list);
                
                int startRowIdx = rowIdx;
                for (int i = 0; i < list.size(); i++) {
                    makeBodyFormExcel(list.get(i), vo.getPartData(list.get(i)).getVal(), sheet, rowIdx++, cellIdx);
                }
                
                makeFooterFromExcel(startRowIdx, cellIdx, rowIdx++, cellIdx + columnNames.length, sheet);
                rowIdx++;
            }
            
            String curDate = new SimpleDateFormat("YYMMdd").format(new Date(System.currentTimeMillis()));
            String makeFilePath = basePath + "/????????????_" + curDate + ".xlsx";
            File makeFile = new File(makeFilePath);
            FileOutputStream fos = new FileOutputStream(makeFile);
            workbook.write(fos);
            workbook.close();
            fos.close();
            
            System.out.println("????????????");
            
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    //???????????? ???
    static private void makeHeaderFormExcel(String formName, XSSFSheet sheet, XSSFCellStyle cellStyle, int rowIdx, int cellIdx) {
        
        int beginCellIdx = cellIdx;
        
        XSSFRow row = null;
        XSSFCell cell = null;
        
        HeaderCellStyle(cellStyle);
        
        row = sheet.createRow(rowIdx); //sheet.getRow(rowIdx);
        
        //cell ?????? ??????
        for (int i = cellIdx + 1; i < cellIdx + 5; i++) {
            sheet.setColumnWidth(i, 4500);
        }
        
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx + 1, beginCellIdx, beginCellIdx));
        cell = row.createCell(beginCellIdx);
        cell.setCellStyle(cellStyle);
        cell.setCellValue("??????");
        
        sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, beginCellIdx + 1, beginCellIdx + 4));
        cell = row.createCell(beginCellIdx + 1);
        cell.setCellValue(formName);
        cell.setCellStyle(cellStyle);
        for (int i = beginCellIdx + 2; i <= beginCellIdx + 4; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(cellStyle);
        }
        
        row  = sheet.createRow(rowIdx + 1);
        cell = row.createCell(beginCellIdx);
        cell.setCellStyle(cellStyle);
        
        
        for (int i = 1; i <= columnNames.length; i++) {
            cell = row.createCell(beginCellIdx + i);
            cell.setCellStyle(cellStyle);
            cell.setCellValue(columnNames[i - 1]);
        }
    }
    //???????????? ???
    static private void makeBodyFormExcel(String month, Map<String, Long> data, XSSFSheet sheet, int rowIdx, int cellIdx) {
        
        XSSFRow row = sheet.createRow(rowIdx);
        XSSFCell cell;
        
        cell = row.createCell(cellIdx);
        cell.setCellValue(month);
        BodyCellStyleMonth(cell, sheet);    //?????? ????????? ????????? ??????????????? ?????? ????????? ??????!
        
        for (int i = 0; i < columnNames.length; i++) {
            String columnName = columnNames[i];
            cell = row.createCell(cellIdx + (i + 1));
            BodyCellStyleValue(cell, sheet);
            if (data.containsKey(columnName)) {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(data.get(columnName));
            } else if ("??????".equals(columnName)) {
                String sColLetter = CellReference.convertNumToColString(cellIdx + 1);
                String eColLetter = CellReference.convertNumToColString(cellIdx + columnNames.length - 1);
                cell.setCellFormula(String.format("SUM(%s%s:%s%s)", sColLetter, rowIdx + 1, eColLetter, rowIdx + 1));
                cell.setCellType(CellType.FORMULA);
            } else {
                cell.setCellType(CellType.NUMERIC);
                cell.setCellValue(0);
            }
        }
    }
    //???????????? ???
    static private void makeFooterFromExcel(int startRowIdx, int startCellIdx, int endRowIdx, int endCellIdx, XSSFSheet sheet) {
        
        XSSFRow row = sheet.createRow(endRowIdx);
        XSSFCell cell;
        
        cell = row.createCell(startCellIdx);
        cell.setCellValue("??????");
        AlignCenter(cell, sheet);
        
        for (int cellIdx = startCellIdx + 1; cellIdx <= endCellIdx; cellIdx++) {
            String sColLetter = CellReference.convertNumToColString(cellIdx);
            cell = row.createCell(cellIdx);
            cell.setCellFormula(String.format("SUM(%s%s:%s%s)", sColLetter, startRowIdx+1, sColLetter, endRowIdx));
            cell.setCellType(CellType.FORMULA);
            BodyCellStyleValue(cell, sheet);
        }
    }
    
    //???????????? ??? ?????????
    static private void HeaderCellStyle(XSSFCellStyle cellStyle) {
        //align ??????
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        
        //border ??????
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
        
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setFillForegroundColor(IndexedColors.AQUA.index);
    }
    //???????????? ??? ??? ?????????
    static private void BodyCellStyleMonth(XSSFCell cell, XSSFSheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
        Border(cellStyle);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setDataFormat(dataFormat.getFormat("@???"));
        cell.setCellStyle(cellStyle);
        cell.setCellType(CellType.STRING);
        
    }
    //???????????? ??? ??? ?????????
    static private void BodyCellStyleValue(XSSFCell cell, XSSFSheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        DataFormat dataFormat = sheet.getWorkbook().createDataFormat();
        Border(cellStyle);
        cellStyle.setDataFormat(dataFormat.getFormat("_-???* #,##0_-;-???* #,##0_-;_-???* \"-\"_-;_-@_-"));
        cell.setCellStyle(cellStyle);
    }
    //??? ?????????, border
    static private void AlignCenter(XSSFCell cell, XSSFSheet sheet) {
        CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
        Border(cellStyle);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellStyle(cellStyle);
        cell.setCellType(CellType.STRING);
    }
    //??? border ?????????
    static private void Border(CellStyle cellStyle) {
        cellStyle.setBorderBottom(BorderStyle.THIN);
        cellStyle.setBorderLeft(BorderStyle.THIN);
        cellStyle.setBorderRight(BorderStyle.THIN);
        cellStyle.setBorderTop(BorderStyle.THIN);
    }
    
}
