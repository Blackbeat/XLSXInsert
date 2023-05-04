import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

public class ReacExcel4 {
    public static List<String> account1 = new ArrayList<String>();
    public static List<String> account2 = new ArrayList<String>();
    public static List<String> account3 = new ArrayList<String>();
    public static List<String> account4 = new ArrayList<String>();

    public static void main(String[] args) throws ClassNotFoundException {

        List<String> list = new ArrayList<>();

        FileInputStream fis = null;
        XSSFWorkbook workbook = null;

        try {

            String filePath = "C:\\Users\\jbt\\Desktop\\메타데이터\\특이 메타데이터2.xlsx";
            fis = new FileInputStream(filePath);
            // XSSFWorkbook은 엑셀파일 전체 내용을 담고 있는 객체
            workbook = new XSSFWorkbook(fis);

            // 탐색에 사용할 Sheet, Row, Cell 객체
            XSSFSheet curSheet;
            XSSFRow curRow;
            XSSFCell curCell;
//           poiList vo;

            // Sheet 탐색 for문
            for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                // 현재 Sheet 반환
                curSheet = workbook.getSheetAt(0);

                // row 탐색 for문
                for (int rowIndex = 0; rowIndex < curSheet.getPhysicalNumberOfRows(); rowIndex++) {
                    // row 0은 헤더정보이기 때문에 무시
                    if (rowIndex != 0) {
                        // 현재 row 반환
                        curRow = curSheet.getRow(rowIndex);
//                       vo = new poiList();
                        String value;
                        // row의 첫번째 cell값이 비어있지 않은 경우 만 cell탐색
                        /*System.out.println("---"+curRow.getCell(0));*/
                        if (curRow.getCell(0) != null) { // 첫번째열 모자를 때 구분
                            if (!"".equals(curRow.getCell(0).getStringCellValue())) {

                                // cell 탐색 for 문
                                for (int cellIndex = 0; cellIndex < curRow.getPhysicalNumberOfCells(); cellIndex++) {
                                    curCell = curRow.getCell(cellIndex);

                                    if (true) {
                                        value = "";
                                        // cell 스타일이 다르더라도 String으로 반환 받음
                                        if (curCell != null) { //각 열의 목록 길이에 따라 분기
                                            switch (curCell.getCellType()) {
                                                case STRING:
                                                    value = curCell.getStringCellValue() + "";
                                                    break;
                                                default:
                                                    value = new String();
                                                    break;
                                            }
                                        }
                                        switch (cellIndex) {
                                            case 0: // name
                                                account1.add(rowIndex - 1, value);
                                                break;

                                            case 1: // 암호
                                                account2.add(rowIndex - 1, value);
                                                break;

                                            case 2: // 여기서 하나씩 분류
                                                account3.add(rowIndex - 1, value);
                                                break;

                                            case 3: // 여기서 하나씩 분류
                                                account4.add(rowIndex - 1, value);
                                                break;

                                            default:
                                                break;
                                        }
                                    }
                                }
                            }
                        }

                    }
                }
            }


        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

        for(int n = 0; n<account1.size(); n++){
            List<String> CellList = new ArrayList<String>();
            try {

                File f = new File(account4.get(n));
                /*System.out.println(account4.get(n));*/

                String fileName = f.getName();

                String sheetName = "데이터 정의서";

                Workbook wb = WorkbookFactory.create(f);
                Sheet sheet = wb.getSheet(sheetName);

                /*int rowIdx = 5;
                int colIdx = 1;*/

                /*if(true){
                    Row row = sheet.getRow(3);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                for(int i = 4; i<5; i++){
                    Row row = sheet.getRow(i);
                    Cell cell = row.getCell(1);

                    Object cellValue = cell.getStringCellValue();
                    *//*System.out.println(cellValue);*//*


                    XSSFRichTextString cellValue2 = (XSSFRichTextString)cell.getRichStringCellValue();
                    *//*System.out.println(cellValue2);*//*

                    String cellValueStr = cellValue2.getString();
                    String newCellValue = "";

                    int numFormattingRuns = cellValue2.numFormattingRuns();
                    for (int fmtIdx = 0; fmtIdx < numFormattingRuns; fmtIdx++) {

                        int begin = cellValue2.getIndexOfFormattingRun(fmtIdx);
                        int length = cellValue2.getLengthOfFormattingRun(fmtIdx);

                        *//*System.out.println(String.format("idx : %d, begin : %d, length : %d", fmtIdx, begin, length));*//*


                        XSSFFont formatFont = cellValue2.getFontOfFormattingRun(fmtIdx);

                        boolean isBold = false;

                        if(formatFont != null){
                            *//*System.out.println(formatFont.getBold());*//*

                            isBold = formatFont.getBold();
                        }


                        if(isBold) {
                            CellList.add(cellValueStr.substring(begin, begin+length));
                        } else {
                            continue;
                        }

                    }

                }*/

                for(int j=3; j<7; j++){
                    Row row = sheet.getRow(j);
                    Cell cell = row.getCell(1);


                    XSSFRichTextString cellValue2 = (XSSFRichTextString)cell.getRichStringCellValue();
                    /* System.out.println(cellValue2);*/

                    String cellValueStr = cellValue2.getString();
                    /*System.out.println(cellValueStr);*/
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(8);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }
                if(true){
                    Row row = sheet.getRow(8);
                    Cell cell = row.getCell(3);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(9);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(9);
                    Cell cell = row.getCell(3);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(10);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(12);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(14);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(14);
                    Cell cell = row.getCell(3);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }
                if(true){
                    Row row = sheet.getRow(16);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }
                if(true){
                    Row row = sheet.getRow(16);
                    Cell cell = row.getCell(3);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(17);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }

                if(true){
                    Row row = sheet.getRow(18);
                    Cell cell = row.getCell(1);

                    XSSFRichTextString cellValue = (XSSFRichTextString)cell.getRichStringCellValue();

                    String cellValueStr = cellValue.getString();
                    CellList.add(cellValueStr);
                }


                /*if(true){
                    Row row = sheet.getRow(20);
                    Cell cell = row.getCell(1);

                    Object cellValue = cell.getStringCellValue();
                    System.out.println(cellValue);


                    XSSFRichTextString cellValue2 = (XSSFRichTextString)cell.getRichStringCellValue();
                    System.out.println(cellValue2);

                    String cellValueStr = cellValue2.getString();
                    String newCellValue = "";

                    int numFormattingRuns = cellValue2.numFormattingRuns();
                    for (int fmtIdx = 0; fmtIdx < numFormattingRuns; fmtIdx++) {

                        int begin = cellValue2.getIndexOfFormattingRun(fmtIdx);
                        int length = cellValue2.getLengthOfFormattingRun(fmtIdx);

                        System.out.println(String.format("idx : %d, begin : %d, length : %d", fmtIdx, begin, length));


                        XSSFFont formatFont = cellValue2.getFontOfFormattingRun(fmtIdx);

                        boolean isBold = false;

                        if(formatFont != null){
                            System.out.println(formatFont.getBold());

                            isBold = formatFont.getBold();
                        }


                        if(isBold) {
                            CellList.add(cellValueStr.substring(begin, begin+length));
                        } else {
                            CellList.add("");
                        }

                    }

                }

                for(int k=21; k<23; k++){
                    Row row = sheet.getRow(k);
                    Cell cell = row.getCell(1);


                    XSSFRichTextString cellValue2 = (XSSFRichTextString)cell.getRichStringCellValue();
                    System.out.println(cellValue2);

                    String cellValueStr = cellValue2.getString();
                    System.out.println(cellValueStr);
                    CellList.add(cellValueStr);
                }*/

                /*for(int l=23; l<25; l++){
                    Row row = sheet.getRow(l);
                    Cell cell = row.getCell(1);


                    XSSFRichTextString cellValue2 = (XSSFRichTextString)cell.getRichStringCellValue();
                    *//*System.out.println(cellValue2);*//*

                    String cellValueStr = cellValue2.getString();
                    *//*System.out.println(cellValueStr);*//*
                    CellList.add(cellValueStr);
                }*/
                /*for(int a = 0; a<23; a++){
                    System.out.println(CellList.get(a));
                }*/


                wb.close();

            } catch(Exception e) {
                e.printStackTrace();
            }

            /*System.out.println(CellList);*/

            Class.forName("org.postgresql.Driver");

            String connurl  = "jdbc:postgresql://192.168.0.6:5432/ehdata";
            String user     = "ehdata";
            String password = "ehdata00";
            PreparedStatement pstmt = null;
            try (Connection connection = DriverManager.getConnection(connurl,user,password);) {
                String sql="INSERT INTO public.tn_meta_table_type2 (meta_id, tbl_id, meta_lbl, data_cat, data_descr, tbl_nm, data_fomt, data_pvsn, data_src, prdctn_cy, coll_mthd, data_rte, spat_rng, ymd_ty, temp_rng, tkcg_inst, pic, note) VALUES(uuid_generate_v4(),?::UUID,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?);";
                /*connection.setAutoCommit(false);*/
                pstmt = connection.prepareStatement(sql);
                pstmt.setString(1,account1.get(n));
                pstmt.setString(2,account3.get(n));
                for(int m = 3; m<17; m++){
                    pstmt.setString(m, CellList.get(m-3));
                }
                pstmt.setString(17,CellList.get(14) + CellList.get(15));
                /*pstmt.setString(18,CellList.get(17));
                pstmt.setString(19,CellList.get(18));
                pstmt.setString(20,CellList.get(19));*/
                System.out.println(pstmt);
                pstmt.executeUpdate();
                pstmt.close();
            }
            catch (SQLException e){
                e.printStackTrace();
            }
        }
    }
}
