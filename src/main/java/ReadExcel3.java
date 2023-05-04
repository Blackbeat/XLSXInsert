import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ReadExcel3 {

    public static List<String> account1 = new ArrayList<String>();
    public static List<String> account2 = new ArrayList<String>();
    public static List<String> account3 = new ArrayList<String>();
    public static List<String> account4 = new ArrayList<String>();
    public static void main(String[] args){
        FileInputStream fis = null;
        XSSFWorkbook workbook = null;

        try {

            String filePath = "C:\\Users\\jbt\\Desktop\\메타데이터 정보.xlsx";
            fis= new FileInputStream(filePath);
            // XSSFWorkbook은 엑셀파일 전체 내용을 담고 있는 객체
            workbook = new XSSFWorkbook (fis);

            // 탐색에 사용할 Sheet, Row, Cell 객체
            XSSFSheet curSheet;
            XSSFRow curRow;
            XSSFCell  curCell;
//           poiList vo;

            // Sheet 탐색 for문
            for(int sheetIndex = 0 ; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
                // 현재 Sheet 반환
                curSheet = workbook.getSheetAt(0);

                // row 탐색 for문
                for(int rowIndex=0; rowIndex < curSheet.getPhysicalNumberOfRows(); rowIndex++) {
                    // row 0은 헤더정보이기 때문에 무시
                    if(rowIndex != 0) {
                        // 현재 row 반환
                        curRow = curSheet.getRow(rowIndex);
//                       vo = new poiList();
                        String value;
                        // row의 첫번째 cell값이 비어있지 않은 경우 만 cell탐색
                        /*System.out.println("---"+curRow.getCell(0));*/
                        if(curRow.getCell(0) != null){ // 첫번째열 모자를 때 구분
                            if(!"".equals(curRow.getCell(0).getStringCellValue())) {

                                // cell 탐색 for 문
                                for(int cellIndex=0;cellIndex<curRow.getPhysicalNumberOfCells(); cellIndex++) {
                                    curCell = curRow.getCell(cellIndex);

                                    if(true) {
                                        value = "";
                                        // cell 스타일이 다르더라도 String으로 반환 받음
                                        if(curCell != null){ //각 열의 목록 길이에 따라 분기
                                            switch (curCell.getCellType()){
                                                case STRING:
                                                    value = curCell.getStringCellValue()+"";
                                                    break;
                                                default:
                                                    value = new String();
                                                    break;
                                            }
                                        }
                                        switch (cellIndex) {
                                            case 0: // name
                                                account1.add(rowIndex-1, value);
                                                break;

                                            case 1: // 암호
                                                account2.add(rowIndex-1, value);
                                                break;

                                            case 2: // 여기서 하나씩 분류
                                                account3.add(rowIndex-1, value);
                                                break;

                                            case 3: // 여기서 하나씩 분류
                                                account4.add(rowIndex-1, value);
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

            System.out.println(account3.get(0));
        } catch (FileNotFoundException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
