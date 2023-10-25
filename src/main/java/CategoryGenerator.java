import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddressList;

import java.util.ArrayList;

public class CategoryGenerator {
    
    private Workbook workbook;
    private final String sheetName = "카테고리";
    private int row;
    private Sheet sheet;
    
    private CategoryGenerator() {}
    
    public CategoryGenerator(Workbook workbook) {
        this.row = 0;
        this.workbook = workbook;
        this.sheet = workbook.createSheet(sheetName);
    }
    
    /**
     * 카테고리 생성명 요소 생성
     * @param name 카테고리명
     * @param arr 카테고리 요소
     */
    public void createCategory(String name, String[]  arr) {
    
        /**
         *  카테고리고명 | 요소1 | 요소2  | 요소3 ...
         *  품목         | 과일  | 채소   |
         *  과일         | 사과  | 바나나 | 망고
         *  채소         | 당긍  | 오이   | 파프리카
         */
    
        Row rows = this.sheet.createRow(this.row); // 현재시트에서 새로운 로우 작성 0부터 시작
        rows.createCell(0).setCellValue(name); // 해당하는 로우에서 첫번제의 내요을 카태고리명으로
        
        // 들어있는 요소의 수만큼 반복해서 데이타를 입력
        int j = 1;
        for(String data : arr) {
            rows.createCell(j++).setCellValue(data);
        }
    
        // 이름관리자에 등록
        createName(name, sheetName + "!"
                + new CellAddress(this.row, 1).formatAsString() +":"
                + new CellAddress(this.row, j-1).formatAsString());
        this.row++;
    }
    
    /**
     * 이름관리자에 등록하는 함수
     * @param name 생성하는 카테고리명
     * @param ref 참조하는 주소 ex) 시트명!D1:D10
     * @return Name
     */
    private Name createName(String name, String ref) {
        Name namedCell = workbook.createName();
        namedCell.setNameName(name);
        namedCell.setRefersToFormula(ref);
        
        return namedCell;
    }
}
