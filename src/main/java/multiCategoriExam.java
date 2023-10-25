import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;

import static org.apache.poi.ss.util.CellUtil.createCell;

public class multiCategoriExam {
    public static void main(String[] args) throws FileNotFoundException {
    
        Workbook wb = new XSSFWorkbook();
        
        CategoryGenerator cg = new CategoryGenerator(wb);
        
        // 카테고리 생성
        cg.createCategory("품목", new String[]{"과일", "채소"});
        cg.createCategory("과일", new String[]{"사과", "바나나", "망고"});
        cg.createCategory("채소", new String[]{"당근", "파프리카", "오이"});
        
        // 시트생성
        Sheet sheet = wb.createSheet("목록들");
        
        // 제약조건 주기 =INDIRECT 함수에서 이름관리자에 있는 내용을 기반으로 가져옴
        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint constraint1 = dvHelper.createFormulaListConstraint("=INDIRECT(\"품목\")");
        DataValidationConstraint constraint2 = dvHelper.createFormulaListConstraint("=INDIRECT(B2)");

//        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);

        // 해당 조건을 어느영역만큼 할것인지에 대한 내용
        // 1열 1 - 100
        DataValidation valid1 = dvHelper.createValidation(constraint1, new CellRangeAddressList(1, 100, 1, 1));
        // 2열 1 - 100
        DataValidation valid2 = dvHelper.createValidation(constraint2, new CellRangeAddressList(1, 100, 2, 2));

        // 제약조건 최종적으로 추가
        sheet.addValidationData(valid1);
        sheet.addValidationData(valid2);
        
        // 파일 생성현재경로에 결과 파일생성
        writeFile(wb, "test");
        
    }
    
    private static void writeFile(Workbook wb, String name) {
    
        // Write the output to a file
        try (OutputStream fileOut = new FileOutputStream(name + ".xlsx")) {
            wb.write(fileOut);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }finally {
            try { wb.close(); } catch (IOException ignored) { }
        }
        
    }
}
