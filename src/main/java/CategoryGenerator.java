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
    
    public void createCategory(String name, String[]  arr) {
    
        Row rows = this.sheet.createRow(this.row);
        rows.createCell(0).setCellValue(name);
        
        int j = 1;
        for(String data : arr) {
            rows.createCell(j++).setCellValue(data);
        }
    
        createName(name, sheetName + "!"
                + new CellAddress(this.row, 1).formatAsString() +":"
                + new CellAddress(this.row, j-1).formatAsString());
        this.row++;
    }
    
    private Name createName(String name, String ref) {
        Name namedCell = workbook.createName();
        namedCell.setNameName(name);
        namedCell.setRefersToFormula(ref);
        
        return namedCell;
    }
}
