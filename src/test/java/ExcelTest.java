import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.Iterator;

public class ExcelTest {
    public static void main(String[] args) throws IOException {
        XSSFWorkbook workBook=new XSSFWorkbook("Excel.xlsx");
        int num_sheets=workBook.getNumberOfSheets();
        System.out.println(num_sheets);

        for(int i=0;i<num_sheets;i++){
            if(workBook.getSheetName(i).equalsIgnoreCase("Testdata")){
                XSSFSheet sheet =workBook.getSheetAt(i);
                Iterator<Row> rows=sheet.iterator();
                Row firstrow=rows.next();

                Iterator<Cell> cells=firstrow.cellIterator();

                int column=0;
                while(cells.hasNext()){
                    Cell currentcell=cells.next();
                    if(currentcell.getStringCellValue().equalsIgnoreCase("Testcase")){
                        //get columns
                        System.out.println("Column ="+column);
                        break;
                    }
                    column++;
                }

                while(rows.hasNext()){
                    Row row=rows.next();
                    Cell cellsofreqdcolumn=row.getCell(column);
                    if(cellsofreqdcolumn.getStringCellValue().equalsIgnoreCase("Imran")){
                        Iterator<Cell> reqdrowcell=row.cellIterator();
                        while (reqdrowcell.hasNext())
                            System.out.println(reqdrowcell.next().getStringCellValue());
                    }

                }
            }




        }


    }
}
