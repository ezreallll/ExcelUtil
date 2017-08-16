import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

/**
 * Created by hongpengsun on 17/8/16.
 *
 * 获取合并表格的值
 *
 * =====================
 * name1    1   2   3
 * =====================
 *          1   2   3
 *       ===============
 * name2    1   2   3
 *       ===============
 *          1   2   3
 * =====================
 */
public class Excel2Data {

    /**
     * 获取检查地点的json字符串
     * @param
     * @return
     * @throws IOException
     */
    public static HashMap getCheckSiteJson(String filePath) throws IOException {
        Sheet sheet=getSheet(filePath);
        //获取所有的表格的行数
        int count = sheet.getLastRowNum()+1;

        HashMap result=new HashMap();
        //从第一行开始循环，排除标题
        for(int i=1;i<count;i++){

            String cellName=getCellValue(sheet.getRow(i).getCell(0));
            //如果是合并表格，用来储存所有的行
            List<HashMap> maps=new ArrayList<HashMap>();

            //如果是合并表格
            if(isMergedRegion(sheet,i,0)){
                //获取合并表格的最后一行
                int lastRow=getMergedLastRow(sheet,i);
                for (;i<=lastRow;i++){
                    Row row =sheet.getRow(i);
                    HashMap mapCell=new HashMap();
                    mapCell.put("col1",getCellValue(row.getCell(1)));
                    mapCell.put("col2",getCellValue(row.getCell(2)));
                    mapCell.put("col3",getCellValue(row.getCell(3)));
                    maps.add(mapCell);
                }
                result.put(cellName,maps);
                i--;
            }else {//如果是单行
                Row row =sheet.getRow(i);

                HashMap mapCell=new HashMap();
                mapCell.put("col1",getCellValue(row.getCell(1)));
                mapCell.put("col2",getCellValue(row.getCell(2)));
                mapCell.put("col3",getCellValue(row.getCell(3)));
                maps.add(mapCell);

                result.put(cellName,maps);
            }

        }

        return result;
    }

    /**
     * 获取sheet
     * @param filePath
     * @return
     * @throws IOException
     */
    private static Sheet getSheet(String filePath)throws IOException{
        boolean isE2007 = false;    //判断是否是excel2007格式
        if(filePath.endsWith("xlsx")){
            isE2007 = true;
        }
        Workbook wb=null;
        InputStream inputStream=new FileInputStream(filePath);
        if(isE2007){
            wb = new XSSFWorkbook(inputStream);
        }else{
            wb = new HSSFWorkbook(inputStream);
        }
        return wb.getSheetAt(0);
    }


    /**
     * 获取单元格的值
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell){
        if(cell == null) return "";
        if(cell.getCellType() == Cell.CELL_TYPE_STRING){
            return cell.getStringCellValue();
        }else if(cell.getCellType() == Cell.CELL_TYPE_BOOLEAN){
            return String.valueOf(cell.getBooleanCellValue());
        }else if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){
            return cell.getCellFormula() ;
        }else if(cell.getCellType() == Cell.CELL_TYPE_NUMERIC){
            return String.valueOf(cell.getNumericCellValue());
        }
        return "";
    }


    /**
     * 判断是否是合并单元格
     * @param sheet
     * @param row
     * @return
     */
    private static boolean isMergedRegion(Sheet sheet,int row ,int col) {
        //获取表格内的合并单元格数量
        int sheetMergeCount = sheet.getNumMergedRegions();
        //循环所有的合并表格
        for (int i = 0; i < sheetMergeCount; i++) {
            //获取第i个合并表格
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            int firstColumn=range.getFirstColumn();
            int lastColumn=range.getLastColumn();
            if(row >= firstRow && row <= lastRow){ //判断当前行是否在合并表格范围内
                if(col>=firstColumn&&col<=lastColumn) { //判断当前列是否在合并表格范围内
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 获取合并单元格的最后一行
     * @param sheet
     * @param row
     * @return
     */
    private static int getMergedLastRow(Sheet sheet,int row){
        int last=0;
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                last = lastRow;
            }
        }
        return last;
    }

    public static void main(String[] args){

    }
}
