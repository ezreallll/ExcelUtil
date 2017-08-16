import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;

/**
 * Created by hongpengsun on 17/8/16.
 */
public class Data2Excel {


    public void data2Excel(List<HashMap> students) throws IOException {

        //创建2003版
        HSSFWorkbook hwb = new HSSFWorkbook();

        //创建一个分页
        HSSFSheet sheet = hwb.createSheet("sheet");

        //生成表头
        HSSFRow firstRow = sheet.createRow(0);

        String[] titles = {"序号", "姓名", "班级", "性别"};
        //创建一个以titles为长度的cell数据
        HSSFCell[] firstCell = new HSSFCell[titles.length];

        //循环，为第一行生成cell 并为标题赋值
        for (int i = 0; i < titles.length; i++) {
            firstCell[i] = firstRow.createCell(i);
            firstCell[i].setCellValue(titles[i]);
        }

        //填充表格内容
        for (int i = 0; i < students.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            HashMap dto = students.get(i);

            HSSFCell[] contentCell = new HSSFCell[titles.length];

            //循环，为第i行生成cell 并设置样式
            for (int j = 0; j < titles.length; j++) {
                contentCell[j] = row.createCell(j);
                //设置cell样式
                contentCell[j].setCellStyle(getStyle(hwb));
            }
            //为每一个cell赋值
            contentCell[0].setCellValue((String) dto.get("id"));
            contentCell[1].setCellValue((String) dto.get("name"));
            contentCell[2].setCellValue((String) dto.get("class"));
            contentCell[3].setCellValue((String) dto.get("gender"));

        }
        OutputStream out = new FileOutputStream("/users/excel/test.xls");
        hwb.write(out);
        out.close();
    }


    /**
     * 获取样式
     * @param hwb
     * @return
     */

    private HSSFCellStyle getStyle(HSSFWorkbook hwb) {
        // 生成一个样式
        HSSFCellStyle style = hwb.createCellStyle();
        // 设置这些样式
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中

        // 背景色
        style.setFillForegroundColor(HSSFColor.YELLOW.index);
        style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
        style.setFillBackgroundColor(HSSFColor.YELLOW.index);

        // 设置边框
        style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        style.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        style.setBorderRight(HSSFCellStyle.BORDER_THIN);
        style.setBorderTop(HSSFCellStyle.BORDER_THIN);
        // 自动换行
        style.setWrapText(true);

        // 生成一个字体
        HSSFFont font = hwb.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setColor(HSSFColor.RED.index);
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontName("宋体");

        // 把字体 应用到当前样式
        style.setFont(font);

        return style;
    }

}
