package com.example.demo.poi;



import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.UnsupportedEncodingException;

/**
 * Created zft on 2018/6/3.
 */
public class UPoi {

//    public static CellStyle bigTitle(Workbook wb) {
//        CellStyle style = wb.createCellStyle();
//        Font font = wb.createFont();
//        font.setFontName("宋体");
//        font.setFontHeightInPoints((short) 16);
//        font.setBoldweight(Font.BOLDWEIGHT_BOLD);                    //字体加粗
//
//        style.setFont(font);
//
//        style.setAlignment(CellStyle.ALIGN_CENTER);                    //横向居中
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);        //纵向居中
//
//        return style;
//    }
//
//
//    public static CellStyle title(Workbook wb) {
//        CellStyle style = wb.createCellStyle();
//        Font font = wb.createFont();
//        font.setFontName("黑体");
//        font.setFontHeightInPoints((short) 12);
//        style.setFont(font);
//
//        style.setAlignment(CellStyle.ALIGN_CENTER);                    //横向居中
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);        //纵向居中
//
//        style.setBorderTop(CellStyle.BORDER_THIN);                    //上细线
//        style.setBorderBottom(CellStyle.BORDER_THIN);                //下细线
//        style.setBorderLeft(CellStyle.BORDER_THIN);                    //左细线
//        style.setBorderRight(CellStyle.BORDER_THIN);                //右细线
//
//        return style;
//    }
//
//    //文字样式
//    public static CellStyle text(Workbook wb) {
//        CellStyle style = wb.createCellStyle();
//        Font font = wb.createFont();
//        font.setFontName("Times New Roman");
//        font.setFontHeightInPoints((short) 10);
//
//        style.setFont(font);
//
//        style.setAlignment(CellStyle.ALIGN_LEFT);                    //横向居左
//        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);        //纵向居中
//
//        style.setBorderTop(CellStyle.BORDER_THIN);                    //上细线
//        style.setBorderBottom(CellStyle.BORDER_THIN);                //下细线
//        style.setBorderLeft(CellStyle.BORDER_THIN);                    //左细线
//        style.setBorderRight(CellStyle.BORDER_THIN);                //右细线
//
//        return style;
//    }
//
//    public static void download(ByteArrayOutputStream byteArrayOutputStream, HttpServletResponse response, String returnName) throws IOException {
//        response.setContentType("application/octet-stream;charset=utf-8");
//        try {
//            returnName = response.encodeURL(new String(returnName.getBytes("utf-8"), "iso8859-1"));
//        } catch (UnsupportedEncodingException e) {
//            e.printStackTrace();
//        }
//        response.addHeader("Content-Disposition", "attachment;filename=" + returnName);
//        response.setContentLength(byteArrayOutputStream.size());
//
//        ServletOutputStream outputstream = response.getOutputStream();
//        byteArrayOutputStream.writeTo(outputstream);
//        byteArrayOutputStream.close();
//        outputstream.flush();
//    }

}
