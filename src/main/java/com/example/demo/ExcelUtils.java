package com.example.demo;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.net.URLEncoder;

/**
 * Created zft on 2018/6/3.
 */
public class ExcelUtils {

    private static final Logger logger = LoggerFactory.getLogger(ExcelUtils.class);

    /**
     * 导出Excel
     * @param path  文件路径
     * @param response
     * @param fileName  文件名
     */
    public static void down(String path, HttpServletResponse response, String fileName){
        try {
            File file = new File(path);
            // 以流的形式下载文件。
            InputStream fis = new BufferedInputStream(new FileInputStream(path));
            byte[] buffer = new byte[fis.available()];
            try {
                fis.read(buffer);
            } catch (Exception e) {
                logger.error(e.toString());
            } finally {
                fis.close();
            }

            // 清空response
            response.reset();
            // 设置response的Header
            response.addHeader("Content-Disposition","attachment;fileName=" + URLEncoder.encode(fileName,"UTF-8") + ".xlsx");
            response.addHeader("Content-Length", "" + file.length());
            response.setContentType("application/vnd.ms-excel");
            try (OutputStream outputStream = new BufferedOutputStream(response.getOutputStream())) {
                outputStream.write(buffer);
                outputStream.flush();
            } catch (Exception e) {
                logger.error(e.toString());
            }
        } catch (IOException ex) {
            logger.error(ex.toString());
        }
    }

}
