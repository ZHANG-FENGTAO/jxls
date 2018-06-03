package com.example.demo;

import net.sf.jxls.transformer.XLSTransformer;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created zft on 2018/6/3.
 */
@Controller
@RequestMapping("/api/jxls")
public class JxlsTestController {

    private static final Logger logger = LoggerFactory.getLogger(JxlsTestController.class);

    @RequestMapping("/export")
    public void export(HttpServletResponse response) throws Exception {
        User user = new User();
        user.setName("张三");
        user.setSex(1);
        User user1 = new User();
        user1.setName("李四");
        user1.setSex(2);
        List<User> list = new ArrayList<User>();
        list.add(user);
        list.add(user1);
        Map<String, Object> data = new HashMap<>(16);
        data.put("list", list);
        String fileName = "test";
        try {
            XLSTransformer xlsTransformer = new XLSTransformer();
            File tmpFile = File.createTempFile("test", ".xlsx");
            String tplName = "test.xls";
            String templatePath = Thread.currentThread()
                    .getContextClassLoader().getResource(tplName).getPath();
            xlsTransformer.transformXLS(templatePath, data, tmpFile.getAbsolutePath());
            ExcelUtils.down(tmpFile.getAbsolutePath(), response, fileName);
        } catch (Exception e) {
            logger.error("export error", e);
        }
    }
}
