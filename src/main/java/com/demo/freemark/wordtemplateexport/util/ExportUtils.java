package com.demo.freemark.wordtemplateexport.util;

import freemarker.template.Configuration;
import freemarker.template.Template;

import javax.servlet.http.HttpServletResponse;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Map;

public class ExportUtils {
    public static void template2Word(String templateName, Map<String, Object> data, String fileName,
                                     HttpServletResponse response) throws Exception{
        fileName = new String(fileName.getBytes("GBK"), StandardCharsets.ISO_8859_1) + ".doc";
        response.setContentType("application/x-msdownload;charset=UTF-8");
        response.addHeader("Content-Disposition", "attachment;filename=" + fileName + ";charset=UTF-8");

        Configuration configuration = new Configuration(Configuration.getVersion());
        configuration.setDefaultEncoding("UTF-8");
        // 设置模板存放的路径
        configuration.setClassForTemplateLoading(ExportUtils.class, "/template");
        Writer writer = new OutputStreamWriter(response.getOutputStream());
        Template t = configuration.getTemplate(templateName);
        t.process(data, writer);
    }
}
