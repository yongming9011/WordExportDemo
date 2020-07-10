package com.demo.freemark.wordtemplateexport.controller;

import com.demo.freemark.wordtemplateexport.service.IExportService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import javax.servlet.http.HttpServletResponse;

@Controller
@RequestMapping("export")
public class TemplateExportController {
    @Autowired
    private IExportService exportService;

    @GetMapping("word")
    public void exportWord(HttpServletResponse response) throws Exception {
        exportService.exportWord(response);
    }
}
