package com.demo.freemark.wordtemplateexport.service;

import javax.servlet.http.HttpServletResponse;

public interface IExportService {
    void exportWord(HttpServletResponse response) throws Exception;
}
