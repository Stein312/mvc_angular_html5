package com.example.tzForMBOIC.controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.servlet.ModelAndView;

@Controller
public class ExcelController {

    @RequestMapping(value = "/excel", method= RequestMethod.GET)
    public ModelAndView excel() {
        System.out.println("ExcelPDFController excel is called");

        //excelDocument - look excel-pdf-config.xml
        return new ModelAndView("xlsHandler");

    }
    @RequestMapping(value = "/pdf", method= RequestMethod.GET)
    public ModelAndView pdf() {
        System.out.println("ExcelPDFController excel is called");

        //excelDocument - look excel-pdf-config.xml
        return new ModelAndView("generatePDF");

    }
}
