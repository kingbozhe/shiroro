package com.bozhong.excel.controller;

import net.sf.json.JSONObject;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.servlet.ModelAndView;

@Controller
@RequestMapping(value = "/shiroro")
public class PageController {

    @RequestMapping("/index")
    public ModelAndView indexPage() {
        new ExcelHandle().delete("d:/shiroro/upload");
        new ExcelHandle().fileMap.clear();
        new ExcelHandle().fileList.clear();
        ModelAndView modelAndView = new ModelAndView();
        modelAndView.setViewName("load");
        return modelAndView;
    }


}
