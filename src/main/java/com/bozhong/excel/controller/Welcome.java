package com.bozhong.excel.controller;

import net.sf.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;

@Controller
@RequestMapping(value = "/cat")
public class Welcome {


    @ResponseBody
    @RequestMapping(value = "/welcome")
    String welcome(@RequestParam(value = "userid") String userId){
        return null;
    }


}
