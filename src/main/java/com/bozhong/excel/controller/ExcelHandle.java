package com.bozhong.excel.controller;

import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.IOException;

@RestController
@RequestMapping(value = "/excel")
public class ExcelHandle {

    @ResponseBody
    @RequestMapping(value = "/receive")
    String receive(HttpServletRequest request, HttpServletResponse response, MultipartFile file) throws IOException {
        System.out.println("receive");
        System.out.println(file.getName());//获取上传文件的表单名称
        System.out.println(file.getContentType());//MIME类型
        System.out.println(file.getSize());//文件大小
        System.out.println(file.getOriginalFilename());//获取上传文件的完整名称
        File destFile= new File("d:/shiroro/upload");
        if(!destFile.exists()) {
            destFile.mkdirs();
        }
        File uploadFile = new File(destFile, file.getOriginalFilename());
        file.transferTo(uploadFile);
        return null;
    }


}
