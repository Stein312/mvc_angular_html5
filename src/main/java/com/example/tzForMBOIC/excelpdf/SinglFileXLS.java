package com.example.tzForMBOIC.excelpdf;

import org.springframework.web.multipart.MultipartFile;

import java.io.File;


public class SinglFileXLS {

    private static File instance;
    private static String filename="excelDownload.xls";
    private static String rootPath = System.getProperty("catalina.home");
    private SinglFileXLS(){}
    public static File getInstance(){
        if(instance==null){

            instance=new File(rootPath + File.separator + filename);
        }
        return instance;
    }

    public static String getFilename() {
        return filename;
    }

    public static void setFilename(String filename) {
        SinglFileXLS.filename = filename;
        instance=new File(rootPath + File.separator + filename);

    }
}
