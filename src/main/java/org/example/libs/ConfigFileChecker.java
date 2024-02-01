package org.example.libs;

import java.io.File;

public class ConfigFileChecker {
    public static Response configFileChecker(String filePath,String fileName){
        File configDir = new File(filePath);
        File[] files = configDir.listFiles();

        if (files == null || files.length == 0) {
            System.exit(0);
            return new Response(false, "No Config file Found in the directory.");
        }

        try {
            for (File file : files) {
                String filename = file.getName();
                fileName = file.getAbsolutePath();
                if(filename.endsWith(".xml")) {
                    return new Response(true,filename);
                }else{
                    return new Response(false,"No valid config file found.");
                }
            }
        } catch (Exception e) {
            return new Response(false,e.getMessage());
        }

        return new Response(false, "Error processing the config file, Please make sure the config is in good shape and try again");

    }
}
