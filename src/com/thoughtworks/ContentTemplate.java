package com.thoughtworks;

import org.apache.commons.io.FileUtils;
import org.stringtemplate.v4.ST;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.CharsetDecoder;
import java.nio.charset.CharsetEncoder;
import java.nio.file.Files;
import java.nio.file.Paths;

public class ContentTemplate {

    private static String rawTemplate = "";

    public ContentTemplate()
    {
        Charset encoding = Charset.defaultCharset();
        if (rawTemplate.equals(""))
        {
            try
            {
                String myPath = this.getClass().getClassLoader().getResource("").getPath();
                File templateFile = new File(myPath + File.separator + "MessageContents.stg");
                rawTemplate = FileUtils.readFileToString(templateFile);
            }
            catch (IOException ex)
            {
                ex.printStackTrace();
            }
        }
    }

    public String generate(PhoneBill phoneBill)
    {
        System.out.println("Entering ContentTemplate.generate()");
        ST template = new ST(rawTemplate, '$', '$');
        template.add("phoneBill", phoneBill);
        String renderedText = template.render();
        System.out.println("Exiting ContentTemplate.generate()");
        return renderedText;
    }
}
