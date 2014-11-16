package com.thoughtworks;

public class Main {

    public static void main(String[] args)
    {

        if (args.length != 4) usage();

        PhoneBill phoneBill = new PhoneBill();
        phoneBill.month = Config.valueOf("MONTH");
        phoneBill.employeeID = "12345";
        phoneBill.employeeName = "KC";
        phoneBill.employeeEmail = "khushroc@thoughtworks.com";

        ContentTemplate t= new ContentTemplate();
        String mailContents = t.generate(phoneBill);

        System.out.println(mailContents);

        //Mail.Send(mailContents, phoneBill);
    }

    private static void usage()
    {
        System.out.println()
    }
}
