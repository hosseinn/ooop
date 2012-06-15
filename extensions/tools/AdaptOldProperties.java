import java.io.*;



public class AdaptOldProperties {

    private static String endChar = "&";

    public static void main(String[] args) {
        if((args.length == 3 || args.length == 4) && (args[0].equals("-addNums") || args[0].equals("-adapt")) && !args[1].equals("") && !args[2].equals("")){
            if(args.length == 4 && args[3].length() != 1){
                System.out.println("A 4. argumentum csak egy karakter lehet");
                printUsing();
            } else{
                File file1 = new File(args[1]);
                File file2 = new File(args[2]);
                if(!file1.exists() || !file2.exists()){
                    if(!file1.exists())
                        System.out.println("Nem létező fájl: " + args[1]);
                    if(!file2.exists())
                        System.out.println("Nem létező fájl: " + args[2]);
                    printUsing();
                }else{
                    if(args.length == 4)
                        endChar = args[3];
                    if(args[0].equals("-addNums"))
                        addNums(args);
                    if(args[0].equals("-adapt"))
                        adapt(args);
                }
            }
	} else {
            printUsing();
	}
    }

    public static void printUsing(){
        System.out.println("\nHasználat:\n"
          + "1. lépés - módosuló sorok megszámozása:\n"
          + "java AdaptOldProperties -addNums old.properties new_en_US.properties\n"
          + "2. lépes - régi tulajdonságok adaptálása, minden properties fájlon:\n"
          + "java AdaptOldProperties -adapt old.properties_ old-xx.properties\n"
          + "a 4. paraméter a számok záró karaktere - opcionális");
    }

    public static void addNums(String[] args){
        try{
  		FileInputStream fstream2 = new FileInputStream(args[2]);
  		DataInputStream in2 = new DataInputStream(fstream2);
  		BufferedReader brNewProps = new BufferedReader(new InputStreamReader(in2));
                String newFileName = args[1].split(".properties")[0] + "_.properties";
		BufferedWriter outputFile = new BufferedWriter(new FileWriter(newFileName));
                System.out.println("\n------------------------------ addNums -----------------------------");
                System.out.println(newFileName + " fájl létrehozva.");
                System.out.println("Módosított sorok " + args[2] + " alapján:");
  		String strLineOldProp, strLineNewProp, sNewValue, sOldNum;

  		while((strLineNewProp = brNewProps.readLine()) != null)   {
                    int index = strLineNewProp.indexOf("=");
                    sNewValue = strLineNewProp.substring(index + 1);
                    sOldNum = "";
                    if(!sNewValue.equals("")){
                        FileInputStream fstream1 = new FileInputStream(args[1]);
                        DataInputStream in1 = new DataInputStream(fstream1);
                        BufferedReader brOldProps = new BufferedReader(new InputStreamReader(in1));
                        while((strLineOldProp = brOldProps.readLine()) != null) {
                            index = strLineOldProp.indexOf("=");
                            if(strLineOldProp.substring(index + 1).equals(sNewValue)){
                                index = strLineOldProp.indexOf(".");
                                if(index > 0)
                                    sOldNum = strLineOldProp.substring(0, index);
                                break;
                            }
                        }
                        in1.close();
                    }	
                    if(!sOldNum.equals(""))
                        sOldNum = " " + sOldNum + endChar;
                    outputFile.write(strLineNewProp + sOldNum + "\n");
                    if(!sOldNum.equals(""))
                        System.out.println(strLineNewProp + sOldNum);
  		}
  		in2.close();
		outputFile.close();
                System.out.println("-------------------------------    --------------------------------\n");
    	}catch (Exception e){
  		System.err.println("Hiba: " + e.getLocalizedMessage());
  	}
    }

    public static void adapt(String[] args){
        try{
            FileInputStream fstream1 = new FileInputStream(args[1]);
            DataInputStream in1 = new DataInputStream(fstream1);
            BufferedReader brNewProps = new BufferedReader(new InputStreamReader(in1));
            String outputFileName = args[2].indexOf("-") > -1 ? args[2].replace('-', '_') : args[2] + "_";
            BufferedWriter outputFile = new BufferedWriter(new FileWriter(outputFileName));
            System.out.println("\n------------------------------- adapt ------------------------------");
            System.out.println(outputFileName + " fájl létrehozva.");
            System.out.println("Módosított sorok " + args[1] + " alapján:");
            String strLineNewProp, strLineOldProp;

            while ((strLineNewProp = brNewProps.readLine()) != null)   {
  		if(strLineNewProp.endsWith(endChar)){
                    String[] parts = strLineNewProp.split(" ");
                    String lastPart = parts[parts.length - 1];
                    String sNum = lastPart.substring(0, lastPart.length() - 1);
                    FileInputStream fstream2 = new FileInputStream(args[2]);
                    DataInputStream in2 = new DataInputStream(fstream2);
                    BufferedReader brOldProps = new BufferedReader(new InputStreamReader(in2));
                    String sValue = "";
                    while ((strLineOldProp = brOldProps.readLine()) != null)   {
                        if(strLineOldProp.startsWith(sNum)){
                            int index = strLineOldProp.indexOf("=");
                            sValue = strLineOldProp.substring(index + 1);
                            String newValuePart = parts[parts.length - 2];
                            if(newValuePart.endsWith(":") && !sValue.endsWith(":")){
                                sValue += ":";
                                System.out.println("Hiányzó kettőspont a " + args[2] + " -ben\n" + strLineOldProp + " - javítva");
                            }
                            break;
                        }
                    }
                    in2.close();
                    int index = strLineNewProp.indexOf("=");
                    String subStr = strLineNewProp.substring(0, index + 1);
                    outputFile.write(subStr + sValue + "\n");
                    System.out.println(subStr + sValue);
                }else{
                    outputFile.write(strLineNewProp + "\n");
                }
            }
            in1.close();
            outputFile.close();
            System.out.println("-------------------------------    --------------------------------\n");
    	}catch (Exception e){
  		System.err.println("Hiba: " + e.getLocalizedMessage());
  	}
    }

}
