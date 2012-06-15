import java.io.*;

public class RemoveEmptyRowsFromProperties {

    public static void main(String[] args) {
        if(args.length == 1 || (args.length == 2 && args[0].equals("-all"))){
            if(args[0].equals("-all")){
                File dir = new File(args[1]);
                if(!dir.exists() || !dir.isDirectory()){
                    System.out.println("Nem létező könyvtár: " + args[1] + "\n");
                    printUsage();
                }else{
                    workWithDirectiory(dir);
                }
            }else{
                File file = new File(args[0]);
                if(!file.exists() || !file.isFile()){
                    System.out.println("Nem létező fájl: " + args[0] + "\n");
                    printUsage();
                }else{
                    workWithFile(args[0]);
                }
            }
        }else{
            printUsage();
        }
    }

    public static void printUsage(){
        System.out.println("Használat:\n" +
                            "Új properties fájl létrehozása üres tulajdonságok nélkül:\n" +
                            "java RemoveEmptyRowsFromProperties fajlnev\n" +
                            "Properties fájlok másolása üres tulajdonságok nélkül egy új könyvtárban:\n" +
                            "java RemoveEmptyRowsFromProperties -all konyvtarNev");
    }

    public static void workWithDirectiory(File dir){
        try{
            System.out.println("\n-------------------  RemoveEmptyRowsFromProperties  --------------------------");
            String newDirName = dir.getName() + "_";
            File newDir = new File(newDirName);
            boolean success = newDir.mkdir();
            if (success) {
                System.out.println("Könyvtár létrehozva: " + newDirName);
                /*String[] fileList =  dir.list(new FilenameFilter(){
                                                    public boolean accept(File dir, String name) {
                                                        return name.toLowerCase().endsWith(".properties");
                                                    }
                                                });
                 */
                String[] fileList = dir.list();
                for(int i=0; i<fileList.length; i++){
                    if(fileList[i].toLowerCase().endsWith(".properties")){
                        BufferedReader inputFile = new BufferedReader(new InputStreamReader(new DataInputStream(new FileInputStream(dir + File.separator + fileList[i]))));
                        BufferedWriter outputFile = new BufferedWriter(new FileWriter(newDirName + File.separator + fileList[i]));
                        System.out.println("\n" + newDirName + File.separator + fileList[i] + " fájl létrehozva.");
                        System.out.println("Törölt sorok:");

                        String strLine, sValue;
                        while ((strLine = inputFile.readLine()) != null) {
                            int index = strLine.indexOf("=");
                            sValue = strLine.substring(index + 1);
                            if(sValue.equals("")){
                                System.out.println(strLine);
                            }else{
                                outputFile.write(strLine + "\n");
                            }
                        }
                        outputFile.close();
                        inputFile.close();
                    }
                }
            }else{
                System.out.println("Nem tudtam létrehozni a következő könyvtárat: " + newDirName);
            }
            System.out.println("------------------------------------------------------------------------------");
        }catch(Exception e){

        }
    }

    public static void workWithFile(String fileName){
        try{
            System.out.println("\n-------------------  RemoveEmptyRowsFromProperties  --------------------------");
         
            if(fileName.toLowerCase().endsWith(".properties")){
                BufferedReader inputFile = new BufferedReader(new InputStreamReader(new DataInputStream(new FileInputStream(fileName))));
                String newFileName = fileName.substring(0, fileName.indexOf(".properties")) + "_" + ".properties";
                BufferedWriter outputFile = new BufferedWriter(new FileWriter(newFileName));
                System.out.println("\n" + newFileName + " fájl létrehozva.");
                System.out.println("Törölt sorok:");

                String strLine, sValue;
                while ((strLine = inputFile.readLine()) != null) {
                    int index = strLine.indexOf("=");
                    sValue = strLine.substring(index + 1);
                    if(sValue.equals("")){
                        System.out.println(strLine);
                    }else{
                        outputFile.write(strLine + "\n");
                    }
                }
                outputFile.close();
                inputFile.close();
            }else{
                System.out.println("Nem properties kiterjesztésű a paraméterben megadott fájl: " + fileName);
            }
            System.out.println("------------------------------------------------------------------------------");
        }catch(Exception e){

        }
    }

}
