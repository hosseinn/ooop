import java.io.*;
import java.util.logging.Level;
import java.util.logging.Logger;

public class Native2ascii {

    public static void main(String[] args) {
        if(args.length != 1){
            printUsage();
        }else{
            File dir = new File(args[0]);
            if(!dir.exists() || !dir.isDirectory()){
                System.out.println("Nem létező könyvtár: " + args[0] + "\n");
                printUsage();
            }else{
                work(dir);
            }
        }
    }

    public static void printUsage(){
        System.out.println("Használat:\n" +
                            "Properties fájlok másolása és UTF8 kódolása egy új könyvtárban:\n" +
                            "java Native2ascii konyvtarnev\n");
    }

    public static void work(File dir){
        try {
            System.out.println("\n-------------------  Native2ascii  --------------------------");
            String newDirName = dir.getName() + "_";
            File newDir = new File(newDirName);
            boolean success = newDir.mkdir();
            if (success) {
                System.out.println("Könyvtár létrehozva: " + newDirName);
                String[] fileList = dir.list();
                String command = "";
                for(int i=0; i<fileList.length; i++){
                    String sOldFile = dir.getName() + File.separator + fileList[i];
                    String sNewFile = newDirName + File.separator + fileList[i].replaceAll("-", "_");
                    if(fileList[i].toLowerCase().endsWith(".properties"))
                        command = "native2ascii -encoding UTF-8 " + sOldFile + " " + sNewFile;
                    else
                        command = "cp " + sOldFile + " " + sNewFile;
                    System.out.println(command);
                    Process p = Runtime.getRuntime().exec(command);
                    p.waitFor();
                    BufferedReader reader=new BufferedReader(new InputStreamReader(p.getInputStream()));
                    String line;
                    while((line = reader.readLine()) !=null)
                        System.out.println(line);
                }
            }else{
                System.out.println("Nem tudtam létrehozni a következő könyvtárat: " + newDirName);
            }
            System.out.println("---------------------------------------------------------------");
        } catch (InterruptedException ex) {
            Logger.getLogger(Native2ascii.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException e) {
            System.err.println(e.getLocalizedMessage());
        }

    }
}