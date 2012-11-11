

import java.io.File;


public class ChangeStringInFileNames {
    public static void main(String[] args) {
        if(args.length == 3){
            File dir = new File(args[2]);
            if(dir != null && dir.isDirectory()){
                changeStrings(args[0], args[1], dir);
            }else{
                System.out.println("Nem létezik, vagy nem könyvtár: " + args[2]);
                printUsing();
            }
        }else{
            System.out.println("Három paramétert adjon meg!");
            printUsing();
        }
    }
    
    public static void printUsing(){
        System.out.println("A program átnevezi a paraméterként megadott könyvtárban lévő fájlokat a megadott paraméterek alapján.\n"
        + "Használat:\n"
        + "java ChangeStringInFileNames oldString newString dir");
    }
    
    public static void changeStrings(String oldString, String newString, File dir){
        File[] files = dir.listFiles();
        String oldName = "";
        String newName = "";
        for(File f : files){
            if(f.getName().contains(oldString)){
                oldName = f.getName();
                newName = dir.getAbsolutePath() + File.separator + oldName.replace(oldString, newString);
                f.renameTo(new File(newName));
                System.out.println("Rename: " + dir.getName() + File.separator + oldName + " to: " + dir.getName() + File.separator + newString);
            }
        }
    }
}
