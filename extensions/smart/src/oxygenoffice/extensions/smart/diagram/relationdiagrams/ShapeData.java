package oxygenoffice.extensions.smart.diagram.relationdiagrams;


public class ShapeData {

    private int     id = -1;
    private String  text = "";

    ShapeData(int id, String text){
        this.id     = id;
        this.text   = text;
    }

    public int getID(){
        return id;
    }
    public void setID(int id){
        this.id = id;
    }
    public String getText(){
        return text;
    }
    public void setText(String text){
        this.text = text;
    }
}
