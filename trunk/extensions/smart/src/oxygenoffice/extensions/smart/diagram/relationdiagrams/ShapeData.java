package oxygenoffice.extensions.smart.diagram.relationdiagrams;


public class ShapeData {

    private int     id = -1;
    private String  text = "";
    private Color   oColor = null;

    ShapeData(int id, String text, Color oColor){
        this.id     = id;
        this.text   = text;
        this.oColor  = oColor;
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
    public Color getColor(){
        return oColor;
    }
    /*
    public void setColor(Color oColor){
        this.oColor = oColor;
    }
     */
}
