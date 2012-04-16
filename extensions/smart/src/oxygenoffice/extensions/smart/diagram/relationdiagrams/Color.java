package oxygenoffice.extensions.smart.diagram.relationdiagrams;

public class Color {

    private boolean isGradient;
    private int     color;
    private int     startColor;
    private int     endColor;

    public Color(boolean isGradient, int color, int startColor, int endColor){
        this.isGradient = isGradient;
        this.color = color;
        this.startColor = startColor;
        this.endColor = endColor;
    }

    public boolean isGradient(){
        return isGradient;
    }

    public int getColor(){
        return color;
    }

    public int getStartColor(){
        return startColor;
    }

    public int getEndColor(){
        return endColor;
    }
}
