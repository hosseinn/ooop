package oxygenoffice.extensions.smart.diagram;

// PageProps struct stores the XDrawPage properties
public class PageProps {

    public int Width;
    public int Height;
    public int BorderLeft;
    public int BorderRight;
    public int BorderTop;
    public int BorderBottom;

    PageProps(int Width, int Height, int BorderLeft, int BorderRight, int BorderTop, int BorderBottom){
        this.Width = Width;
        this.Height = Height;
        this.BorderLeft = BorderLeft;
        this.BorderRight = BorderRight;
        this.BorderTop = BorderTop;
        this.BorderBottom = BorderBottom;
    }
}
