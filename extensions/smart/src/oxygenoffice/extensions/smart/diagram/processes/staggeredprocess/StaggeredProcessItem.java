package oxygenoffice.extensions.smart.diagram.processes.staggeredprocess;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;


public class StaggeredProcessItem extends ProcessDiagramItem {
    
    // rates of measure of rectangles height and space
    public final static int     _iShapeHeightMeasure    = 3;
    public final static int     _iSpaceMeasure          = 1;
    private final static int    GAP                     = 50;

    StaggeredProcessItem(StaggeredProcess sProcess, int id, XShape xTextShape, XShape xSecondShape) {
        super(sProcess, id, xTextShape, xSecondShape);
    }
    
    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePos, int lastRectXPos) {
        try {
            int slide, space, height, itemHeight;
            slide = space = 0;
            int width = controlShapeSize.Width - GAP;
            int xCoord = controlShapePos.X;
            height = itemHeight = controlShapeSize.Height;
            int yCoord = controlShapePos.Y;
            
            if(numOfItems > 1 && numOfItems <= 8 )
                slide = width / 2 / 7;
            if(numOfItems > 8 )
                slide = width / 2 / (numOfItems - 1);
            width -= slide * (numOfItems - 1);
            xCoord += slide * (id - 1);
            if(numOfItems > 1){
                space = controlShapeSize.Height / ( (numOfItems * _iShapeHeightMeasure) + ((numOfItems - 1) * _iSpaceMeasure) );
                height = space * _iShapeHeightMeasure;
                itemHeight = height + space;
                yCoord = controlShapePos.Y + itemHeight * (id - 1);
            }
                
            if(xMainShape != null) {
                xMainShape.setPosition(new Point(xCoord, yCoord));
                xMainShape.setSize(new Size(width, height));
            }
            if(xSecondShape != null) {
                Point a, b, c, d, e, f, g;
                a = b = c = d = e = f = g = new Point(controlShapePos.X, controlShapePos.Y);
                if(numOfItems != id) {
                    int partW, starX, partH, startY;
                    partW = width / 40;
                    starX = xCoord + partW * 35;
                    partH = height / 6;
                    startY = yCoord + 5 * partH;
                    a = new Point(starX, startY);
                    b = new Point(starX + partW * 4, startY);
                    c = new Point(starX + partW * 4, startY + space);
                    d = new Point(starX + partW * 5, startY + space);
                    e = new Point(starX + partW * 2, startY + space + 2 *partH);
                    f = new Point(starX - partW, startY + space);
                    g = new Point(starX, startY + space);
                    
                }
      
                Point[] points1 = new Point[7];
                points1[0] = new Point(a.X, a.Y);
                points1[1] = new Point(b.X, b.Y);
                points1[2] = new Point(c.X, c.Y);
                points1[3] = new Point(d.X, d.Y);
                points1[4] = new Point(e.X, e.Y);
                points1[5] = new Point(f.X, f.Y);
                points1[6] = new Point(g.X, g.Y);

                Point[] points2 = new Point[2];
                points2[0] = new Point(a.X, a.Y);
                points2[1] = new Point(g.X, g.Y);

                Point[][] allPoints = new Point[2][];
                allPoints[0] = points1;
                allPoints[1] = points2;

                XPropertySet xPolygonShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xSecondShape);
                xPolygonShapeProps.setPropertyValue("PolyPolygon", allPoints);
                
                setInvisibleFeaturesOfShape(xSecondShape, numOfItems == id);
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        
    }
    
}
