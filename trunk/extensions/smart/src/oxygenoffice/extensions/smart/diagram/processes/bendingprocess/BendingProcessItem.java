package oxygenoffice.extensions.smart.diagram.processes.bendingprocess;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;


public class BendingProcessItem extends ProcessDiagramItem {
    
    // rates of measure of rectangles height and space
    private final static int    _iShapeWidthMeasure     = 5;
    private final static int    _iHSpaceMeasure         = 2;
    private final static int    _iShapeHeightMeasure    = 3;
    private final static int    _iVSpaceMeasure         = 2;
    private final static int    GAP                     = 50;

    BendingProcessItem(BendingProcess bProcess, int id, XShape xRectangleShape, XShape xPolygonShape) {
        super(bProcess, id, xRectangleShape, xPolygonShape);
    }
    
    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePos) {
        try {
            int numOfRows = numOfItems;
            if(numOfItems > 2){
                numOfRows = 1;
                do{
                    numOfRows++;
                }while(numOfRows * numOfRows < numOfItems);
            }
            int hPart = controlShapeSize.Width / (numOfRows * _iShapeWidthMeasure + (numOfRows - 1) * _iHSpaceMeasure);
            int width = hPart * _iShapeWidthMeasure - GAP;
            int hSpace = hPart * _iHSpaceMeasure;
            int xPosNum = (id - 1) % numOfRows;
            boolean reverse = (((id - 1) / numOfRows) % 2) == 1;
            if(reverse)
                xPosNum = (numOfRows - 1) - xPosNum;
            int xCoord = controlShapePos.X + (xPosNum) * (width + hSpace);
            
            int numOfFields = numOfRows;
            if(numOfRows * numOfRows - numOfItems >= numOfRows)
                numOfFields--;
            int vPart = controlShapeSize.Height / (numOfFields * _iShapeHeightMeasure + (numOfFields - 1) * _iVSpaceMeasure);
            int height = vPart * _iShapeHeightMeasure;
            int vSpace = vPart * _iVSpaceMeasure;
            if(numOfItems == 2)
                height /= 2;
            if(numOfItems >= 3 && numOfItems <= 6){
                height = height / 5 * 4;
                vSpace = vSpace / 5 * 4;
             //   yCoord += (height / 4); 
            }
            int yPosNum = (id - 1) / numOfRows;
            int yCoord = controlShapePos.Y + (yPosNum) * (height + vSpace);
            if(numOfItems == 2)
                yCoord += (height / 2); 
            if(numOfItems >= 3 && numOfItems <= 6)
                yCoord += (height *2 / 5); 
            
            
            
            if(xMainShape != null) {
                xMainShape.setPosition(new Point(xCoord, yCoord));
                xMainShape.setSize(new Size(width, height));
            }

            if(xSecondShape != null) {
                Point a, b, c, d, e, f, g;
                a = b = c = d = e = f = g = new Point(controlShapePos.X, controlShapePos.Y);
                if(numOfItems != id) {
                    int partW, starX, partH, startY;
                    partW = hSpace / 9;
                    partH = height / 20;
                    
                    if((xPosNum == 0 && reverse) || (xPosNum == numOfRows - 1 && !reverse)){
                        
                        starX = xCoord + (width - partH * 4) / 2;
                        startY = yCoord + height + (vSpace - partW * 5) / 2;
                        a = new Point(starX, startY);
                        b = new Point(starX, startY + partW * 3);
                        c = new Point(starX - partH, startY + partW * 3);
                        d = new Point(starX + partH * 2, startY + partW * 5);
                        e = new Point(starX + partH * 5, startY + partW * 3);
                        f = new Point(starX + partH * 4, startY + partW * 3);
                        g = new Point(starX + partH * 4, startY);
                    } else if(reverse){
                        starX = xCoord - partW * 2;
                        startY = yCoord + partH * 8;
                        a = new Point(starX, startY);
                        b = new Point(starX - partW * 3, startY);
                        c = new Point(starX - partW * 3, startY - partH);
                        d = new Point(starX - partW * 5, startY + partH * 2);
                        e = new Point(starX - partW * 3, startY + partH * 5);
                        f = new Point(starX - partW * 3, startY + partH * 4);
                        g = new Point(starX, startY + partH * 4);
                    }else{
                        starX = xCoord + width + partW * 2;
                        startY = yCoord + partH * 8;
                        a = new Point(starX, startY);
                        b = new Point(starX + partW * 3, startY);
                        c = new Point(starX + partW * 3, startY - partH);
                        d = new Point(starX + partW * 5, startY + partH * 2);
                        e = new Point(starX + partW * 3, startY + partH * 5);
                        f = new Point(starX + partW * 3, startY + partH * 4);
                        g = new Point(starX, startY + partH * 4);
                    }
                    
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
