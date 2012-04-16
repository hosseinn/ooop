package oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.PolyPolygonBezierCoords;
import com.sun.star.drawing.PolygonFlags;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class CycleDiagramItem extends RelationDiagramItem{

    public CycleDiagramItem(CycleDiagram cDiagram, int shapeID, XShape xBezierShape, XShape xRectangleShape) {
        super(cDiagram, shapeID, xBezierShape, xRectangleShape);
    }

    @Override
    public boolean isInjuredItem() {
        return !(getRDiagram().isInGruopShapes(xMainShape) && getRDiagram().isInGruopShapes(xTextShape));
    }

    @Override
    public void setShapesProps() {
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().setShapeProperties(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().setShapeProperties(xTextShape);
        }
    }

    @Override
    public void setShapeFontMeausereProps(){
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().setFontPropertiesOfShape(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().setFontPropertiesOfShape(xTextShape);
        }
    }
    
    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePoint, Point middlePoint, Point rectMiddlePoint, int rectShapeWidth, int rectShapeHeight, double angle, int radius, int radius2, int radius3, int radius4, int radius5, int radius6, double point1Diff, double point2Diff, double point3Diff, double point4Diff, double point5Diff, double point6Diff, double point7Diff, double point8Diff, double point9Diff) {
        
        angle += (2.0 * Math.PI / numOfItems) * (id - 1);
            //XShape xRectShape = null;
            //int xCoord;
            //int yCoord;
        if(xMainShape != null){
            Point point1 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point1Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point1Diff)));
            Point point2 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point2Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point2Diff)));
            Point point3 = new Point((int)(middlePoint.X + radius2 * Math.cos(angle + point3Diff)),(int)(middlePoint.Y + radius2 * Math.sin(angle + point3Diff)));
            Point point4 = new Point((int)(middlePoint.X + radius3 * Math.cos(angle + point4Diff)),(int)(middlePoint.Y + radius3 * Math.sin(angle + point4Diff)));
            Point point5 = new Point((int)(middlePoint.X + radius6 * Math.cos(angle + point5Diff)),(int)(middlePoint.Y + radius6 * Math.sin(angle + point5Diff)));
            Point point6 = new Point((int)(middlePoint.X + radius4 * Math.cos(angle + point6Diff)),(int)(middlePoint.Y + radius4 * Math.sin(angle + point6Diff)));
            Point point7 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point7Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point7Diff)));
            Point point8 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point8Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point8Diff)));
            Point point9 = new Point((int)(middlePoint.X + radius5 * Math.cos(angle + point9Diff)),(int)(middlePoint.Y + radius5 * Math.sin(angle + point9Diff)));
            setBezierShapeSize(xMainShape, point1, point2, point3, point4, point5 ,point6 ,point7, point8, point9);
        }

        if(xTextShape != null){
            try {
                double rectRadius = 0.0;
                if (numOfItems == 2)
                    rectRadius = radius * 0.8;
                else
                    rectRadius = radius;
                int xCoord = (int) (rectMiddlePoint.X + rectRadius * Math.cos(angle + Math.PI/numOfItems));
                int yCoord = (int) (rectMiddlePoint.Y + rectRadius * Math.sin(angle + Math.PI/numOfItems));
                xTextShape.setPosition(new Point(xCoord, yCoord));
                xTextShape.setSize(new Size(rectShapeWidth, rectShapeHeight));
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    // need to pass bezierShape and its 9 points
    public void setBezierShapeSize(XShape xBezierShape, Point point1, Point point2, Point point3, Point point4, Point point5, Point point6, Point point7, Point point8, Point point9 ){
        try {

            PolyPolygonBezierCoords aCoords = new PolyPolygonBezierCoords();

            Point[] pointCoords = new Point[11];
            pointCoords[0] = point1;
            pointCoords[1] = point2;
            pointCoords[2] = point2;
            pointCoords[3] = point3;
            pointCoords[4] = point4;
            pointCoords[5] = point5;
            pointCoords[6] = point6;
            pointCoords[7] = point7;
            pointCoords[8] = point8;
            pointCoords[9] = point8;
            pointCoords[10] = point9;

            Point[][] points = new Point[1][];
            points[0] = pointCoords;
            aCoords.Coordinates = points;

            PolygonFlags[] flags = new PolygonFlags[11];
            flags[0] = PolygonFlags.NORMAL;
            flags[1] = PolygonFlags.CONTROL;
            flags[2] = PolygonFlags.CONTROL;
            flags[3] = PolygonFlags.NORMAL;
            flags[4] = PolygonFlags.NORMAL;
            flags[5] = PolygonFlags.NORMAL;
            flags[6] = PolygonFlags.NORMAL;
            flags[7] = PolygonFlags.NORMAL;
            flags[8] = PolygonFlags.CONTROL;
            flags[9] = PolygonFlags.CONTROL;
            flags[10] = PolygonFlags.NORMAL;

            PolygonFlags[][] flagsArray = new PolygonFlags[1][];
            flagsArray[0] = flags;
            aCoords.Flags = flagsArray;

            XPropertySet xBezierShapeProps = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xBezierShape);
            xBezierShapeProps.setPropertyValue("PolyPolygonBezier", aCoords);

        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

}
