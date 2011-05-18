package oxygenoffice.extensions.smart.diagram.relationdiagrams.pyramiddiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class PyramidDiagramItem extends RelationDiagramItem{

    PyramidDiagramItem(PyramidDiagram pDiagram, int shapeID, XShape xTrapezeShape) {
        super(pDiagram, shapeID, xTrapezeShape, xTrapezeShape);
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
    }

    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePos) {
        if(xMainShape != null) {
            Point a = new Point();
            Point b = new Point();
            Point c = new Point();
            Point d = new Point();
            a.X = controlShapePos.X + controlShapeSize.Width / 2 - controlShapeSize.Width / numOfItems * id / 2;
            a.Y = controlShapePos.Y + controlShapeSize.Height / numOfItems * id;
            b.X = controlShapePos.X + controlShapeSize.Width / 2 + controlShapeSize.Width / numOfItems * id / 2;
            b.Y = controlShapePos.Y + controlShapeSize.Height / numOfItems * id;
            c.X = controlShapePos.X + controlShapeSize.Width / 2 - controlShapeSize.Width / numOfItems * (id-1) / 2;
            c.Y = controlShapePos.Y + controlShapeSize.Height / numOfItems * (id-1);
            d.X = controlShapePos.X + controlShapeSize.Width / 2 + controlShapeSize.Width / numOfItems * (id-1) / 2;
            d.Y = controlShapePos.Y + controlShapeSize.Height / numOfItems * (id-1);
            try {
                Point[] points1 = new Point[4];
                points1[0] = new Point(a.X, a.Y);
                points1[1] = new Point(c.X, c.Y);
                points1[2] = new Point(d.X, d.Y);
                points1[3] = new Point(b.X, b.Y);

                Point[] points2 = new Point[2];
                points2[0] = new Point(a.X, a.Y);
                points2[1] = new Point(b.X, b.Y);

                Point[][] allPoints = new Point[2][];
                allPoints[0] = points1;
                allPoints[1] = points2;

                XPropertySet xPolygonShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xMainShape);
                xPolygonShapeProps.setPropertyValue("PolyPolygon", allPoints);
            } catch (com.sun.star.beans.PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (IllegalArgumentException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (UnknownPropertyException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }




}
