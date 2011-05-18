package oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class VennDiagramItem extends RelationDiagramItem{

    VennDiagramItem(VennDiagram vDiagram, int id, XShape xMainShape, XShape xTextShape){
        super(vDiagram, id, xMainShape, xTextShape);
        setZOrder(id);
    }

    private void setZOrder(int n){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xMainShape);
            xPropText.setPropertyValue("ZOrder", new Integer(n));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void setID(int id){
        super.setID(id);
        setZOrder(id);
    }

    public void setPosition(int numCircle, Size controlShapeSize, Point controlShapePoint, Point middlePoint, double angle, double radius, double radius2) {
        if(numCircle == 1){
            xMainShape.setPosition(new Point(controlShapePoint.X, controlShapePoint.Y));
            //may be rectangle (pair of ellipse) had been removed by user
            if(xTextShape != null)
                xTextShape.setPosition(new Point(middlePoint.X - controlShapeSize.Width / 4, controlShapePoint.Y - controlShapeSize.Height / 2));
        }else{
            //make numCircle(number of circle) angles around the middlePoint and instantate the shapes
            angle += (2.0 * Math.PI / numCircle) * (id - 1);
            int xCoord = (int)(controlShapePoint.X + radius * Math.cos(angle));
            int yCoord = (int)(controlShapePoint.Y + radius * Math.sin(angle));
            xMainShape.setPosition(new Point(xCoord, yCoord));
            if(xTextShape != null){
                xCoord = (int)(middlePoint.X + radius2 * Math.cos(angle));
                yCoord = (int)(middlePoint.Y + radius2 * Math.sin(angle));
                xTextShape.setPosition(new Point(xCoord - controlShapeSize.Width / 4, yCoord - controlShapeSize.Height / 8 ));
            }
        }
    }
    
    @Override
    public boolean isInjuredItem(){
        return !getRDiagram().isInGruopShapes(xMainShape);
    }

    @Override
    public void setShapesProps(){
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().setShapeProperties(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().setShapeProperties(xTextShape);
        }
    }

}
