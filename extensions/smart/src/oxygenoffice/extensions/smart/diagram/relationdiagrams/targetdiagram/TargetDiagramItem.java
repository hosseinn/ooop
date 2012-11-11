package oxygenoffice.extensions.smart.diagram.relationdiagrams.targetdiagram;

import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XIdentifierContainer;
import com.sun.star.drawing.Alignment;
import com.sun.star.drawing.EscapeDirection;
import com.sun.star.drawing.GluePoint2;
import com.sun.star.drawing.XGluePointsSupplier;
import com.sun.star.drawing.XShape;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;


public class TargetDiagramItem extends RelationDiagramItem{

    XShape xConnShape = null;

    TargetDiagramItem(TargetDiagram bDiagram, int shapeID, XShape xEllipseShape, XShape xRectangleShape, XShape xConnShape) {
        super(bDiagram, shapeID, xEllipseShape, xRectangleShape);
        this.xConnShape = xConnShape;
        setZOrder();
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
        if(xConnShape != null){
            if(getRDiagram().isInGruopShapes(xConnShape))
                getRDiagram().setShapeProperties(xConnShape);
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

    private void setZOrder(){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xMainShape);
            xPropText.setPropertyValue("ZOrder", new Integer(id));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void setID(int id){
        super.setID(id);
        setZOrder();
    }

    void setPosition(int numOfItems, Size controlEllipseSize, Point controlEllipsePos, Point middlePoint, Size controlRectSize, Point controlRectPos) {
        if(xMainShape != null){
            try {
                int width, height, xCoord, yCoord;
                if(((TargetDiagram)getRDiagram()).isLeftLayoutProperty()){
                    if(numOfItems < 6){
                        width = controlEllipseSize.Width / 3 + controlEllipseSize.Width    * 2 / 3 * (numOfItems - id) / 5;
                        height = controlEllipseSize.Height / 3 + controlEllipseSize.Height * 2 / 3 * (numOfItems - id) / 5;
                    }else{
                        width = controlEllipseSize.Width / 3 + controlEllipseSize.Width    * 2 / 3 * (numOfItems - id) / (numOfItems - 1);
                        height = controlEllipseSize.Height / 3 + controlEllipseSize.Height * 2 / 3 * (numOfItems - id) / (numOfItems - 1);
                    }
                    xCoord = controlEllipsePos.X;
                }else{
                    if(numOfItems == 1){
                        width = controlEllipseSize.Width / 2;
                        height = controlEllipseSize.Height / 2;
                    }else if(numOfItems == 2){
                        width = controlEllipseSize.Width / 3 * 2 / numOfItems * (numOfItems - (id - 1));
                        height = controlEllipseSize.Height / 3 * 2 / numOfItems * (numOfItems - (id - 1));
                    }else{
                        width = controlEllipseSize.Width / numOfItems * (numOfItems - (id - 1));
                        height = controlEllipseSize.Height / numOfItems * (numOfItems - (id - 1));
                    }
                    xCoord = middlePoint.X - width / 2;
                }
                yCoord = middlePoint.Y - height / 2;
                xMainShape.setPosition(new Point(xCoord, yCoord));
                xMainShape.setSize(new Size(width, height));
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        if(xTextShape != null){
            try {
                int defaultHeight = controlRectSize.Height / 4;
                int height = controlRectSize.Height / numOfItems;
                if(height > defaultHeight)
                    height = defaultHeight;
                int xCoord = controlRectPos.X;
                int yCoord = controlRectPos.Y + height * (numOfItems - id);
                xTextShape.setPosition(new Point(xCoord, yCoord));
                xTextShape.setSize(new Size(controlRectSize.Width, height));
            } catch (PropertyVetoException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        if(xConnShape != null){
            if(xMainShape != null){
                int xOffset = 0;
                if(numOfItems == 1)
                    xOffset = xMainShape.getSize().Width / 4;
                else if(numOfItems == 2)
                    xOffset = xMainShape.getSize().Width / 2 - controlEllipseSize.Width / 3 * 2 / 8;
                else
                    xOffset = xMainShape.getSize().Width / 2 - controlEllipseSize.Width / 4 / numOfItems;
                setConnectorShapePos(xOffset);
            }else{
                setConnectorShapePos(500);
            }
        }
    }

    public void setConnectorShapePos(int xOffset){
        try{
            GluePoint2  aGluePoint = new GluePoint2();
            aGluePoint.IsRelative = false;
            aGluePoint.PositionAlignment = Alignment.CENTER;
            aGluePoint.Escape = EscapeDirection.SMART;
            aGluePoint.IsUserDefined = true;
            aGluePoint.Position.X = xOffset;
            aGluePoint.Position.Y = 0;

            XGluePointsSupplier xGluePointsSupplier = (XGluePointsSupplier)UnoRuntime.queryInterface(XGluePointsSupplier.class, xMainShape);
            XIdentifierContainer xIdentifierContainer = (XIdentifierContainer)UnoRuntime.queryInterface(XIdentifierContainer.class, xGluePointsSupplier.getGluePoints());
            int nIndexOfGluePoint = xIdentifierContainer.insert(aGluePoint);

            XPropertySet xPropSet = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
            xPropSet.setPropertyValue("StartGluePointIndex", new Integer(nIndexOfGluePoint));
        }catch(Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void removeItem(){
        if(xConnShape != null){
            if(getRDiagram().isInGruopShapes(xConnShape))
                getRDiagram().removeShapeFromGroup(xConnShape);
        }
        super.removeItem();
    }

    @Override
    public void removeShapes(){
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().removeShapeFromGroup(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().removeShapeFromGroup(xTextShape);
        }
        if(xConnShape != null){
            if(getRDiagram().isInGruopShapes(xConnShape))
                getRDiagram().removeShapeFromGroup(xConnShape);
        }
    }

    @Override
    public void setColor(Color oColor){
        if(getRDiagram() != null){
            if(oColor.isGradient()){
                getRDiagram().setGradient(xMainShape, GradientStyle.RADIAL, oColor.getStartColor(), oColor.getEndColor(), (short)0, (short)0, (short)50, (short)50, (short)100, (short)100);
                //getRDiagram().setGradient(xTextShape, GradientStyle.RECT, oColor.getStartColor(), oColor.getEndColor(), (short)0, (short)10, (short)50, (short)50, (short)100, (short)100);
                ((TargetDiagram)getRDiagram()).setInvisibleFeatures(xTextShape);
                ((TargetDiagram)getRDiagram()).setTextColorOfShape(xTextShape, oColor.getStartColor());
            }else{
                getRDiagram().setColorOfShape(xMainShape, oColor.getColor());
                ((TargetDiagram)getRDiagram()).setInvisibleFeatures(xTextShape);
                ((TargetDiagram)getRDiagram()).setTextColorOfShape(xTextShape, oColor.getColor());
            }
        }
    }

}
