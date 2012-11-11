package oxygenoffice.extensions.smart.diagram.processes.continuousblockprocess;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;


public class ContinuousBlockProcessItem extends ProcessDiagramItem {
    
    public static int           _iShapeWidth  = 7;
    public static int           _iSpace       = 1;
    private final static int    GAP           = 50;

    ContinuousBlockProcessItem(ContinuousBlockProcess cbProcess, int id, XShape xRectangleShape) {
        super(cbProcess, id, xRectangleShape, null);
    }

    @Override
    public void setShapesProps() {
        if(xMainShape != null){
            if(getProcess().isInGruopShapes(xMainShape))
                getProcess().setShapeProperties(xMainShape);
        }
    }

    @Override
    public boolean isInjuredItem() {
        return !(getProcess().isInGruopShapes(xMainShape));
    }
    
    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePos) {
        try {
            int part = controlShapeSize.Width / (numOfItems * _iShapeWidth + _iSpace * (numOfItems - 1));
            int width = part * _iShapeWidth - GAP;
            int space = part * _iSpace;
            int xCoord = controlShapePos.X + (id - 1) * (width + space);
            int height = controlShapeSize.Height / 20 * 8;
            int yCoord = controlShapePos.Y + controlShapeSize.Height / 4 + controlShapeSize.Height / 20;       
            if(xMainShape != null) {
                xMainShape.setPosition(new Point(xCoord, yCoord));
                xMainShape.setSize(new Size(width, height));
            }
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } 
        
    }

    @Override
    public void removeShapes(){
        if(xMainShape != null){
            if(getProcess().isInGruopShapes(xMainShape))
                getProcess().removeShapeFromGroup(xMainShape);
        }
    }
    
    @Override
    public void setColor(Color oColor){
        if(getProcess() != null){
            if(oColor.isGradient()){
                getProcess().setGradient(xMainShape, oColor.getStartColor(), oColor.getEndColor());
            }else{
                getProcess().setColorOfShape(xMainShape, oColor.getColor());
            }
        }
    }
    
    @Override
    public void setLineColor(){
        if(getProcess() != null){
            getProcess().setLineColorOfShape(xMainShape);
        }
    }
}
