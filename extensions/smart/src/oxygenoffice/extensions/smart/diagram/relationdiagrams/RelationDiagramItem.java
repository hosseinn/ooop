package oxygenoffice.extensions.smart.diagram.relationdiagrams;

import com.sun.star.awt.Gradient;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;

public abstract class RelationDiagramItem {

    protected RelationDiagram rDiagram    = null;
    protected int             id          = -1;
    protected XShape          xMainShape  = null;
    protected XShape          xTextShape  = null;

    public RelationDiagramItem(RelationDiagram rDiagram, int id, XShape xMainShape, XShape xTextShape){
        this.rDiagram = rDiagram;
        this.id         = id;
        this.xMainShape = xMainShape;
        this.xTextShape = xTextShape;
    }

    public RelationDiagram getRDiagram(){
        return rDiagram;
    }

    public XShape getMainShape(){
        return xMainShape;
    }
    
    public XShape getTextShape(){
        return xTextShape;
    }

    public Color getColorObj(){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xMainShape);
            //int lineColor = AnyConverter.toInt(xProp.getPropertyValue("LineColor"));
            if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID){
                int color = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                return new Color(false, color, 0, 0);
            } else{
                int startColor = ((Gradient)xProp.getPropertyValue("FillGradient")).StartColor;
                int endColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                return new Color(true, startColor, startColor, endColor);
            }
        }  catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }
  
    public void setColor(Color oColor){
        if(getRDiagram() != null){
            if(oColor.isGradient() && getRDiagram().getController().getDiagramType() != Controller.VENNDIAGRAM){
                getRDiagram().setGradient(xMainShape, oColor.getStartColor(), oColor.getEndColor());
                getRDiagram().setGradient(xTextShape, oColor.getStartColor(), oColor.getEndColor());
            }else{
                getRDiagram().setColorOfShape(xMainShape, oColor.getColor());
                getRDiagram().setColorOfShape(xTextShape, oColor.getColor());
            }
        }
    }
    
    public void setLineColor(){
        if(getRDiagram() != null){
            getRDiagram().setLineColorOfShape(xMainShape);
            getRDiagram().setLineColorOfShape(xTextShape);
        }
    }

    public int getID(){
        if(getRDiagram() != null)
            return getRDiagram().getController().getShapeID(getRDiagram().getShapeName(xMainShape));
        return -1;
    }

    public void setID(int id){
        this.id = id;
        if(xMainShape != null){
            String mainShapeName = getRDiagram().getShapeName(xMainShape);
            String newShapeName = mainShapeName.split("Shape", 2)[0] + "Shape" + id;
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xMainShape );
            xNamed.setName(newShapeName);
        }
        if(xTextShape != null){
            String textShapeName = getRDiagram().getShapeName(xTextShape);
            String newShapeName = textShapeName.split("Shape", 2)[0] + "Shape" + id;
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xTextShape );
            xNamed.setName(newShapeName);
        }
    }

    public void increaseID(){
        setID(getID() + 1);
    }

    public void decreaseID(){
        setID(getID() - 1);
    }

    public String getText(){
        if(xTextShape != null)
            return ((XText) UnoRuntime.queryInterface(XText.class, xTextShape)).getString();
        return "";
    }

    public void setText(String str){
        if(xTextShape != null)
            ((XText) UnoRuntime.queryInterface(XText.class, xTextShape)).setString(str);
    }

    public void setDefaultText(){
        String defaultText = getRDiagram().getGui().getDialogPropertyValue("ControlDialog1", "ControlDialog1.Text.Label");
        if(getRDiagram().getController().getDiagramType() == Controller.TARGETDIAGRAM)
            defaultText = defaultText.replaceAll("\n", " ");
        setText(defaultText);
    }

    public void removeItem(){
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().removeShapeFromGroup(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().removeShapeFromGroup(xTextShape);
        }
        if(getRDiagram() != null && getRDiagram().items != null)
            getRDiagram().items.remove(this);
    }

    public void removeShapes(){
        if(xMainShape != null){
            if(getRDiagram().isInGruopShapes(xMainShape))
                getRDiagram().removeShapeFromGroup(xMainShape);
        }
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().removeShapeFromGroup(xTextShape);
        }
    }

    public void removeTextShape(){
        if(xTextShape != null){
            if(getRDiagram().isInGruopShapes(xTextShape))
                getRDiagram().removeShapeFromGroup(xTextShape);
            //xTextShape = null;
        }
    }

    public abstract boolean isInjuredItem();

    public abstract void setShapesProps();
    
    public void setShapesProps(boolean modifyColor){
        if(modifyColor){
            getRDiagram().setModifyColorsProp(true);
            setShapesProps();
            getRDiagram().setModifyColorsProp(false);
        } else {
            setShapesProps();
        }
    }
    
    public abstract void setShapeFontMeausereProps();
    
}
