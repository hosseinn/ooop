package oxygenoffice.extensions.smart.diagram.processes;

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


public abstract class ProcessDiagramItem {
    
    protected ProcessDiagram   process     = null;
    protected int       id          = -1;
    protected XShape    xMainShape  = null;
    protected XShape    xSecondShape  = null;

    public ProcessDiagramItem(ProcessDiagram process, int id, XShape xMainShape, XShape xSecondShape){
        this.process    = process;
        this.id         = id;
        this.xMainShape = xMainShape;
        this.xSecondShape = xSecondShape;
    }

    public ProcessDiagram getProcess(){
        return process;
    }
    
    public Color getColorObj(){
        try {
            XShape xShape = xMainShape;
            if(this.getProcess().getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                xShape = xSecondShape;
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
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
    
    public void setShapesProps() {
        if(xMainShape != null){
            if(getProcess().isInGruopShapes(xMainShape))
                getProcess().setShapeProperties(xMainShape);
        }
        if(xSecondShape != null){
            if(getProcess().isInGruopShapes(xSecondShape))
                getProcess().setShapeProperties(xSecondShape);
        }
    }
    
    public void setShapesProps(boolean modifyColor){
        if(modifyColor){
            getProcess().setModifyColorsProp(true);
            setShapesProps();
            getProcess().setModifyColorsProp(false);
        } else {
            setShapesProps();
        }
    }
    
    public boolean isInjuredItem() {
        return !(getProcess().isInGruopShapes(xMainShape) && getProcess().isInGruopShapes(xSecondShape));
    }
    
    public int getID(){
        if(getProcess() != null)
            return getProcess().getController().getShapeID(getProcess().getShapeName(xMainShape));
        return -1;
    }
    
    public void setID(int id){
        this.id = id;
        if(xMainShape != null){
            String mainShapeName = getProcess().getShapeName(xMainShape);
            String newShapeName = mainShapeName.split("Shape", 2)[0] + "Shape" + id;
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xMainShape );
            xNamed.setName(newShapeName);
        }
        if(xSecondShape != null){
            String secondShapeName = getProcess().getShapeName(xSecondShape);
            String newShapeName = secondShapeName.split("Shape", 2)[0] + "Shape" + id;
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface( XNamed.class, xSecondShape );
            xNamed.setName(newShapeName);
        }
    }
    
    public void increaseID(){
        setID(getID() + 1);
    }

    public void decreaseID(){
        setID(getID() - 1);
    }
    
    public void removeItem(){
        if(xMainShape != null){
            if(getProcess().isInGruopShapes(xMainShape))
                getProcess().removeShapeFromGroup(xMainShape);
        }
        if(xSecondShape != null){
            if(getProcess().isInGruopShapes(xSecondShape))
                getProcess().removeShapeFromGroup(xSecondShape);
        }
        if(getProcess() != null && getProcess().items != null)
            getProcess().items.remove(this);
    }
    
    public XShape getTextShape(){
        return getMainShape();
    }
    
    public XShape getMainShape(){
        return this.xMainShape;
    }
    
    public XShape getSecondShape(){
        return this.xSecondShape;
    }
    
    public void setText(String str){
        if(xMainShape != null)
            ((XText) UnoRuntime.queryInterface(XText.class, xMainShape)).setString(str);
    }
    
    public String getText(){
        XShape xTextShape = getTextShape();
        if(xTextShape != null)
            return ((XText) UnoRuntime.queryInterface(XText.class, xTextShape)).getString();
        return "";
    }
    
    public void setInvisibleFeaturesOfShape(XShape xShape, boolean bool){
        try {
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if(bool){
                xShapeProps.setPropertyValue("FillStyle", FillStyle.NONE);
            }else{
                if(getProcess().isColorSchemeMode() || getProcess().isBaseColorsWithGradientMode() || getProcess().isColorThemeGradientMode())
                    xShapeProps.setPropertyValue("FillStyle", FillStyle.GRADIENT);
                else
                    xShapeProps.setPropertyValue("FillStyle", FillStyle.SOLID);
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void removeShapes(){
        if(xMainShape != null){
            if(getProcess().isInGruopShapes(xMainShape))
                getProcess().removeShapeFromGroup(xMainShape);
        }
        if(xSecondShape != null){
            if(getProcess().isInGruopShapes(xSecondShape))
                getProcess().removeShapeFromGroup(xSecondShape);
        }
    }
    
    public void setColor(Color oColor){
        if(getProcess() != null){
            if(oColor.isGradient()){
                getProcess().setGradient(xMainShape, oColor.getStartColor(), oColor.getEndColor());
                getProcess().setGradient(xSecondShape, oColor.getStartColor(), oColor.getEndColor());
            }else{
                getProcess().setColorOfShape(xMainShape, oColor.getColor());
                getProcess().setColorOfShape(xSecondShape, oColor.getColor());
            }
        }
    }
    
    public void setLineColor(){
        if(getProcess() != null){
            getProcess().setLineColorOfShape(xMainShape);
            getProcess().setLineColorOfShape(xSecondShape);
        }
    }
}
