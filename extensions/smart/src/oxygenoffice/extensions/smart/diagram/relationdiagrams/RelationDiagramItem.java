package oxygenoffice.extensions.smart.diagram.relationdiagrams;

import com.sun.star.container.XNamed;
import com.sun.star.drawing.XShape;
import com.sun.star.text.XText;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;

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
    
}
