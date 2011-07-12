package oxygenoffice.extensions.smart.diagram.relationdiagrams;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.Diagram;


public abstract class RelationDiagram extends Diagram{

    protected int                               m_GroupSize     = 0;
    protected int                               m_iHalfDiff     = 0;
    protected XShape                            m_xControlShape = null;
    protected ArrayList<RelationDiagramItem>    items           = null;

    public final static short                   CORNER_RADIUS   = 300;

    public final int                            COLOR           = 255;
    public final int[]  aCOLORS         = { 65280, 255, 16711680, 16776960, 9699435, 16737843, 47359, 12076800 };
    public final int[]  aRED_COLORS     = { 11674146, 14160145, 16711680, 16744192, 16776960 };
    public final int[]  aBLUE_COLORS    = { 85, 170, 255, 7573742, 11393254 };

    public RelationDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
        items = new ArrayList<RelationDiagramItem>();
    }

    public void createDiagram(int n){
        setDrawArea();
        createControlShape();
        getController().removeSelectionListener();
        if(m_xDrawPage != null && m_xShapes != null){
            for(int i = 1; i <= n; i++)
                addShape(i);
            refreshDiagram();
        }
        getController().addSelectionListener();
    }

    public void createDiagram(ArrayList<ShapeData> shapeDatas){
        super.createDiagram();
        setDrawArea();
        createControlShape();
        getController().removeSelectionListener();
        if(m_xDrawPage != null && m_xShapes != null){
            if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
                getGui().setImageColorOfControlDialog(aCOLORS[0]);
            for(ShapeData shapeData : shapeDatas){
                createItem(shapeData.getID(), shapeData.getText());
                if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(aCOLORS[shapeData.getID() % 8]);
            }
            refreshDiagram();
        }
        getController().addSelectionListener();
    }

    public void setDrawArea(){
        try {
            m_GroupSize = m_DrawAreaWidth <= m_DrawAreaHeight ? m_DrawAreaWidth : m_DrawAreaHeight;
            m_xGroupShape.setSize( new Size( m_GroupSize, m_GroupSize ) );
            if(m_GroupSize < m_DrawAreaWidth)
                m_iHalfDiff = (m_DrawAreaWidth - m_GroupSize) / 2;
            m_xGroupShape.setPosition( new Point( m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop ) );
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public abstract void createControlShape();

    public void setControlShapeProps(XShape xShape){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillStyle", FillStyle.NONE);
            xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            xProp.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void addShape(int n){
        createItem(n);
        if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
            getGui().setImageColorOfControlDialog(aCOLORS[n % 8]);
    }

    public void removeAllShapesFromDrawPage(){
        for(RelationDiagramItem item : items)
            item.removeShapes();
        if(m_xShapes != null){
            try {
                XShape xCurrShape;
                String currShapeName = "";
                String sDiagramID = getDiagramTypeName() + getDiagramID();
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if(currShapeName.startsWith(sDiagramID))
                        m_xShapes.remove(xCurrShape);
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }

    public XShape getPairOfMainShape(int id){
        return null;
    }

    public abstract int getSelectedShapeID();

    public abstract int getTopShapeID();

    public abstract void createItem(int shapeID, String str);

    public void createItem(int shapeID){
        createItem(shapeID, "DefaultText");
    }

    public ArrayList<ShapeData> getShapeDatas(){
        ArrayList<ShapeData> shapeDatas = new ArrayList<ShapeData>();
        for(RelationDiagramItem item : items)
            shapeDatas.add(new ShapeData(item.getID(), item.getText()));
        return shapeDatas;
    }

    public int getNumOfItems(){
        return items.size();
    }

    public void addItem(RelationDiagramItem item){
        items.add(item);
    }

    public void decreaseItemsIDs(int id){
        for(RelationDiagramItem item : items)
            if(item.getID() > id)
                item.decreaseID();
    }

    public void increaseItemsIDs(int id){
        for(RelationDiagramItem item : items)
            if(item.getID() > id)
                item.increaseID();
    }

    public RelationDiagramItem getFirstItem(){
        for(RelationDiagramItem item : items)
            if(item.getID() == 1)
                return item;
        return null;
    }

    public RelationDiagramItem getItem(XShape xShape){
        for(RelationDiagramItem item : items)
            if(item.xMainShape.equals(xShape) || item.xTextShape.equals(xShape))
                return item;
        return null;
    }

    public abstract XShape getNextShape();

    public RelationDiagramItem getNextItem(){
        int id = getSelectedShapeID() + 1;
        if(id > getTopShapeID())
            return getFirstItem();
        else
            for(RelationDiagramItem item : items)
                if(item.getID() == id)
                    return item;
        return null;
    }

    public abstract XShape getPreviousShape();

    public RelationDiagramItem getPreviousItem(){
        int id = getSelectedShapeID();
        if(id > 1)
            --id;
        else
            id = getTopShapeID();
        for(RelationDiagramItem item : items)
            if(item.getID() == id)
                return item;
        return null;
    }

    public void removeSingleItems(){
        boolean isValid = false;
        while(!isValid){
            isValid = true;
            for(RelationDiagramItem item : items){
                if(item.isInjuredItem()){
                    isValid = false;
                    int id = item.getID();
                    decreaseItemsIDs(id);
                    item.removeItem();
                    break;
                }
            }
        }
    }

    public void setShapeProperties(XShape xShape) {
        if (getShapeName(xShape).contains("EllipseShape"))
            setShapeProperties(xShape, "EllipseShape");
        if (getShapeName(xShape).contains("RectangleShape"))
            setShapeProperties(xShape, "RectangleShape");
        if (getShapeName(xShape).contains("ClosedBezierShape"))
            setShapeProperties(xShape, "ClosedBezierShape");
        if (getShapeName(xShape).contains("PolyPolygonShape"))
            setShapeProperties(xShape, "PolyPolygonShape");
    }

    public abstract void setSelectedShapesProperties();

    public void setAllShapeProperties(){
        for(RelationDiagramItem item : items)
            item.setShapesProps();
    }

    public int getNextColor(){
        return aCOLORS[getSelectedShapeID() % 8];
    }

    public int getColor(int shapeID){
        int color = -1;
        if(getGui() != null && getGui().getControlDialogWindow() != null)
            color = getGui().getImageColorOfControlDialog();
        if(color < 0){
            if(isBaseColorsProps())
                color = aCOLORS[(shapeID - 1) % 8];
            else
                color = m_iColor;
        }
        return color;
    }
    
}
