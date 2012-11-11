package oxygenoffice.extensions.smart.diagram.processes;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.TextHorizontalAdjust;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.DataOfDiagram;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.ShapeData;
import oxygenoffice.extensions.smart.diagram.processes.continuousblockprocess.ContinuousBlockProcess;
import oxygenoffice.extensions.smart.diagram.processes.upwardarrowprocess.UpwardArrowProcess;
import oxygenoffice.extensions.smart.gui.Gui;


public abstract class ProcessDiagram  extends Diagram {
    
    protected int                       m_iHalfDiff     = 0;
    protected int                       m_iGroupWidth;
    protected int                       m_iGroupHeight;
    protected XShape                    m_xControlShape = null;
    protected ArrayList<ProcessDiagramItem>    items           = null;

    public ProcessDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
        setDefaultProps();
        items = new ArrayList<ProcessDiagramItem>();
    }
    
    public void createDiagram(int n){
        setDrawArea();
        if(isBaseColorsMode() || isBaseColorsWithGradientMode())
            setColorProp(Diagram._aBaseColors[0]);
        createControlShape();
        setColorModeAndStyeOfControlShape(m_xControlShape);
        if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
            ((ContinuousBlockProcess)this).createArrowShape(true);
        if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
            ((UpwardArrowProcess)this).createArrowShape(true);
        if(m_xDrawPage != null && m_xShapes != null){
            for(int i = 1; i <= n; i++)
                addShape(i);
            refreshDiagram();
        }
    }
    
    public abstract void createControlShape();
    
    public void addShape(int n){
        createItem(n);
        if(isBaseColorsMode() || isBaseColorsWithGradientMode())
            setColorProp(Diagram._aBaseColors[n % 8]);
    }
    
    public abstract short getUserDefineStyleValue();
   
    public abstract ProcessDiagramItem createItem(int shapeID);
    
    public abstract ProcessDiagramItem createItem(int shapeID, String str);
    
    public abstract ProcessDiagramItem createItem(int shapeID, String str, Color oColor);
     
    public ArrayList<ShapeData> getShapeDatas(){
        ArrayList<ShapeData> shapeDatas = new ArrayList<ShapeData>();
        //have to order items by id
        int lastID = -1;
        for(ProcessDiagramItem item : items)
            if(lastID < item.id)
                lastID = item.id;
        for(int i = 0; i <= lastID; i++){
            for(ProcessDiagramItem item : items)
                if(item.id == i)
                    shapeDatas.add(new ShapeData(item.getID(), item.getText(), item.getColorObj()));
        }
        return shapeDatas;
    }
    
    public void createDiagram(ArrayList<ShapeData> shapeDatas){
        super.createDiagram();
        setDrawArea();
        createControlShape();
        if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
            ((ContinuousBlockProcess)this).createArrowShape(false);
        if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
            ((UpwardArrowProcess)this).createArrowShape(false);
        setColorModeAndStyeOfControlShape(m_xControlShape);
        if(m_xDrawPage != null && m_xShapes != null){
            for(ShapeData shapeData : shapeDatas)
                createItem(shapeData.getID(), shapeData.getText(), shapeData.getColor());
            refreshDiagram();
            if(isBaseColorsMode())
                setColorProp(getNextColor());
        }
    }
    
    @Override
    public void createDiagram(DataOfDiagram datas) {
        if(!datas.isEmpty()){
            super.createDiagram();
            setDrawArea();
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS || getController().getDiagramType() == Controller.UPWARDARROWPROCESS){
                setTextFitProp(false);
                setFontSizeProp((float)32.0);
            }
            if(getController().getDiagramType() == Controller.STAGGEREDPROCESS || getController().getDiagramType() == Controller.BENDINGPROCESS)    
                setRoundedProp(Diagram.NULL_ROUNDED);
            
            if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                setColorProp(Diagram._aBaseColors[0]);
        
            createControlShape();
            setColorModeAndStyeOfControlShape(m_xControlShape);
    
            if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
                ((ContinuousBlockProcess)this).createArrowShape(true);
            if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
                ((UpwardArrowProcess)this).createArrowShape(true);
            if(m_xDrawPage != null && m_xShapes != null){
                int n = 1;
                String text = datas.get(0).getValue();
                for(int i = 1; i < datas.size(); i++){
                    if(datas.get(i).getLevel() == 0){
                        createItemAndSetNextColor(n++, text);
                        text = datas.get(i).getValue();
                    } else {
                        text += "\n" + getController().getSpace(datas.get(i).getLevel()) + datas.get(i).getValue();
                    }
                }
                createItemAndSetNextColor(n, text);
                refreshDiagram();
            }
        }
    }
    
    public ProcessDiagramItem createItemAndSetNextColor(int n, String str){
        ProcessDiagramItem item = createItem(n, str);
        if(isBaseColorsMode() || isBaseColorsWithGradientMode())
            setColorProp(Diagram._aBaseColors[n % 8]);
        return item;
    }

    @Override
    public void setDrawArea() {
        try {
            //allow horizontal place for shadow properties
            m_PageProps.BorderTop += (SHADOW_DIST1 + 100);
            m_DrawAreaWidth -= ( 2 * SHADOW_DIST1 + 100);
            m_DrawAreaHeight -= (SHADOW_DIST1 + 100);

            int orignGSWidth = m_DrawAreaWidth;
            if ((m_DrawAreaWidth / m_iGroupWidth) <= (m_DrawAreaHeight / m_iGroupHeight)) {
                m_DrawAreaHeight = m_DrawAreaWidth * m_iGroupHeight / m_iGroupWidth;
            } else {
                m_DrawAreaWidth = m_DrawAreaHeight * m_iGroupWidth / m_iGroupHeight;
            }
            // set new size of m_xGroupShape for Organigram
            m_xGroupShape.setSize(new Size(m_DrawAreaWidth, m_DrawAreaHeight));
            m_iHalfDiff = 0;
            if (orignGSWidth > m_DrawAreaWidth)
                m_iHalfDiff = (orignGSWidth - m_DrawAreaWidth) / 2;
            m_iHalfDiff += SHADOW_DIST1;
            m_xGroupShape.setPosition(new Point(m_PageProps.BorderLeft + m_iHalfDiff, m_PageProps.BorderTop));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public abstract XShape getControlShape();
    
    public void addItem(ProcessDiagramItem item){
        items.add(item);
    }
    
    public void setShapeProperties(XShape xShape) {
        if (getShapeName(xShape).contains("Shape0")) {
            setShapeProperties(xShape, "BaseShape");
        } else {
            if (getShapeName(xShape).contains("EllipseShape"))
                setShapeProperties(xShape, "EllipseShape");
            if (getShapeName(xShape).contains("RectangleShape"))
                setShapeProperties(xShape, "RectangleShape");
            if (getShapeName(xShape).contains("PolyPolygonShape"))
                setShapeProperties(xShape, "PolyPolygonShape");
        }
    }
    
    public void removeSingleItems(){
        boolean isValid = false;
        while(!isValid){
            isValid = true;
            for(ProcessDiagramItem item : items){
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
    
    public void decreaseItemsIDs(int id){
        for(ProcessDiagramItem item : items)
            if(item.getID() > id)
                item.decreaseID();
    }
    
    public int getNumOfItems(){
        return items.size();
    }
    
    public abstract XShape getPairOfMainShape(int id);

    public ProcessDiagramItem getItem(XShape xShape){
        for(ProcessDiagramItem item : items){
            if(item.xMainShape.equals(xShape))
                return item;
            if(item.xSecondShape != null && item.xSecondShape.equals(xShape))
                return item;
        }
        return null;
    }
    
    public abstract int getSelectedShapeID();
    
    public void increaseItemsIDs(int id){
        for(ProcessDiagramItem item : items)
            if(item.getID() > id)
                item.increaseID();
    }
    
    public ProcessDiagramItem getFirstItem(){
        for(ProcessDiagramItem item : items)
            if(item.getID() == 1)
                return item;
        return null;
    }
    
    public abstract int getTopShapeID();
    
    public abstract XShape getNextShape();
    
    public ProcessDiagramItem getNextItem(){
        int id = getSelectedShapeID() + 1;
        if(id > getTopShapeID())
            return getFirstItem();
        else
            for(ProcessDiagramItem item : items)
                if(item.getID() == id)
                    return item;
        return null;
    }
    
    public abstract XShape getPreviousShape();

    public ProcessDiagramItem getPreviousItem(){
        int id = getSelectedShapeID();
        if(id > 1)
            --id;
        else
            id = getTopShapeID();
        for(ProcessDiagramItem item : items)
            if(item.getID() == id)
                return item;
        return null;
    }
    
    public int getNextColor(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID == -1)
            selectedShapeID = 0;
        return Diagram._aBaseColors[selectedShapeID % 8];
    }
    
    @Override
    public void setColorProp(int color){
        super.setColorProp(color);
        if(getGui() != null && getGui().getControlDialogWindow() != null)
            getGui().setImageColorOfControlDialog(color);
    }
    
    public void setLeftHorAdjustOfShape(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("TextHorizontalAdjust", TextHorizontalAdjust.LEFT);
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setAllShapeProperties(){
        setShapeProperties(getControlShape());
        if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
            setShapeProperties(((ContinuousBlockProcess)this).getArrowShape(), "PolyPolygonShape");
        if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
            setShapeProperties(((UpwardArrowProcess)this).getArrowShape(), "PolyPolygonShape");
        if((isBaseColorsMode() || isBaseColorsWithGradientMode()) && isModifyColorsProp()){
            setColorProp(Diagram._aBaseColors[0]);
            for(ProcessDiagramItem item : items){
                item.setShapesProps();
                setColorProp(Diagram._aBaseColors[item.getID() % 8]);
            }
            int id = getSelectedShapeID();
            if(id > 0)
                setColorProp(getNextColor());
        } else {
            for(ProcessDiagramItem item : items)
                item.setShapesProps();
        }
    }
    
    public void setSelectedShapesProperties(){
        XShapes xShapes = getController().getSelectedShapes();
        XShape xShape = null;
        try {
            for(int i = 0; i < xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xShapes.getByIndex(i));
                if(xShape != null){
                    String shapeName = getShapeName(xShape);
                    if(shapeName.contains("GroupShape")){
                        setAllShapeProperties();
                    }else{
                        if(!shapeName.endsWith("RectangleShape0") && !shapeName.equals("")){
                            ProcessDiagramItem item = getItem(xShape);
                            if(item != null)
                                item.setShapesProps();
                        }
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void setFontPropertyValues(){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, getFirstItem().getTextShape());
            TextFitToSizeType textFit = (TextFitToSizeType)xPropText.getPropertyValue("TextFitToSize");
            if(textFit.getValue() == TextFitToSizeType.NONE_value){
                setTextFitProp(false);
            } else{
                setTextFitProp(true);
            }
            float fontSizeValue = AnyConverter.toFloat(xPropText.getPropertyValue("CharHeight"));
            setFontSizeProp(fontSizeValue);
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void initColorModeAndStyle(){
        initColorModeAndStyle(getControlShape());
    }

    @Override
    public void initProperties() {
        XShape xControlShape = getControlShape();
        if(xControlShape != null && items != null && !items.isEmpty())
            initProperties(xControlShape, items);
    }
    
    public void initProperties(XShape xControlShape, ArrayList<ProcessDiagramItem> items){
        setDefaultProps();
        initColorModeAndStyle();
        if(isSimpleColorMode())
            setColorProp(items.get(0).getColorObj().getColor());
        if(isColorThemeMode())
            setColorThemeColors();
        if(isBaseColorsMode() || isBaseColorsWithGradientMode()){
            for(ProcessDiagramItem item : items){
                if(item.getID() > 0 && item.getID() <= 8)
                    Diagram._aBaseColors[item.getID() - 1] = item.getColorObj().getColor();
            }
            setColorProp(getNextColor());
        }
    }
    
    public void removeAllShapesFromDrawPage(){
        for(ProcessDiagramItem item : items)
            item.removeShapes();
        if(getController().getDiagramType() == Controller.CONTINUOUSBLOCKPROCESS)
            ((ContinuousBlockProcess)this).removeArrowShape();
        if(getController().getDiagramType() == Controller.UPWARDARROWPROCESS)
            ((UpwardArrowProcess)this).removeArrowShape();
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
    
}
