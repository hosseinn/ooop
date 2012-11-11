package oxygenoffice.extensions.smart.diagram.relationdiagrams;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.container.XIndexAccess;
import com.sun.star.container.XIndexReplace;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XServiceInfo;
import com.sun.star.style.NumberingType;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.DataOfDiagram;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.ShapeData;
import oxygenoffice.extensions.smart.gui.Gui;


public abstract class RelationDiagram extends Diagram{

    protected int                               m_GroupSize     = 0;
    protected int                               m_iHalfDiff     = 0;
    protected XShape                            m_xControlShape = null;
    protected ArrayList<RelationDiagramItem>    items           = null;

   
    public RelationDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
        setDefaultProps();
        items = new ArrayList<RelationDiagramItem>();
    }
    
    @Override
    public void initColorModeAndStyle(){
        initColorModeAndStyle(getControlShape());
    }

    @Override
    public void initProperties(){
        XShape xControlShape = getControlShape();
        if(xControlShape != null && items != null && !items.isEmpty())
            initProperties(xControlShape, items);
    }
    
    public void initProperties(XShape xControlShape, ArrayList<RelationDiagramItem> items){
        setDefaultProps();
        initColorModeAndStyle();
        if(isSimpleColorMode())
            setColorProp(items.get(0).getColorObj().getColor());
        if(isColorThemeMode())
            setColorThemeColors();
        if(isBaseColorsMode() || isBaseColorsWithGradientMode()){
            for(RelationDiagramItem item : items){
                if(item.getID() > 0 && item.getID() <= 8)
                    Diagram._aBaseColors[item.getID() - 1] = item.getColorObj().getColor();
            }
            setColorProp(getNextColor());
        }
    }
    
    public void createDiagram(int n){
        setDrawArea();
        if(isBaseColorsMode() || isBaseColorsWithGradientMode())
            setColorProp(Diagram._aBaseColors[0]);
        createControlShape();
        setColorModeAndStyeOfControlShape(m_xControlShape);
        if(m_xDrawPage != null && m_xShapes != null){
            for(int i = 1; i <= n; i++)
                addShape(i);
            refreshDiagram();
        }
    }

    public void createDiagram(ArrayList<ShapeData> shapeDatas){
        super.createDiagram();
        setDrawArea();
        createControlShape();
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
    public void createDiagram(DataOfDiagram datas){
        if(!datas.isEmpty()){
            super.createDiagram();
            setDrawArea();
            if(getController().getDiagramType() == Controller.PYRAMIDDIAGRAM){
                setTextFitProp(false);
                setFontSizeProp((float)32.0);
            }
            if(isBaseColorsMode() || isBaseColorsWithGradientMode())
                setColorProp(Diagram._aBaseColors[0]);
            
            createControlShape();
            setColorModeAndStyeOfControlShape(m_xControlShape);
      
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
                
                /* with numbering
                DataOfDiagram itemData = new DataOfDiagram();
                itemData.add(datas.get(0).getLevel(), datas.get(0).getValue());
                int n = 1;
                for(int i = 1; i < datas.size(); i++){
                    if(datas.get(i).getLevel() == 0){
                        RelationDiagramItem item = createItemAndSetNextColor(n++);
                        setNumbering(item.getTextShape(), itemData);
                        itemData.clear();
                        itemData.add(datas.get(i).getLevel(), datas.get(i).getValue());
                    } else {
                        itemData.add(datas.get(i).getLevel(), datas.get(i).getValue());
                    }
                }
                RelationDiagramItem item = createItemAndSetNextColor(n);
                setNumbering(item.getTextShape(), itemData);
                */
                refreshDiagram();
            }
        }
    }
    
    public RelationDiagramItem createItemAndSetNextColor(int n, String str){
        RelationDiagramItem item = createItem(n, str);
        if(isBaseColorsMode() || isBaseColorsWithGradientMode())
            setColorProp(Diagram._aBaseColors[n % 8]);
        return item;
    }
    
    public void setNumbering(XShape xShape, DataOfDiagram datas){
        try {
            String sText = datas.get(0).getValue();
            for(int i = 1; i < datas.size(); i++){
                sText += "\n" + datas.get(i).getValue();
            }
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xShape);
            xText.setString(sText);
     
            XEnumerationAccess xEnumerationAccess = (XEnumerationAccess) UnoRuntime.queryInterface(XEnumerationAccess.class, xText);
            XEnumeration xParagraphEnumeration = xEnumerationAccess.createEnumeration();
            short m = 0;
            while (xParagraphEnumeration.hasMoreElements()) {
                Object o = xParagraphEnumeration.nextElement();
                XServiceInfo xInfo = (XServiceInfo) UnoRuntime.queryInterface(XServiceInfo.class, o);
                if (xInfo.supportsService("com.sun.star.text.Paragraph")) {
                    XPropertySet xSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xInfo);
                    short level = datas.get(m).getLevel();
                    try{
                        xSet.setPropertyValue("NumberingLevel", new Short(level));
                    }catch(Exception ex){ }
                    m++;
                }
            }    
            
            Object oNumberingRules = m_xMSF.createInstance("com.sun.star.text.NumberingRules");
            XIndexAccess iNumRules = (XIndexAccess)UnoRuntime.queryInterface(XIndexAccess.class, oNumberingRules);
            XIndexReplace xReplace = (XIndexReplace)UnoRuntime.queryInterface(XIndexReplace.class, iNumRules);
            for(int i = 0; i < iNumRules.getCount(); i++){
                PropertyValue[] aProps = (PropertyValue[]) iNumRules.getByIndex(i);
                for (int j = 0 ; j < aProps.length ; ++j) 
                    if (aProps[j].Name.equals ("NumberingType"))
                        aProps[j].Value = new Short(NumberingType.NUMBER_NONE);
                xReplace.replaceByIndex (i, aProps);
            }
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xText);
            xProps.setPropertyValue("NumberingRules", iNumRules); 
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        } 
    }

    @Override
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
    
    @Override
    public void setColorProp(int color){
        super.setColorProp(color);
        if(getGui() != null && getGui().getControlDialogWindow() != null)
            getGui().setImageColorOfControlDialog(color);
    }

    public void addShape(int n){
        createItem(n);
        if(isBaseColorsMode())
            setColorProp(Diagram._aBaseColors[n % 8]);
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

    public abstract short getUserDefineStyleValue();
    
    public abstract RelationDiagramItem createItem(int shapeID, String str, Color oColor);

    public abstract RelationDiagramItem createItem(int shapeID, String str);

    public RelationDiagramItem createItem(int shapeID){
        RelationDiagramItem item = createItem(shapeID, "DefaultText");
        return item;
    }

    public ArrayList<ShapeData> getShapeDatas(){
        ArrayList<ShapeData> shapeDatas = new ArrayList<ShapeData>();
        //have to order items by id
        int lastID = -1;
        for(RelationDiagramItem item : items)
            if(lastID < item.id)
                lastID = item.id;
        for(int i = 0; i <= lastID; i++){
            for(RelationDiagramItem item : items)
                if(item.id == i)
                    shapeDatas.add(new ShapeData(item.getID(), item.getText(), item.getColorObj()));
        }
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

    public void setShapeProperties(XShape xShape) {
        if (getShapeName(xShape).contains("Shape0")){
            setShapeProperties(xShape,"BaseShape");
        } else {
            if (getShapeName(xShape).contains("EllipseShape"))
                setShapeProperties(xShape, "EllipseShape");
            if (getShapeName(xShape).contains("RectangleShape"))
                setShapeProperties(xShape, "RectangleShape");
            if (getShapeName(xShape).contains("ClosedBezierShape"))
                setShapeProperties(xShape, "ClosedBezierShape");
            if (getShapeName(xShape).contains("PolyPolygonShape"))
                setShapeProperties(xShape, "PolyPolygonShape");
            if (getShapeName(xShape).contains("ConnectorShape"))
                setShapeProperties(xShape, "ConnectorShape");
        }
    }
    
    public abstract XShape getControlShape();

    public abstract void setSelectedShapesProperties();

    public void setAllShapeProperties(){
        setShapeProperties(getControlShape());
        if(isBaseColorsMode() && isModifyColorsProp())
            setColorProp(Diagram._aBaseColors[0]);
        for(RelationDiagramItem item : items){
            item.setShapesProps();
            if(isBaseColorsMode() && isModifyColorsProp())
                setColorProp(Diagram._aBaseColors[item.getID() % 8]);
        }
        if(isBaseColorsMode() && isModifyColorsProp()){
            int id = getSelectedShapeID();
                if(id > 0)
                    setColorProp(getNextColor());
        }
    }

    public void setAllShapeFontMeausereProperties(){
        for(RelationDiagramItem item : items)
            item.setShapeFontMeausereProps();
    }

    public int getNextColor(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID == -1)
            selectedShapeID = 0;
        return Diagram._aBaseColors[selectedShapeID % 8];
    }
    
}
