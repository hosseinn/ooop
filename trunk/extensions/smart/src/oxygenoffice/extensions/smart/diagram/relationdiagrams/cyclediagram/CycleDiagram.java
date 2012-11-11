package oxygenoffice.extensions.smart.diagram.relationdiagrams.cyclediagram;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.SchemeDefinitions;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class CycleDiagram extends RelationDiagram {
    

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_SHADOW         = 2;
    public final static short   WITH_FRAME          = 3;
    public final static short   GREEN_DARK      = 4;
    public final static short   GREEN_BRIGHT    = 5;
    public final static short   BLUE_DARK       = 6;
    public final static short   BLUE_BRIGHT     = 7;
    public final static short   PURPLE_DARK     = 8;
    public final static short   PURPLE_BRIGHT   = 9;
    public final static short   ORANGE_DARK     = 10;
    public final static short   ORANGE_BRIGHT   = 11;
    public final static short   YELLOW_DARK     = 12;
    public final static short   YELLOW_BRIGHT   = 13;
    public final static short   BLUE_SCHEME      = 14;
    public final static short   AQUA_SCHEME      = 15;
    public final static short   RED_SCHEME       = 16;
    public final static short   FIRE_SCHEME      = 17;
    public final static short   SUN_SCHEME       = 18;
    public final static short   GREEN_SCHEME     = 19;
    public final static short   OLIVE_SCHEME     = 20;
    public final static short   PURPLE_SCHEME    = 21;
    public final static short   PINK_SCHEME      = 22;
    public final static short   INDIAN_SCHEME    = 23;
    public final static short   MAROON_SCHEME    = 24;
    public final static short   BROWN_SCHEME     = 25;
    public final static short   USER_DEFINE         = 26;

    public final static short   FIRST_COLORTHEME_STYLE_VALUE = 4;
    public final static short   FIRST_COLORSCHEME_STYLE_VALUE = 14;


    public CycleDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
        setFrameProp(false);
    }

    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
    }
    
    @Override
    public boolean isColorSchemeStyle(short style){
        return  style == BLUE_SCHEME || style == AQUA_SCHEME ||
                style == RED_SCHEME || style == FIRE_SCHEME ||
                style == SUN_SCHEME || style == GREEN_SCHEME ||
                style == OLIVE_SCHEME || style == PURPLE_SCHEME ||
                style == PINK_SCHEME || style == INDIAN_SCHEME ||
                style == MAROON_SCHEME || style == BROWN_SCHEME;
    }

    @Override
    public short getColorModeOfSchemeStyle(short style){
        return (short)(style - FIRST_COLORSCHEME_STYLE_VALUE + Diagram.FIRST_COLORSCHEME_MODE_VALUE);
    }
    
    @Override
    public boolean isColorThemeStyle(short style){
        return  style == GREEN_DARK || style == GREEN_BRIGHT ||
                style == BLUE_DARK || style == BLUE_BRIGHT || 
                style == PURPLE_DARK || style == PURPLE_BRIGHT ||
                style == ORANGE_DARK || style == ORANGE_BRIGHT ||
                style == YELLOW_DARK || style == YELLOW_BRIGHT;
    }
    
    @Override
    public short getColorModeOfThemeStyle(short style){
        return (short)(style - FIRST_COLORTHEME_STYLE_VALUE + Diagram.FIRST_COLORTHEME_MODE_VALUE);
    }

    @Override
    public String getDiagramTypeName(){
        return "CycleDiagram";
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<RelationDiagramItem> items){
        super.initProperties(xControlShape, items);
        setFrameProp(false);
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            //if(getStyleProps() == VennDiagram.DEFAULT){ }
            if(getStyleProp() == CycleDiagram.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == CycleDiagram.WITH_FRAME)
                setFrameProp(true);
            if(getStyleProp() == CycleDiagram.WITH_SHADOW)
                setShadowProp(true);
            if(getStyleProp() == CycleDiagram.USER_DEFINE){
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
                            
                if(AnyConverter.toBoolean(xProps.getPropertyValue("Shadow")))
                    setShadowProp(true);
                
                xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getTextShape());
                LineStyle lineStyle = (LineStyle)xProps.getPropertyValue("LineStyle");
                if(lineStyle.getValue() == LineStyle.SOLID_value)
                    setFrameProp(true);                
            }

            setFontPropertyValues();
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, items.get(0).getTextShape());
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            setTextColorProp(AnyConverter.toInt(xTextProps.getPropertyValue("CharColor")));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (com.sun.star.lang.IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("CycleDiagram") && (shapeName.contains("ClosedBezierShape") || shapeName.contains("RectangleShape") || shapeName.contains("EllipseShape"))){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public int getTopShapeID(){

        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")) {
                    shapeID = getController().getShapeID(currShapeName);
                    if (shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return iTopShapeID;
    }

    @Override
    public XShape getPreviousShape(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("ClosedBezierShape"))
            return getPreviousItem().getMainShape();
        if(shapeName.contains("RectangleShape"))
            return getPreviousItem().getTextShape();
        return null;
    }

    @Override
    public XShape getNextShape(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("ClosedBezierShape"))
            return getNextItem().getMainShape();
        if(shapeName.contains("RectangleShape"))
            return getNextItem().getTextShape();
        return null;
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(2);
    }

    @Override
    public void createControlShape(){
        int controlEllipseSize = m_GroupSize / 4 * 3;
        Point middlePoint = new Point(m_GroupSize / 2 + m_PageProps.BorderLeft + m_iHalfDiff, m_GroupSize / 2 + m_PageProps.BorderTop);
        XShape xControlEllipseShape = createShape("EllipseShape", 0, middlePoint.X - controlEllipseSize / 2, middlePoint.Y - controlEllipseSize / 2, controlEllipseSize, controlEllipseSize);
        m_xShapes.add(xControlEllipseShape);
        setControlShapeProps(getControlShape());
    }

    @Override
    public void initDiagram(){
        super.initDiagram();
        if(!items.isEmpty())
            items.clear();
        try {
            XShape xCurrShape;
            String currShapeName = "";
            int currShapeID;
            for(int i = 0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("EllipseShape0"))
                    m_xControlShape = xCurrShape;
            }
            int topShapeID = getTopShapeID();
            for(int i = 1; i <= topShapeID; i++ ){
                for(int j = 0; j < m_xShapes.getCount(); j++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(j));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.contains("ClosedBezierShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new CycleDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID)));
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapePropsWithoutTextProps(getControlShape());
    }

    @Override
    public XShape getPairOfMainShape(int id){
        XShape xCurrShape;
        String currShapeName = "";
        int currShapeID;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.contains("RectangleShape")) {
                    currShapeID = getController().getShapeID(currShapeName);
                    if(currShapeID == id)
                        return xCurrShape;
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return null;
    }

    @Override
    public void addShape(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0)
            increaseItemsIDs(selectedShapeID);
        if(selectedShapeID == -1)
            selectedShapeID = getTopShapeID();
        if(selectedShapeID == -1 ){
            super.createDiagram();
            createDiagram(2);
        }
        if(selectedShapeID == 0 ){
            addShape(1);
            addShape(2);
        }
        if(selectedShapeID > 0){
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
        }
    }


    @Override
    public RelationDiagramItem createItem(int shapeID, String str, Color oColor){
        XShape xBezierShape = createShape("ClosedBezierShape", shapeID);
        m_xShapes.add(xBezierShape);
        setMoveProtectOfShape(xBezierShape);

        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setInvisibleFeatures(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);

        CycleDiagramItem item = new CycleDiagramItem(this, shapeID, xBezierShape, xRectangleShape);
        addItem(item);

        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        item.setLineColor();
        getController().setSelectedShape(xBezierShape);

        return item;
    }
    
    @Override
    public RelationDiagramItem createItem(int shapeID, String str){
        XShape xBezierShape = createShape("ClosedBezierShape", shapeID);
        m_xShapes.add(xBezierShape);
        setMoveProtectOfShape(xBezierShape);

        XShape xRectangleShape = createShape("RectangleShape", shapeID);
        m_xShapes.add(xRectangleShape);
        setInvisibleFeatures(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);
        
        CycleDiagramItem item = new CycleDiagramItem(this, shapeID, xBezierShape, xRectangleShape);
        addItem(item);

        if(str.equals(""))
            str = " ";
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);
        item.setShapesProps(true);
        setTextColorOfShape(item.getTextShape());
        
        XShape selectedShape = getController().getSelectedShape();
        if(selectedShape != null){
            String selectedShapeName = getShapeName(selectedShape);
            if(selectedShapeName.contains("RectangleShape"))
                getController().setSelectedShape(xRectangleShape);
            else
                getController().setSelectedShape(xBezierShape);
        }else
            getController().setSelectedShape(xBezierShape);
        
        return item;
    }

    @Override
    public XShape getControlShape(){
        if(m_xControlShape == null && m_xShapes != null){
            XShape xCurrShape = null;
            String currShapeName = "";
            try {
                for( int i=0; i < m_xShapes.getCount(); i++ ){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    if (currShapeName.endsWith("EllipseShape0"))
                        m_xControlShape = xCurrShape;
                }
            } catch (IndexOutOfBoundsException ex) {
                System.err.println(ex.getLocalizedMessage());
            } catch (WrappedTargetException ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
        return m_xControlShape;
    }

    @Override
    public void refreshDiagram() {
        removeSingleItems();
        int numOfItems = getNumOfItems();
        if(numOfItems > 0){
            //init
            Point controlShapePoint = getControlShape().getPosition();
            Size controlShapeSize = getControlShape().getSize();
            Point middlePoint = new Point(controlShapePoint.X + controlShapeSize.Width / 2, controlShapePoint.Y + controlShapeSize.Height / 2);
            int rectShapeWidth = (int)(controlShapeSize.Width / numOfItems * 1.8);
            int rectShapeHeight = controlShapeSize.Height / numOfItems;
            Point rectMiddlePoint = new Point( middlePoint.X - rectShapeWidth / 2, middlePoint.Y - rectShapeHeight / 2);
            //set positions of shapes
            double    angle = - Math.PI/2;
            int radius = controlShapeSize.Width/2 <= controlShapeSize.Height/2 ? controlShapeSize.Width/2 : controlShapeSize.Height/2;

            if(numOfItems > 4 ){
                double newRadius;
                if(numOfItems == 5)
                    newRadius = (numOfItems / 100.0 + 1.0) * radius;
                else
                    newRadius = (numOfItems / 80.0 + 1.0) * radius;
                if(numOfItems > 20 )
                    newRadius = 1.25 * radius;
                radius = (int)newRadius;
            }
            int radius2, radius3, radius4, radius5, radius6;
            radius2 = radius3 = radius4 = radius5 = radius6 = 0;

            double point1Diff, point2Diff, point3Diff, point4Diff, point5Diff, point6Diff, point7Diff, point8Diff, point9Diff;
            point1Diff = point2Diff = point3Diff = point4Diff = point5Diff =
            point6Diff = point7Diff= point8Diff = point9Diff = 0.0;

            if(numOfItems == 2){
                rectShapeWidth = (int)(rectShapeWidth * 0.5);
                rectShapeHeight = (int)(rectShapeHeight * 0.6);
                rectMiddlePoint = new Point(middlePoint.X - rectShapeWidth / 2, middlePoint.Y - rectShapeHeight / 2);
                radius2 = (int)(radius + radius / (numOfItems * 3));
                radius3 = (int)(radius + radius / (numOfItems * 1.4));
                radius4 = (int)(radius - radius / (numOfItems * 1.3));
                radius5 = (int)(radius - radius / (numOfItems * 3));
                radius6 = (int)(radius);
                double diff = Math.PI / 16 + 2 * Math.PI / numOfItems;
                point1Diff = - Math.PI / numOfItems / 2 + diff;
                point2Diff = - Math.PI / numOfItems / 4 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI / numOfItems / 3 / 20 + diff;
                point5Diff = Math.PI / numOfItems / 3 + diff;
                point6Diff = - Math.PI / numOfItems / 3 / 10 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI / numOfItems / 4 + diff;
                point9Diff =  - Math.PI / numOfItems / 2 + diff;
            }
            if(numOfItems == 3){
                rectShapeWidth = (int)(controlShapeSize.Width / numOfItems * 1.35);
                rectShapeHeight = (int)(controlShapeSize.Height / numOfItems * 0.85);
                rectMiddlePoint.X =  rectMiddlePoint.X + rectShapeWidth / 4;
                rectMiddlePoint.Y =  rectMiddlePoint.Y + rectShapeHeight / 4;
                radius2 = (int)(radius + radius / (numOfItems * 2.5));
                radius3 = (int)(radius + radius / (numOfItems * 1.2));
                radius4 = (int)(radius - radius / (numOfItems * 1.2));
                radius5 = (int)(radius - radius / (numOfItems * 2.5));
                radius6 = (int)(radius - radius / (numOfItems * 10));
                double diff = 2 * Math.PI / numOfItems;
                point1Diff = - Math.PI / numOfItems / 2 + diff;
                point2Diff = - Math.PI / numOfItems / 4 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI / numOfItems / 3 / 8 + diff;
                point5Diff = Math.PI / numOfItems / 3 + diff;
                point6Diff = - Math.PI / numOfItems / 3 / 8 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI / numOfItems / 4 + diff;
                point9Diff =  - Math.PI / numOfItems / 2 + diff;
            }
            if(numOfItems == 4){
                rectShapeWidth = (int)(controlShapeSize.Width / numOfItems * 1.6);
                rectMiddlePoint.X =  rectMiddlePoint.X + rectShapeWidth / 16;
            }
            if(numOfItems >= 4){
                radius2 = (int)(radius + radius / (numOfItems * 2));
                radius3 = (int)(radius + radius / numOfItems);
                radius4 = (int)(radius - radius / numOfItems);
                radius5 = (int)(radius - radius / (numOfItems * 2));
                radius6 = (int)(radius - radius / (numOfItems * 10));
                double diff = 2 * Math.PI / numOfItems;
                point1Diff = - Math.PI / numOfItems / 3 + diff;
                point2Diff = - Math.PI / numOfItems / 3 / 2 + diff;
                point3Diff = 0 + diff;
                point4Diff = Math.PI / numOfItems / 3 / 8 + diff;
                point5Diff = Math.PI / numOfItems / 3 + diff;
                point6Diff = - Math.PI / numOfItems / 3 / 8 + diff;
                point7Diff = 0 + diff;
                point8Diff =  - Math.PI / numOfItems / 3 / 2 + diff;
                point9Diff =  - Math.PI / numOfItems / 3 + diff;
            }

            for(RelationDiagramItem item : items){
                ((CycleDiagramItem)item).setPosition(numOfItems, controlShapeSize, controlShapePoint, middlePoint, rectMiddlePoint, rectShapeWidth, rectShapeHeight, angle, radius, radius2, radius3, radius4, radius5, radius6, point1Diff, point2Diff, point3Diff, point4Diff, point5Diff, point6Diff, point7Diff, point8Diff, point9Diff);
                if(isColorSchemeMode())
                    setGradientColor(item.getMainShape());
            }
        }
    }

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            RelationDiagramItem selectedItem = getItem(xSelectedShape);
            
            if(selectedItem != null){
                XShape xNextSelectedShape = null;
                if(getTopShapeID() > 2){
                    xNextSelectedShape = getPreviousShape();
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }else{
                    xNextSelectedShape = getControlShape();
                    RelationDiagramItem previousItem = getPreviousItem();
                    selectedItem.removeItem();
                    previousItem.removeItem();
                }
            
                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode())
                    setColorProp(getNextColor());
            }
        }
    }

    @Override
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
                        if(!shapeName.endsWith("EllipseShape0") && !shapeName.equals("")){
                            RelationDiagramItem item = getItem(xShape);
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

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, boolean isOutline, int lineWidth, boolean isShadow, boolean isFrame){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setShadowProp(isShadow);
        setFrameProp(isFrame);
    }

    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfCycleDiagram();
            if(isSelectedAllShapesProp()){
                setAllShapeProperties();
                setModifyColorsProp(false);
            } else {
                setSelectedShapesProperties();
                setModifyColorsProp(false);
                setAllShapeProperties();
            }
            
            getController().getDiagram().refreshDiagram();
        } 
        getGui().enableAndSetFocusControlDialog();
    }

    public short getAngle(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        int sector = (short)(3600 / getNumOfItems());
        int angle = sector * shapeID - 900;
        if(angle < 0)
            angle = Math.abs(angle);
        else
            angle = 3600 - angle;
        return (short)(angle % 3600);
    }

    public void setGradientColor(XShape xShape){
        int shapeID = getController().getShapeID(getShapeName(xShape));
        if(shapeID > 0){
            setSchemesColors(shapeID);
            setGradient(xShape, getAngle(xShape));
        }
    }
    
    public void setColorSettingsOfShape(XShape xShape){
            if(isColorSchemeMode()){
                int shapeID = getController().getShapeID(getShapeName(xShape));
                if(shapeID > 0)
                    setSchemesColors(shapeID);
                setGradient(xShape, getAngle(xShape));
                setLineColorProp(getDefaultLineColor());
            }else{
                if(isSimpleColorMode())
                    setLineColorProp(getDefaultLineColor());    
                if(isBaseColorsMode())
                    setLineColorProp(getDefaultLineColor());
                if(isColorThemeMode())
                    setColorThemeColors();
            
                setColorOfShape(xShape);
            }
        
        setLineColorOfShape(xShape);
    }

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        setFontPropertiesOfShape(xShape);
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);

            if(type.equals("BaseShape")){
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
            }
            
            if(type.equals("ClosedBezierShape")){
                if(isModifyColorsProp())
                    setColorSettingsOfShape(xShape);
                if(m_IsShadow){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    int numOfItems = getNumOfItems();
                    int shadowDist = SHADOW_DIST1;
                    if(numOfItems >= 4 && numOfItems <= 6)
                        shadowDist = SHADOW_DIST1 * 5 / 6;
                    if(numOfItems >= 7 && numOfItems <= 9)
                        shadowDist =  SHADOW_DIST1 * 4 / 6;
                    if(numOfItems >= 10)
                        shadowDist = SHADOW_DIST1 / 2;
                    xProp.setPropertyValue("ShadowXDistance", new Integer(shadowDist));
                    xProp.setPropertyValue("ShadowYDistance", new Integer(-shadowDist));
                    xProp.setPropertyValue("ShadowTransparence", new Integer(SHADOW_TRANSP));
                    int shadowColor = -1;
                    if(((FillStyle)xProp.getPropertyValue("FillStyle")) == FillStyle.SOLID)
                        shadowColor = AnyConverter.toInt(xProp.getPropertyValue("FillColor"));
                    else
                        shadowColor = ((Gradient)xProp.getPropertyValue("FillGradient")).EndColor;
                    if(shadowColor == -1)
                        shadowColor = 8421504;
                    xProp.setPropertyValue("ShadowColor", new Integer(shadowColor));
                }else{
                    xProp.setPropertyValue("Shadow", new Boolean(false));
                }
            
                if(m_IsOutline){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                } else {
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
            }

            if(type.equals("RectangleShape")){
                if(m_IsFrame)
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }
           
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setSchemesColors(int shapeID){
        --shapeID;
        if(isColorSchemeMode()){
            int topShapeID = getTopShapeID();
            boolean isIrr = topShapeID % 2 == 1;
            int iSteps = topShapeID / 2;
            int i1 = shapeID;
            int i2 = i1 + 1;
            if(i1 >= iSteps){
                i1 = topShapeID - shapeID;
                i2 = i1 - 1;
                if(isIrr && shapeID == iSteps)
                    i1 = i2 = iSteps;
            }
            int colorCode = getColorModeProp() - Diagram.FIRST_COLORSCHEME_MODE_VALUE;
            setStartColorProp(SchemeDefinitions.getGradientColor(colorCode, i1, iSteps));
            setEndColorProp(SchemeDefinitions.getGradientColor(colorCode, i2, iSteps));
        }
    }

    public void setInvisibleFeatures(XShape xShape){
        try {
            XPropertySet xShapeProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xShapeProps.setPropertyValue("FillColor", new Integer(0xFFFFFF));
            xShapeProps.setPropertyValue("FillTransparence", new Integer(1000));
            xShapeProps.setPropertyValue("LineStyle", LineStyle.NONE);
            xShapeProps.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

}
