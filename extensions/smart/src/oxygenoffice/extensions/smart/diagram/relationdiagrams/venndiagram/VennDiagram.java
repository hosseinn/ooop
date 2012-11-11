package oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
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
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.gui.Gui;


public class VennDiagram extends RelationDiagram {


    public final static short   DEFAULT         = 0;
    public final static short   WITHOUT_OUTLINE = 1;
    public final static short   WITHOUT_FRAME   = 2;
    public final static short   NOT_ROUNDED     = 3;
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
    public final static short   USER_DEFINE    = 14;
    
    public final static short   FIRST_COLORTHEME_STYLE_VALUE = 4;
   
    
    public VennDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
        setDefaultProps();
        setColorModeProp(Diagram.BASE_COLORS_MODE);
    }
    
    @Override
    public short getUserDefineStyleValue(){
        return USER_DEFINE;
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
        return "VennDiagram";
    }
    
    @Override
    public void initProperties(XShape xControlShape, ArrayList<RelationDiagramItem> items){
        super.initProperties(xControlShape, items);
        XPropertySet xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getMainShape());
        try {
            //if(getStyleProps() == VennDiagram.DEFAULT){ }
            if(getStyleProp() == VennDiagram.WITHOUT_OUTLINE)
                setOutlineProp(false);
            if(getStyleProp() == VennDiagram.WITHOUT_FRAME)
                setFrameProp(false);
            if(getStyleProp() == VennDiagram.NOT_ROUNDED)
                setRoundedFrameProp(true);
            if(getStyleProp() == VennDiagram.USER_DEFINE){
                
                int transparence = AnyConverter.toInt(xProps.getPropertyValue("FillTransparence"));
                if(transparence <= 10)
                    setTransparencyProp(Diagram.NULL_TRANSP);
                if(transparence > 10 && transparence < 40)
                    setTransparencyProp(Diagram.MEDIUM_TRANSP);
                if(transparence >= 40)
                    setTransparencyProp(Diagram.EXTRA_TRANSP);
             
                if(((LineStyle)xProps.getPropertyValue("LineStyle")).getValue() == LineStyle.NONE_value)
                    setOutlineProp(false);
                setShapesLineWidthProp(AnyConverter.toInt(xProps.getPropertyValue("LineWidth")));
                
                xProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, items.get(0).getTextShape());
                transparence = AnyConverter.toInt(xProps.getPropertyValue("FillTransparence"));
                setFrameProp(transparence != 100);
                if(isFrameProp()){
                    int cornerRadius = AnyConverter.toInt(xProps.getPropertyValue("CornerRadius"));
                    setRoundedFrameProp(cornerRadius > 0);
                }
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
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }

    @Override
    public void createControlShape(){
        int ellipseSize = m_GroupSize / 3;
        Point middlePoint = new Point(m_GroupSize / 2 + m_PageProps.BorderLeft + m_iHalfDiff, m_GroupSize / 2 + m_PageProps.BorderTop);
        //ControlEllipse
        Point coord = new Point(middlePoint.X - ellipseSize / 2, middlePoint.Y - ellipseSize / 2);
        Size size = new Size(ellipseSize, ellipseSize);
        m_xControlShape = createShape("EllipseShape", 0, coord.X, coord.Y, size.Width, size.Height);
        m_xShapes.add(m_xControlShape);
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
                    if (currShapeName.contains("EllipseShape")) {
                        currShapeID = getController().getShapeID(currShapeName);
                        if(currShapeID == i)
                            addItem(new VennDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID)));
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
    public int getTopShapeID(){
        XShape xCurrShape = null;
        String currShapeName = "";
        int shapeID = -1;
        int iTopShapeID = -1;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("EllipseShape")) {
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
    public void addShape(){
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0)
            increaseItemsIDs(selectedShapeID);
        if(selectedShapeID == -1)
            selectedShapeID = getTopShapeID();
        if(selectedShapeID == -1 ){
            super.createDiagram();
            createDiagram(1);
        }else{
            createItem(selectedShapeID + 1);
            if(isBaseColorsMode())
                setColorProp(_aBaseColors[(selectedShapeID + 1) % 8]);
        }
    }

    @Override
    public int getSelectedShapeID(){
        XShape xShape = getController().getSelectedShape();
        if(xShape != null){
            String shapeName = getShapeName(xShape);
            if(shapeName.contains("VennDiagram") && (shapeName.contains("EllipseShape") || shapeName.contains("RectangleShape"))){
                int currShapeID = getController().getShapeID(shapeName);
                if(currShapeID >= 0)
                    return currShapeID;
            }
        }
        return -1;
    }

    @Override
    public RelationDiagramItem createItem(int shapeID, String str, Color oColor){
        if(oColor.isGradient())
            setColorProp(oColor.getStartColor());
        Size size = getControlShape().getSize();
        XShape xEllipseShape = createShape("EllipseShape", shapeID, size.Width, size.Height);
        m_xShapes.add(xEllipseShape);
        setMoveProtectOfShape(xEllipseShape);

        XShape xRectangleShape = createShape( "RectangleShape", shapeID, size.Width/2, size.Height/4 );
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);

        VennDiagramItem item = new VennDiagramItem(this, shapeID, xEllipseShape, xRectangleShape);
        addItem(item);
        item.setText(str);
        item.setShapesProps();
        item.setColor(oColor);
        item.setLineColor();
        getController().setSelectedShape((Object)xEllipseShape);

        return item;
    }
    
    @Override
    public RelationDiagramItem createItem(int shapeID, String str){
        Size size = getControlShape().getSize();
        XShape xEllipseShape = createShape("EllipseShape", shapeID, size.Width, size.Height);
        m_xShapes.add(xEllipseShape);
        setMoveProtectOfShape(xEllipseShape);

        XShape xRectangleShape = createShape( "RectangleShape", shapeID, size.Width/2, size.Height/4 );
        m_xShapes.add(xRectangleShape);
        setMoveProtectOfShape(xRectangleShape);

        VennDiagramItem item = new VennDiagramItem(this, shapeID, xEllipseShape, xRectangleShape);
        addItem(item);

        if(str.equals(""))
            str = " ";
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);

        item.setShapesProps(true);
        setTextColorOfShape(item.getTextShape());
        getController().setSelectedShape((Object)xEllipseShape);

        return item;
    }

    @Override
    public void refreshDiagram() {
        removeSingleItems();
        //init
        Point controlShapePoint = getControlShape().getPosition();
        Size controlShapeSize = getControlShape().getSize();
        Point middlePoint = new Point( controlShapePoint.X + controlShapeSize.Width / 2, controlShapePoint.Y + controlShapeSize.Height / 2);
        int numCircle = getNumOfItems();
        double angle = 0.0;
        if(numCircle == 3)
            angle = -Math.PI/2;
        int radius = controlShapeSize.Width/3 <= controlShapeSize.Height/3 ? controlShapeSize.Width/3 : controlShapeSize.Height/3;
        int radius2 = controlShapeSize.Width <= controlShapeSize.Height ? controlShapeSize.Width : controlShapeSize.Height;
        radius2 *= 1.2;
        for(RelationDiagramItem item : items)
            ((VennDiagramItem)item).setPosition(numCircle, controlShapeSize, controlShapePoint, middlePoint, angle, radius, radius2);
    }

    @Override
    public void removeShape() {
        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0){
            XShape xSelectedShape = getController().getSelectedShape();
            RelationDiagramItem selectedItem = getItem(xSelectedShape);

            if(selectedItem != null){
                String selectedShapeName = getShapeName(xSelectedShape);
                XShape xNextSelectedShape = getPreviousShape();

                if(selectedShapeName.contains("RectangleShape")){
                    if(xNextSelectedShape.equals(selectedItem.getTextShape()))
                        xNextSelectedShape = selectedItem.getMainShape();
                    selectedItem.removeTextShape();
                }
        
                if(selectedShapeName.contains("EllipseShape") && !selectedShapeName.endsWith("EllipseShape0")){
                    if(selectedShapeID == 1 && getTopShapeID() == 1)
                        xNextSelectedShape = getControlShape();
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }

                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
                if(isBaseColorsMode())
                    setColorProp(getNextColor());
            }
        }
    }

    @Override
    public XShape getPreviousShape(){
        return getAppropriateShapeInItem(getPreviousItem());
    }

    @Override
    public XShape getNextShape(){
        return getAppropriateShapeInItem(getNextItem());
    }

    public XShape getAppropriateShapeInItem(RelationDiagramItem item){
        XShape xShape = null;
        if(item != null){
            String selectedShapeName = getController().getDiagram().getShapeName(getController().getSelectedShape());
            if(selectedShapeName.contains("EllipseShape") && !selectedShapeName.contains("EllipseShape0"))
                xShape = item.getMainShape();
            if(selectedShapeName.contains("RectangleShape")){
                xShape = item.getTextShape();
                if(xShape == null || !getController().getDiagram().isInGruopShapes(xShape))
                    xShape = item.getMainShape();
            }
        }
        return xShape;
    }

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, short sColorMode, short sTransparency, boolean isOutline, int lineWidth, boolean isFrame, boolean isRoundedFrame){
        setSelectedAllShapesProp(isSelectAllShape);
        setModifyColorsProp(isModifyColors);
        setColorModeProp(sColorMode);
        setTransparencyProp(sTransparency);
        setOutlineProp(isOutline);
        setShapesLineWidthProp(lineWidth);
        setFrameProp(isFrame);
        setRoundedFrameProp(isRoundedFrame);
    }

    
    @Override
    public void showPropertyDialog(){
        getGui().enableControlDialogWindow(false);
        short exec = getGui().executePropertiesDialog();
        if(exec == 1){  
            getGui().setPropertiesOfVennDiagram();
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
    
    public void setColorSettingsOfShape(XShape xShape){
        if(isSimpleColorMode())
            setLineColorProp(getDefaultLineColor());
        if(isBaseColorsMode())
            setLineColorProp(getDefaultLineColor());
        if(isColorThemeMode())
            setColorThemeColors();
        
        setColorOfShape(xShape);
        setLineColorOfShape(xShape);
    }
    
    @Override
    public void setShapeProperties(XShape xShape, String type) {
        setFontPropertiesOfShape(xShape);
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if(type.equals("BaseShape"))
                if(!isTextFitProp())
                    xProp.setPropertyValue("CharHeight", new Float(40.0));
            
            if(!type.equals("BaseShape")){
                if(isModifyColorsProp())
                    setColorSettingsOfShape(xShape);
                
                if(m_sTransparency == Diagram.NULL_TRANSP)
                    xProp.setPropertyValue("FillTransparence", new Integer(0));
                if(m_sTransparency == Diagram.MEDIUM_TRANSP)
                    xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE1));
                if(m_sTransparency == Diagram.EXTRA_TRANSP)
                    xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE2));
/*
                if(m_sTransparency == Diagram.NULL_TRANSP)
                    xProp.setPropertyValue("LineTransparence", new Integer(0));
                if(m_sTransparency == Diagram.MEDIUM_TRANSP)
                    xProp.setPropertyValue("LineTransparence", new Integer(TRANSP_VALUE1));
                if(m_sTransparency == Diagram.EXTRA_TRANSP)
                    xProp.setPropertyValue("LineTransparence", new Integer(TRANSP_VALUE2));
*/                
                if(isOutlineProp()){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineWidth", new Integer(getShapesLineWidhtProp()));
                }else{
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }
                
                if(type.equals("RectangleShape")){
                    if(isFrameProp()){
                        if(isRoundedFrameProp()){
                            xProp.setPropertyValue("CornerRadius", new Integer(RelationDiagram.CORNER_RADIUS1));
                        }else{
                            xProp.setPropertyValue("CornerRadius", new Integer(0));
                        }
                    }else{
                        xProp.setPropertyValue("FillTransparence", new Integer(100));
                        xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                    }
                }
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
}
