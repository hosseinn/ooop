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
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.Color;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;


public class CycleDiagram extends RelationDiagram {
    

    public final static short   DEFAULT             = 0;
    public final static short   WITHOUT_OUTLINE     = 1;
    public final static short   WITH_SHADOW         = 2;
    public final static short   WITH_FRAME          = 3;
    public final static short   BLUE_GRADIENTS      = 4;
    public final static short   RED_GRADIENTS       = 5;
    public final static short   USER_DEFINE         = 6;

    public final static int     SHADOW_DIST         = 300;
    private final static int    SHADOW_TRANSP       = 30;


    public CycleDiagram(Controller controller, Gui gui, XFrame xFrame) {
        super(controller, gui, xFrame);
        setSelectedAllShapesProps(true);
        setModifyColorsProps(false);
        setBaseColorsProps(true);
        m_IsBlueGradients = false;
        m_IsRedGradients = false;
        m_IsShadow = false;
        m_IsOutline = true;
        m_IsFrame = false;

        m_IsAction = false;
    }

    @Override
    public String getDiagramTypeName(){
        return "CycleDiagram";
    }
    
    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("CycleDiagram") && (shapeName.contains("ClosedBezierShape") || shapeName.contains("RectangleShape"))){
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
        setControlShapeProps(xControlEllipseShape);
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

            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if(currShapeName.endsWith("EllipseShape0"))
                    m_xControlShape = xCurrShape;
                if (currShapeName.contains("ClosedBezierShape")) {
                    currShapeID = getController().getShapeID(currShapeName);
                    addItem(new CycleDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID)));
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        setControlShapeProps(getControlShape());
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
            getGui().setColorModeOfImageOfControlDialog();
        }else{
            createItem(selectedShapeID + 1);
            if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
                getGui().setImageColorOfControlDialog(aCOLORS[ (selectedShapeID + 1) % 8] );
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
        getController().setSelectedShape(xBezierShape);

        return item;

    }

    @Override
    public RelationDiagramItem createItem(int shapeID, String str, int color){
        XShape xBezierShape = createShape("ClosedBezierShape", shapeID);
        m_xShapes.add(xBezierShape);
        if(m_IsBlueGradients || m_IsRedGradients){
            if(shapeID > 0)
                setGradientsStyleColors(shapeID);
            setGradient(xBezierShape, getAngle(xBezierShape));
        }else{
            setColorOfShape(xBezierShape, color);
        }
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
        item.setShapesProps();

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

    public XShape getControlShape(){
        if(m_xControlShape == null){
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
        //init
        Point controlShapePoint = getControlShape().getPosition();
        Size controlShapeSize = getControlShape().getSize();
        Point middlePoint = new Point(controlShapePoint.X + controlShapeSize.Width / 2, controlShapePoint.Y + controlShapeSize.Height / 2);
        int numOfItems = getNumOfItems();
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

        for(RelationDiagramItem item : items)
            ((CycleDiagramItem)item).setPosition(numOfItems, controlShapeSize, controlShapePoint, middlePoint, rectMiddlePoint, rectShapeWidth, rectShapeHeight, angle, radius, radius2, radius3, radius4, radius5, radius6, point1Diff, point2Diff, point3Diff, point4Diff, point5Diff, point6Diff, point7Diff, point8Diff, point9Diff);

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
                    if(isBaseColorsProps() && getGui() != null && getGui().isVisibleControlDialog())
                        getGui().setImageColorOfControlDialog(aCOLORS[0]);
                    selectedItem.removeItem();
                    previousItem.removeItem();
                }
            
                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
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

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, boolean isBaseColors, boolean isBlueGradients, boolean isRedGradients, boolean isOutline, boolean isShadow, boolean isFrame){
        setSelectedAllShapesProps(isSelectAllShape);
        setModifyColorsProps(isModifyColors);
        setBaseColorsProps(isBaseColors);
        setBlueGradientsProps(isBlueGradients);
        setRedGradientsProps(isRedGradients);
        setOutlineProps(isOutline);
        setShadowProps(isShadow);
        setFrameProps(isFrame);
        getGui().setColorModeOfImageOfControlDialog();
    }

    @Override
    public void refreshShapeProperties(){
        // need to memorize members, if user would exit into propsDialog
        boolean isSelectAllShape = isSelectedAllShapesProps();
        boolean isModifyColors = isModifyColorsProps();
        boolean isBaseColors = isBaseColorsProps();
        boolean isBlueGradients = m_IsBlueGradients;
        boolean isRedGradients = m_IsRedGradients;
        boolean isShadow =  m_IsShadow;
        boolean isOutline = m_IsOutline;
        boolean isFrame = m_IsFrame;

        m_IsAction = false;

        getGui().executePropertiesDialog();

        if(m_IsAction){

            if( m_Style == DEFAULT ){
                setPropertiesValues(true, true, true, false, false, true, false, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITHOUT_OUTLINE){
                setPropertiesValues(true, true, true, false, false, false, false, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITH_SHADOW){
                setPropertiesValues(true, true, true, false, false, true, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITH_FRAME){
                setPropertiesValues(true, true, true, false, false, true, false, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == BLUE_GRADIENTS)
                setPropertiesValues(true, true, false, true, false, true, false, false);
            if( m_Style == RED_GRADIENTS)
                setPropertiesValues(true, true, false, false, true, true, false, false);
            if(m_Style == USER_DEFINE){
                setModifyColorsProps(getGui().isModifyColorsPropsInDiagramPropsDialog());
                if(isModifyColorsProps()){
                    setBlueGradientsProps(false);
                    setRedGradientsProps(false);
                    setBaseColorsProps(getGui().isBaseColorsPropsInDiagramPropsDialog());
                    if(isBaseColorsProps()){
                        getGui().setImageColorOfControlDialog(getNextColor());
                    }else{
                        setColorProps(getGui().getImageColorOfControl(getGui().m_xColorImageControlOfPD));
                        getGui().setImageColorOfControlDialog(m_iColor);
                    }
                    getGui().setColorModeOfImageOfControlDialog();
                }
            }

            if(isSelectedAllShapesProps()){
                setAllShapeProperties();
            } else{
                setSelectedShapesProperties();
                setAllShapeFontMeausereProperties();
            }

            setModifyColorsProps(false);
        }else{
            setSelectedAllShapesProps(isSelectAllShape);
            setModifyColorsProps(isModifyColors);
            setBaseColorsProps(isBaseColors);
            m_IsBlueGradients = isBlueGradients;
            m_IsRedGradients = isRedGradients;
            m_IsShadow = isShadow;
            m_IsOutline = isOutline;
            m_IsFrame = isFrame;
        }
        m_IsAction = false;
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

    @Override
    public void setShapeProperties(XShape xShape, String type) {
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);

            if(type.equals("ClosedBezierShape")){
                if(isModifyColorsProps()){
                    int color = -1;
                    int shapeID = getController().getShapeID(getShapeName(xShape));

                    if(m_IsBlueGradients || m_IsRedGradients){
                            if(shapeID > 0)
                                setGradientsStyleColors(shapeID);
                        setGradient(xShape, getAngle(xShape));
                    }else{
                        if(isBaseColorsProps()){
                            color = aCOLORS[(shapeID - 1) % 8];
                        }else{
                            if(getGui() != null && getGui().getControlDialogWindow() != null)
                                color = getGui().getImageColorOfControlDialog();
                            if(color == -1)
                                color = m_iColor;
                        }
                        xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
                        xProp.setPropertyValue("FillColor", new Integer(color));
                    }
                }

                if(m_IsShadow){
                    xProp.setPropertyValue("Shadow", new Boolean(true));
                    int numOfItems = getNumOfItems();
                    int shadowDist = SHADOW_DIST;
                    if(numOfItems >= 4 && numOfItems <= 6)
                        shadowDist = SHADOW_DIST * 5 / 6;
                    if(numOfItems >= 7 && numOfItems <= 9)
                        shadowDist =  SHADOW_DIST * 4 / 6;
                    if(numOfItems >= 10)
                        shadowDist = SHADOW_DIST / 2;
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
            
                if(m_IsOutline)
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }

            if(type.equals("RectangleShape")){
                if(m_IsFrame)
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            }
            
            setFontPropertiesOfShape(xShape);

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

    public void setGradientsStyleColors(int shapeID){
        --shapeID;
        if(m_IsBlueGradients){
            shapeID %= 8;
            if(shapeID < 4){
                setStartColorProps(aBLUE_COLORS[shapeID]);
                setEndColorProps(aBLUE_COLORS[shapeID + 1]);
            }else{
                setStartColorProps(aBLUE_COLORS[8 - shapeID]);
                setEndColorProps(aBLUE_COLORS[8 - (shapeID + 1)]);
            }
        }
        if(m_IsRedGradients){
            shapeID %= 8;
            if(shapeID < 4){
                setStartColorProps(aRED_COLORS[shapeID]);
                setEndColorProps(aRED_COLORS[shapeID + 1]);
            }else{
                setStartColorProps(aRED_COLORS[8 - shapeID]);
                setEndColorProps(aRED_COLORS[8 - (shapeID + 1)]);
            }
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
