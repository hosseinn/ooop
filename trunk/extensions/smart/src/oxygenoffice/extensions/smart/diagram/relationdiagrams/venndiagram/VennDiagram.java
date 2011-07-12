package oxygenoffice.extensions.smart.diagram.relationdiagrams.venndiagram;

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
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.Gui;
import oxygenoffice.extensions.smart.diagram.Diagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagramItem;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;


public class VennDiagram extends RelationDiagram {


    public final static short   DEFAULT          = 0;
    public final static short   WITHOUT_OUTLINE  = 1;
    public final static short   WITHOUT_FRAME    = 2;
    public final static short   NOT_ROUNDED      = 3;
    public final static short   USER_DEFINE      = 4;

    private final static int    TRANSP_VALUE     = 35;
    private final static int    TRANSP_VALUE2    = 50;


    public VennDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
        //default values
        setSelectedAllShapesProps(true);
        setModifyColorsProps(false);
        setBaseColorsProps(true);
        m_sTransparency = Diagram.MEDIUM_TRANSP;
        m_IsOutline = true;
        m_IsFrame = true;
        m_IsRoundedFrame = true;

        m_IsAction = false;
    }

    @Override
    public String getDiagramTypeName(){
        return "VennDiagram";
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
        getController().setSelectedShape(m_xControlShape);
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
                if (currShapeName.contains("EllipseShape")) {
                    if(currShapeName.endsWith("EllipseShape0")){
                        m_xControlShape = xCurrShape;
                    }else{
                        currShapeID = getController().getShapeID(currShapeName);
                        addItem(new VennDiagramItem(this, currShapeID, xCurrShape, getPairOfMainShape(currShapeID)));
                    }
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
    public void addShape(){

        int selectedShapeID = getSelectedShapeID();
        if(selectedShapeID > 0)
            increaseItemsIDs(selectedShapeID);
        if(selectedShapeID == -1)
            selectedShapeID = getTopShapeID();
        if(selectedShapeID == -1 ){
            super.createDiagram();
            createDiagram(1);
            getGui().setColorModeOfImageOfControlDialog();
        }else{
            createItem(selectedShapeID + 1);
            if(isBaseColorsProps() && getGui() != null && getGui().getControlDialogWindow() != null)
                getGui().setImageColorOfControlDialog(aCOLORS[ (selectedShapeID + 1) % 8] );
        }
    }

    @Override
    public int getSelectedShapeID(){
        String shapeName = getShapeName(getController().getSelectedShape());
        if(shapeName.contains("VennDiagram") && (shapeName.contains("EllipseShape") || shapeName.contains("RectangleShape"))){
            int currShapeID = getController().getShapeID(shapeName);
            if(currShapeID >= 0)
                return currShapeID;
        }
        return -1;
    }

    @Override
    public void createItem(int shapeID, String str){
        Size size = getControlShape().getSize();
        int color = getColor(shapeID);
        XShape xEllipseShape = createShape("EllipseShape", shapeID, size.Width, size.Height);
        m_xShapes.add(xEllipseShape);
        setColorOfShape(xEllipseShape, color);
        setMoveProtectOfShape(xEllipseShape);

        XShape xRectangleShape = createShape( "RectangleShape", shapeID, size.Width/2, size.Height/4 );
        m_xShapes.add(xRectangleShape);
        setColorOfShape(xRectangleShape, color);
        setMoveProtectOfShape(xRectangleShape);
        setTextFitToSize(xRectangleShape);

        VennDiagramItem item = new VennDiagramItem(this, shapeID, xEllipseShape, xRectangleShape);
        addItem(item);
        item.setShapesProps();
        if(str.equals("DefaultText"))
            item.setDefaultText();
        else
            item.setText(str);
        
        getController().setSelectedShape((Object)xEllipseShape);
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
                    if(selectedShapeID == 1 && getTopShapeID() == 1){
                        xNextSelectedShape = getControlShape();
                        if(isBaseColorsProps() && getGui() != null && getGui().isVisibleControlDialog())
                            getGui().setImageColorOfControlDialog(aCOLORS[0]);
                    }
                    selectedItem.removeItem();
                    decreaseItemsIDs(selectedShapeID);
                }

                if(xNextSelectedShape != null)
                    getController().setSelectedShape(xNextSelectedShape);
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

    public void setPropertiesValues(boolean isSelectAllShape, boolean isModifyColors, boolean isBaseColors, short sTransparency, boolean isOutline, boolean isFrame, boolean isRoundedFrame){
        setSelectedAllShapesProps(isSelectAllShape);
        setModifyColorsProps(isModifyColors);
        setBaseColorsProps(isBaseColors);
        setTransparencyProps(sTransparency);
        setOutlineProps(isOutline);
        setFrameProps(isFrame);
        setRoundedFrameProps(isRoundedFrame);
        getGui().setColorModeOfImageOfControlDialog();
    }

    @Override
    public void refreshShapeProperties(){
        // need to memorize members, if user would exit into propsDialog
        boolean isSelectAllShape = isSelectedAllShapesProps();
        boolean isModifyColors = isModifyColorsProps();
        boolean isBaseColors = isBaseColorsProps();
        short   sTransparency = m_sTransparency;
        boolean isOutline = m_IsOutline;
        boolean isFrame = m_IsFrame;
        boolean isRoundedFrame = m_IsRoundedFrame;
            
        m_IsAction = false;

        getGui().executePropertiesDialog();

        if(m_IsAction){

            if( m_Style == DEFAULT ){
                setPropertiesValues(true, true, true, (short)1, true, true, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITHOUT_OUTLINE){
                setPropertiesValues(true, true, true, (short)1, false, true, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == WITHOUT_FRAME){
                setPropertiesValues(true, true, true, (short)1, true, false, true);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if( m_Style == NOT_ROUNDED){
                setPropertiesValues(true, true, true, (short)1, true, true, false);
                getGui().setImageColorOfControlDialog(getNextColor());
            }
            if(m_Style == USER_DEFINE){
                setModifyColorsProps(getGui().isModifyColorsPropsInDiagramPropsDialog());
                if(isModifyColorsProps()){
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

            if(!isSelectedAllShapesProps()){
                m_IsOutline = isOutline;
                m_IsFrame = isFrame;
                m_IsRoundedFrame = isRoundedFrame;
                setSelectedShapesProperties();
            }else{
                if(!m_IsFrame)
                    m_IsRoundedFrame = isRoundedFrame;
                setAllShapeProperties();
            }
            setModifyColorsProps(false);
        }else{
            setSelectedAllShapesProps(isSelectAllShape);
            setModifyColorsProps(isModifyColors);
            setBaseColorsProps(isBaseColors);
            m_sTransparency = sTransparency;
            m_IsOutline = isOutline;
            m_IsFrame = isFrame;
            m_IsRoundedFrame = isRoundedFrame;
        }
        m_IsAction = false;
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

    @Override
    public void setShapeProperties(XShape xShape, String type) {

        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            
            if(isModifyColorsProps()){
                int color = -1;
                int shapeID = getController().getShapeID(getShapeName(xShape));
              
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

            if(m_sTransparency == Diagram.NULL_TRANSP)
                xProp.setPropertyValue("FillTransparence", new Integer(0));
            if(m_sTransparency == Diagram.MEDIUM_TRANSP)
                xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE));
            if(m_sTransparency == Diagram.EXTRA_TRANSP)
                xProp.setPropertyValue("FillTransparence", new Integer(TRANSP_VALUE2));

            if(m_IsOutline)
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
            else
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);

            if(type.equals("RectangleShape")){
                if(m_IsFrame){
                    if(m_IsRoundedFrame){
                        xProp.setPropertyValue("CornerRadius", new Integer(RelationDiagram.CORNER_RADIUS));
                    }else{
                        xProp.setPropertyValue("CornerRadius", new Integer(0));
                    }
                }else{
                    xProp.setPropertyValue("FillTransparence", new Integer(100));
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
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
