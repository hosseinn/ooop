package org.openoffice.extensions.diagrams.diagram.venndiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;
import org.openoffice.extensions.diagrams.Controller;
import org.openoffice.extensions.diagrams.Gui;
import org.openoffice.extensions.diagrams.diagram.Diagram;





public class VennDiagram extends Diagram {

    private int             m_GroupSize     = 0;

    protected final short   DEFAULT         = 0;
    protected final short   MONOGRAPHIC     = 1;
    protected final short   ROUNDED         = 2;
    protected final short   WITHOUTFRAME    = 3;
    protected final short   USER_DEFINE     = 4;


    public VennDiagram(Controller controller, Gui gui, XFrame xFrame){
        super(controller, gui, xFrame);
    }

    @Override
    public String getDiagramType(){
        return "VennDiagram";
    }

    @Override
    public void createDiagram(){
        super.createDiagram();
        createDiagram(3);
    }

    public void createDiagram(int n){
        try {

            m_GroupSize = m_GroupSizeWidth <= m_GroupSizeHeight ? m_GroupSizeWidth : m_GroupSizeHeight;
            m_xGroupShape.setSize( new Size( m_GroupSize, m_GroupSize ) );

            int halfDiff = 0;
            if(m_GroupSize < m_GroupSizeWidth)
                halfDiff = (m_GroupSizeWidth - m_GroupSize) / 2;
            m_xGroupShape.setPosition( new Point( m_PageProps.BorderLeft + halfDiff, m_PageProps.BorderTop ) );

            int ellipseSize = m_GroupSize / 3;
            int xR;
            int yR;
            xR = yR = ellipseSize / 2;
            Point middlePoint = new Point(m_GroupSize / 2 + m_PageProps.BorderLeft + halfDiff, m_GroupSize / 2 + m_PageProps.BorderTop);
            //ControlEllipse
            Point coord = new Point(middlePoint.X - xR, middlePoint.Y - yR);
            Size size = new Size(ellipseSize, ellipseSize);
            XShape xShape = this.createShape("EllipseShape", 0, coord.X, coord.Y, size.Width, size.Height);
            m_xShapes.add(xShape);
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillColor", new Integer(0xFFFFFF));
            xProp.setPropertyValue("FillTransparence", new Integer(1000));
            xProp.setPropertyValue("LineColor", new Integer(0xFFFFFF));
            xProp.setPropertyValue("LineTransparence", new Integer(1000));
            xProp.setPropertyValue("MoveProtect", new Boolean(true));
            getController().setSelectedShape((Object)m_xShapes);

            getController().removeSelectionListener();
            if(m_xDrawPage != null && m_xShapes != null){
                for( int i=0; i< n; i++ )
                    addShape();
                refreshDiagram();
            }
            getController().addSelectionListener();
            getController().setSelectedShape((Object)m_xShapes);

        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void addShape(){

        XShape xControlCircleShape = null;
        int iTopShapeID = -1;
        XShape xCurrShape = null;
        String currShapeName = "";

        try {

            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("EllipseShape")) {
                    int shapeID = getController().getNumberOfShape(currShapeName);
                    if ( shapeID > iTopShapeID)
                        iTopShapeID = shapeID;
                    if (currShapeName.endsWith("EllipseShape0"))
                        xControlCircleShape = xCurrShape;
                }
            }
            if( iTopShapeID == -1 ){
                //m_xDrawPage.remove(m_xGroupShape);
                createDiagram(1);
            }else{
                Size size = xControlCircleShape.getSize();
                iTopShapeID++;
                XShape xShape = createShape("EllipseShape",iTopShapeID, size.Width, size.Height);
                m_xShapes.add(xShape);
                setShapeProperties( xShape, "EllipseShape" );
                int color = -1;
                if(getGui() != null && getGui().getControlDialogWindow() != null)
                    color = getGui().getImageColorOfControlDialog();
                if(color == -1)
                    color = aCOLORS[ (iTopShapeID - 1) % 8 ];
                setColorOfShape(xShape, color);
                getController().setSelectedShape((Object)xShape);
                xShape = createShape( "RectangleShape", iTopShapeID, size.Width/2, size.Height/4 );
                m_xShapes.add(xShape);
                setShapeProperties( xShape, "RectangleShape" );
                setColorOfShape(xShape, color);
                setTextOfRectangle(xShape);
                if(getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(aCOLORS[ (iTopShapeID) % 8] );
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void removeShape() {

        XShape xCurrShape = null;
        XShape xControlCircle = null;
        String currShapeName = "";
        int iTopShapeID = 0;
        XShape xSelectedShape = getController().getSelectedShape();
        String selectedShapeName = getShapeName(xSelectedShape);
        String selectedShapeType = "";

        try {
           
            if(selectedShapeName.contains("EllipseShape"))
                selectedShapeType = "EllipseShape";
            if(selectedShapeName.contains("RectangleShape"))
                selectedShapeType = "RectangleShape";
            //user can't remove the controlShapes with the controlPanel
            if(!selectedShapeType.equals("") && !selectedShapeName.endsWith("GroupShape") && !selectedShapeName.endsWith("EllipseShape0")){
                //String sNum = getController().getDaigramShapeNumberStr(selectedShapeName);
                int iSelectedShapeNum = getController().getNumberOfShape(selectedShapeName);
                int iNumBefore = 0;
                int iCurrShapeNumber = 0;
                XShape xRectangleShape = null;
                XShape xPreviousShape = null;  //Circle or Rectangle
                XShape xLastShape = null;      //Circle or Rectangle
                m_xShapes.remove(xSelectedShape);
                //seek shapes
                for(int i=0; i < m_xShapes.getCount(); i++){
                    xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                    currShapeName = getShapeName(xCurrShape);
                    iCurrShapeNumber = getController().getNumberOfShape(currShapeName);
                    //seek the controlCircle
                    if (currShapeName.endsWith("EllipseShape0")) 
                        xControlCircle = xCurrShape;
                    if (currShapeName.contains(selectedShapeType)) {
                        //seek the last Shape and iTopShapeID
                        if (iCurrShapeNumber > iTopShapeID) {
                            iTopShapeID = iCurrShapeNumber;
                            xLastShape = xCurrShape;
                        }
                        if (iCurrShapeNumber < iSelectedShapeNum) {
                            if (iCurrShapeNumber > iNumBefore) {
                                iNumBefore = iCurrShapeNumber;
                                xPreviousShape = xCurrShape;
                            }
                        }
                    }
                    if ( selectedShapeType.equals("EllipseShape") )
                        if ( iCurrShapeNumber == iSelectedShapeNum && currShapeName.contains("RectangleShape") )
                            xRectangleShape = xCurrShape;
                }
                //remove the rectangle too
                if(xRectangleShape != null)
                    m_xShapes.remove(xRectangleShape);
                //and select the previous shape
                if(xPreviousShape != null){
                    getController().setSelectedShape((Object)xPreviousShape);
                }else{
                    if(xLastShape != null){
                        getController().setSelectedShape((Object)xLastShape);
                    }else{
                        getController().setSelectedShape((Object)xControlCircle);
                    }
                }
            }
            if(getGui() != null && getGui().getControlDialogWindow() != null)
                    getGui().setImageColorOfControlDialog(aCOLORS[ (iTopShapeID) % 8] );
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void refreshDiagram() {
        try {
            //init
            XShape xShape = null;
            XShape xControlCircle = null;
            int numCircle = 0;
            int iTopShapeID = 0;
            ArrayList<XShape> circleList = new ArrayList();
            ArrayList<XShape> textList = new ArrayList();
            String shapeName = "";
            Point point;
            Size size;
            int xR;
            int yR;
            Point middlePoint;
            //fill the arrayLists and adjust the controrCircle, numCircle, iTopID
            for(int i=0;i<m_xShapes.getCount();i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                shapeName = this.getShapeName(xShape);
                if (shapeName.contains("EllipseShape")) {
                    if (shapeName.endsWith("EllipseShape0")) {
                        xControlCircle = xShape;
                    } else {
                        numCircle++;
                        int shapeID = getController().getNumberOfShape(shapeName);
                        if (shapeID > iTopShapeID) {
                            iTopShapeID = shapeID;
                        }
                        circleList.add(xShape);
                    }
                }
                if (shapeName.contains("RectangleShape")) {
                    textList.add(xShape);
                }
            }

            // remove the singli rectangles
            int currRectShapeID = -1;
            int currCircleShapeID = -1;
            boolean isValid = false;

            for(XShape currRectXShape : textList){
                isValid = false;
                currRectShapeID = getController().getNumberOfShape(getShapeName(currRectXShape));
                if(currRectShapeID > 0){
                    for(XShape currCircleXShape : circleList){
                        currCircleShapeID = getController().getNumberOfShape(getShapeName(currCircleXShape));
                        if(currRectShapeID == currCircleShapeID)
                            isValid = true;
                    }
                    if(!isValid){
                        m_xShapes.remove(currRectXShape);
                    }
                }
            }
            //use the controlCircle to adjust the scale
            XShape xCircleShape = null;
            XShape xTextBoxShape = null;
            point = xControlCircle.getPosition();
            size = xControlCircle.getSize();
            xR = size.Width/2;
            yR = size.Height/2;
            middlePoint = new Point( point.X + xR, point.Y + yR);
            //set positions of shapes
            if(numCircle == circleList.size()){
                if(numCircle == 1){
                    xCircleShape = circleList.get(0);
                    xCircleShape.setPosition(new Point(point.X, point.Y));
                    xTextBoxShape = textList.get(0);
                    //may be rectangle (pair of ellipse) had been removed by user
                    if(xTextBoxShape != null)
                        xTextBoxShape.setPosition(new Point(middlePoint.X - xR/2, point.Y - yR));
                }else{
                    double angle = 0.0;
                    if(numCircle == 3)
                        angle = -Math.PI/2;
                    int radius = size.Width/3 <= size.Height/3 ? size.Width/3 : size.Height/3;
                    int radius2 = size.Width <= size.Height ? size.Width : size.Height;
                    radius2 *= 1.2;
                    //make numCircle(number of circle) angles around the middlePoint and instantate the shapes
                    for(int i=0; i < numCircle; i++, angle += 2.0 * Math.PI / numCircle){
                        int xCoord = (int)(point.X + radius * Math.cos(angle));
                        int yCoord = (int)(point.Y + radius * Math.sin(angle));
                        xCircleShape = circleList.get(i);
                        xCircleShape.setPosition(new Point(xCoord, yCoord));
                        String sNum = getController().getNumberStrOfShape(getShapeName(xCircleShape));
                        for(XShape item : textList){
                            String sCurrItemNum = getController().getNumberStrOfShape(getShapeName(item));
                            if(sCurrItemNum.equals(sNum)){
                                xTextBoxShape = item;
                                xCoord = (int)(middlePoint.X + radius2 * Math.cos(angle));
                                yCoord = (int)(middlePoint.Y + radius2 * Math.sin(angle));
                                xTextBoxShape.setPosition(new Point(xCoord - xR/2, yCoord - yR/4 ));
                            }
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
    public void initDiagram(){
        try {
            super.initDiagram();
            XShape xCurrShape;
            String currShapeName = "";
            int currShapeNum;
            int iTopShapeID = 0;
        
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains("EllipseShape")) {
                    if (!currShapeName.endsWith("EllipseShape0")) {
                        currShapeNum = getController().getNumberOfShape(currShapeName);
                        if (currShapeNum > iTopShapeID) {
                            iTopShapeID = currShapeNum;
                        }
                    }
                }
            }  
            if( getGui() != null)
                getGui().setImageColorOfControlDialog(aCOLORS[iTopShapeID % 8]);
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    @Override
    public void refreshShapeProperties() {
    
        if( m_Style == USER_DEFINE ){
            // need to memorize members, if user exit into propsDialog
            boolean isSelectAllShape = m_IsSelectAllShape;
            boolean isColor = m_IsColor;
            short   sTransparency = m_sTransparency;
            boolean isMonographic = m_IsMonographic;
            boolean isFrame = m_IsFrame;
            boolean isRoundedFrame = m_IsRoundedFrame;

            //default values
            m_IsSelectAllShape = true;
            m_IsColor = false;
            m_sTransparency = (short)1;
            m_IsMonographic = false;
            m_IsFrame = true;
            m_IsRoundedFrame = false;

            m_IsAction = false;
                
            getGui().showDiagramPropsDialog();

            if(m_IsAction){
                if(!m_IsSelectAllShape){
                    m_IsMonographic = isMonographic;
                    m_IsFrame = isFrame;
                    m_IsRoundedFrame = isRoundedFrame;

                    setSelectedShapeProperties();
                }else{
                    m_IsColor = false;
                    if(!m_IsFrame)
                        m_IsRoundedFrame = isRoundedFrame;

                    setAllShapeProperties();
                }
            }else{
                m_IsSelectAllShape = isSelectAllShape;
                m_IsColor = isColor;
                m_sTransparency = sTransparency;
                m_IsMonographic = isMonographic;
                m_IsFrame = isFrame;
                m_IsRoundedFrame = isRoundedFrame;
            }
            m_IsAction = false;
        }else{

            setAllShapeProperties();
        }
    }

    public void setSelectedShapeProperties(){
        try {


            XShape xShape = getController().getSelectedShape();
            String shapeName = getShapeName(xShape);
            int shapeID;
            XShape xPairOfShape = null;
            XShape xCurrShape = null;
            String currShapeName = "";
            int currShapeID;

            if(shapeName.contains("GroupShape")){
                setAllShapeProperties();
                if(m_IsColor)
                    getGui().setImageColorOfControlDialog(m_iColor);
            }else{
                if(!shapeName.endsWith("EllipseShape0") && !shapeName.equals("")){
                    shapeID = getController().getNumberOfShape(shapeName);
                    System.out.println(shapeID);
                    for(int i=0; i < m_xShapes.getCount(); i++){
                        xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                        currShapeName = getShapeName(xCurrShape);
                        currShapeID = getController().getNumberOfShape(currShapeName);
                        if( currShapeID == shapeID && ( (shapeName.contains("EllipseShape") && currShapeName.contains("RectangleShape")) || ((shapeName.contains("RectangleShape") && currShapeName.contains("EllipseShape")) ) ) )
                            xPairOfShape = xCurrShape;
                    }
                    if(shapeName.contains("EllipseShape")){
                        setShapeProperties(xShape, "EllipseShape");
                        if(xPairOfShape != null)
                            setShapeProperties(xPairOfShape, "RectangleShape");
                    }
                    if(shapeName.contains("RectangleShape")){
                        setShapeProperties(xShape, "RectangleShape");
                        if(xPairOfShape != null)
                            setShapeProperties(xPairOfShape, "EllipseShape");
                    }
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setAllShapeProperties(){
        try {

            XShape xShape = null;
            String name = "";
            for(int i=0; i < m_xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                name = getShapeName(xShape);
                if (name.contains("EllipseShape") && !name.endsWith("EllipseShape0"))
                    setShapeProperties(xShape, "EllipseShape");
                if (name.contains("RectangleShape"))
                    setShapeProperties(xShape, "RectangleShape");
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
            if( m_Style == DEFAULT ){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if (type.equals("RectangleShape")) {
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                }
            }else if( m_Style == MONOGRAPHIC){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineColor", new Integer(0x000000));
                xProp.setPropertyValue("LineTransparence", new Integer(10));
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                }
            }else if( m_Style == ROUNDED){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(300));
                }
            } else if( m_Style == WITHOUTFRAME){
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if(type.equals("EllipseShape")){
                    xProp.setPropertyValue("FillTransparence", new Integer(20));
                }
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                    xProp.setPropertyValue("FillTransparence", new Integer(100));
                }
            }
            else if( m_Style == USER_DEFINE){

                if(!m_IsSelectAllShape)
                    if(m_IsColor)
                        xProp.setPropertyValue("FillColor", new Integer(m_iColor));

                if(m_sTransparency == (short)0)
                    xProp.setPropertyValue("FillTransparence", new Integer(0));
                if(m_sTransparency == (short)1)
                    xProp.setPropertyValue("FillTransparence", new Integer(20));
                if(m_sTransparency == (short)2)
                    xProp.setPropertyValue("FillTransparence", new Integer(40));

                if(m_IsMonographic){
                    xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                    xProp.setPropertyValue("LineColor", new Integer(0x000000));
                    xProp.setPropertyValue("LineTransparence", new Integer(10));
                }else{
                    xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                }

                if(type.equals("RectangleShape")){
                    if(m_IsFrame){
                        if(m_IsRoundedFrame){
                            xProp.setPropertyValue("CornerRadius", new Integer(300));
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

    public void setTextOfRectangle(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
            XText xText = (XText) UnoRuntime.queryInterface(XText.class, xShape);
            xText.setString(getGui().getDialogPropertyValue( "ControlDialog1", "ControlDialog1.Text.Label" ) );
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
  
}
