/*****************************************************************
 * 
 * file: Controller.java
 * 
 * The Venn Diagram Extension has been licensed under term of
 * GNU Lesser General Public License (LGPL). 
 * You can read and understand LGPL license here:
 * http://www.gnu.org/licenses/lgpl.html
 * 
 * This extension is created by: Tibor Hornyák and
 * University of Szeged, Hungary - http://www.u-szeged.hu/english/
 * OxygenOffice Professional Team - http://ooop.sf.net/
 * 
 * license_en-US document V1.0.0
 * 
 ****************************************************************/

package org.openoffice.extensions.venndiagram;

import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNameAccess;
import com.sun.star.container.XNamed;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XController;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.Locale;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XLocalizable;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.resource.StringResourceWithLocation;
import com.sun.star.resource.XStringResourceWithLocation;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import java.util.ArrayList;


public class Controller {
    
    private              VennDiagram          m_VennDiagram = null;
    private              XComponentContext    m_xContext;
    private              XFrame               m_xFrame;
    private              XModel               m_xModel;
    private              XMultiServiceFactory m_xMSF;
    private              XController          m_xController;
    private              XDrawPage            m_xDrawPage;
    private              XShapes              m_xShapes;
    private int                               m_id;
    private              XShape               m_xGroupShape;
    private              Gui                  m_Gui = null;
    protected static final int[]              colors = {0x00FF00,0x0000FF,0xFF0000,0xFFFF33, 
                                                         0x0099FF,0x990099,0x99FF99,0xFF6666 };
    
    private static final short                DEFAULT = 0;
    private static final short                MONOGRAPHIC = 1;
    private static final short                ROUNDED = 2;
    private static final short                WITHOUTFRAME = 3;
    private short                             style;
  
    public Controller(VennDiagram vennDiagram, XComponentContext context, XFrame xFrame){
        m_VennDiagram = vennDiagram;
        m_xContext = context;
        m_xFrame = xFrame;
        m_xController = m_xFrame.getController();
        m_xModel = m_xController.getModel();
        m_xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, m_xModel);
        m_id = 0;
        style = DEFAULT;
    }
    
    public XDrawPage getDrawPage(){
        return m_xDrawPage;
    }
    
    public XShapes getShapes(){
        return m_xShapes;
    }

    public void setVisibleGui(boolean bool) throws Exception{
        if (m_Gui == null && VennDiagram._bool) {
            VennDiagram._bool = false;
            m_Gui = new Gui(this, m_xContext, m_xModel);
        }
        m_Gui.setVisible(bool);
    }
  
    public Locale getLocation() {
        Locale locale = null;
        try {
            XMultiComponentFactory xServiceManager = m_xContext.getServiceManager();
            Object oConfigurationProvider = xServiceManager.createInstanceWithContext("com.sun.star.configuration.ConfigurationProvider", m_xContext);
            XLocalizable xLocalizable = (XLocalizable) UnoRuntime.queryInterface(XLocalizable.class, oConfigurationProvider);
            locale = xLocalizable.getLocale();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return locale;
    }
    
    public void createDiagram(int numer) {
        try {
            if(m_Gui != null)
                m_Gui.setBackgroundColor(colors[0]);
            m_xDrawPage = getCurrentPage();
            m_id = (int)(Math.random()*10000);
            //page properties
            XPropertySet xPageProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xDrawPage);
            int width = AnyConverter.toInt(xPageProperties.getPropertyValue("Width"));
            int height = AnyConverter.toInt(xPageProperties.getPropertyValue("Height"));
            int borderLeft = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderLeft"));
            int borderRight = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderRight"));
            int borderTop = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderTop"));
            int borderBottom = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderBottom"));
            int groupWidth = width - borderLeft - borderRight;
            int groupHeight = height - borderTop - borderBottom;
            int groupSize = groupWidth <= groupHeight ? groupWidth : groupHeight;
            int ellipseSize = groupSize / 3;
            int xR, yR;
            xR = yR = ellipseSize/2;
            Point middlePoint = new Point(groupSize/2+borderLeft,groupSize/2+borderTop);
            
            m_xGroupShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing.GroupShape"));
            m_xGroupShape.setPosition(new Point(borderLeft, borderTop));
            m_xGroupShape.setSize(new Size(groupSize, groupSize));
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,m_xGroupShape);
            xNamed.setName("VennDiagram" + m_id + "-GroupShape");
            m_xDrawPage.add(m_xGroupShape);
            m_xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, m_xGroupShape );
            //base circle
            XShape xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing.EllipseShape"));
            xShape.setPosition(new Point(middlePoint.X - xR, middlePoint.Y - yR));
            xShape.setSize(new Size(ellipseSize, ellipseSize));
            xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
            xNamed.setName("VennDiagram" + m_id + "-EllipseShape0");
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillColor", new Integer(0xFFFFFF));
            xProp.setPropertyValue("FillTransparence", new Integer(1000));
            xProp.setPropertyValue("LineColor", new Integer(0xFFFFFF));
            xProp.setPropertyValue("LineTransparence", new Integer(1000));
            xProp.setPropertyValue("MoveProtect", new Boolean(true));
            m_xShapes.add(xShape);
            if(m_xDrawPage != null && m_xShapes != null){
                for(int i=0;i<numer;i++){
                    addShape();
                }
                refreshDiagram();
            }
            m_VennDiagram.setSelectedShape((Object)m_xShapes);
        } catch (IndexOutOfBoundsException ex) {
            ex.printStackTrace();
        } catch (UnknownPropertyException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (PropertyVetoException ex) {
            ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }
    
    public void initDiagram() throws WrappedTargetException, IndexOutOfBoundsException, UnknownPropertyException, IllegalArgumentException{
        XShape xShape = null;
        String name = "";
        int iTopID = 0; 
        //initial members
        m_xDrawPage = getCurrentPage();
        m_id = parseInt(getCurrentDiagramIdName());
        
        for(int i=0; i < m_xDrawPage.getCount(); i++){
            xShape = (XShape)UnoRuntime.queryInterface(XShape.class, m_xDrawPage.getByIndex(i));
            name = getShapeName(xShape);
            if(name.contains(getCurrentDiagramIdName()) && name.startsWith("VennDiagram") && name.endsWith("GroupShape")){
                m_xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, xShape );
            }
        }
        
        for(int i=0;i<m_xShapes.getCount();i++){
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
            name = this.getShapeName(xShape);
            if (name.contains("EllipseShape")) {
                if (!name.endsWith("EllipseShape0")) {
                    if (Integer.parseInt(getNumer(name)) > iTopID) {
                        iTopID = parseInt(getNumer(name));
                    }
                }
            }
        }  
        if(m_Gui != null){
            m_Gui.setBackgroundColor(colors[iTopID % 8]);      
        }
    }
    
    public void setBackgroundColor(int color){
        m_Gui.setBackgroundColor(color);
    }
    
    public void refreshDiagram() throws IndexOutOfBoundsException, WrappedTargetException {
        //init
        XShape xShape = null;
        XShape xControlCircle = null;
        int numCircle = 0;
        int iTopID = 0; 
        ArrayList<XShape> circleList = new ArrayList();
        ArrayList<XShape> textList = new ArrayList(); 
        String name = "";
        //initial later
        Point point;
        Size size;
        int xR;
        int yR;
        Point middlePoint;  
        //fill the arrayLists and adjust the controrCircle, numCircle, iTopID
        for(int i=0;i<m_xShapes.getCount();i++){
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
            name = this.getShapeName(xShape);
            if (name.contains("EllipseShape")) {
                if (name.endsWith("EllipseShape0")) {
                    xControlCircle = xShape;
                } else {
                    numCircle++;
                    if (Integer.parseInt(getNumer(name)) > iTopID) {
                        iTopID = parseInt(getNumer(name));
                    }
                    circleList.add(xShape);
                }
            }
            if (name.contains("RectangleShape")) {
                textList.add(xShape);
            }
        }
        //use the controlCircle to adjust the scale
        point = xControlCircle.getPosition();
        size = xControlCircle.getSize();
        xR = size.Width/2;
        yR = size.Height/2;
        middlePoint = new Point( point.X + xR, point.Y + yR); 
        //set positions of shapes
        if(numCircle == circleList.size()){
            if(numCircle == 1){
                xShape = circleList.get(0);
                xShape.setPosition(new Point(point.X, point.Y));
                xShape = textList.get(0);
                //may be rectangle (pair of ellipse) had been removed by user
                if(xShape != null)     
                    xShape.setPosition(new Point(middlePoint.X - xR/2, point.Y - yR));         
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
                    xShape = circleList.get(i);
                    xShape.setPosition(new Point(xCoord, yCoord));
                    name = getShapeName(xShape); 
                    String sNum = getNumer(name);  
                    for(XShape item : textList){
                        name = getShapeName(item);
                        if(getNumer(name).equals(sNum)){
                            xShape = item;
                            xCoord = (int)(middlePoint.X + radius2 * Math.cos(angle));
                            yCoord = (int)(middlePoint.Y + radius2 * Math.sin(angle));
                            xShape.setPosition(new Point(xCoord - xR/2, yCoord - yR/4 ));         
                        }
                    }    
                } 
            }
        }  
    }
    
    public void addShape() throws IndexOutOfBoundsException, WrappedTargetException, Exception{
        XShape xControlCircle = null;
        int iTopID = -1;
        XShape xShape = null;
        String name = "";
        
        for(int i=0;i<m_xShapes.getCount();i++){
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
            name = this.getShapeName(xShape);
            if(name.contains("EllipseShape")){
                if(Integer.parseInt(getNumer(name))>iTopID)
                        iTopID = parseInt(getNumer(name));
                if(name.endsWith("EllipseShape0"))
                    xControlCircle = xShape;
            }
        }
        if(iTopID == -1){
            m_xDrawPage.remove(m_xGroupShape);
            createDiagram(1);
        }else{
            Size size = xControlCircle.getSize();
            iTopID++;
            int color = -1;
            if(m_Gui != null) color = m_Gui.getBackgroundColor();
            if(color == -1) color = colors[(iTopID - 1) % 8];
            xShape = createShape(iTopID, size.Width, size.Height, "EllipseShape", color);
            m_xShapes.add(xShape);
            m_VennDiagram.setSelectedShape((Object)xShape);
            xShape = createShape(iTopID, size.Width/2, size.Height/4, "RectangleShape", color);
            m_xShapes.add(xShape);
            XPropertySet xPropText = (XPropertySet)UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue( "TextFitToSize",TextFitToSizeType.PROPORTIONAL);
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xShape);
            xText.setString(getDialogPropertyValue("GuiPanel.Text.Label"));
            if(m_Gui !=null) 
                    m_Gui.setBackgroundColor(colors[(iTopID) % 8]);
        }
        if(iTopID == 3){
            int i =0;
            while(i<2){
                addShape();
                refreshDiagram();
                removeShape();
                refreshDiagram();
                i++;
            }
        }    
    }
    
    public String getDialogPropertyValue(String name){
        String result = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            String location =  xPIP.getPackageLocation("org.openoffice.extensions.venndiagram.VennDiagram");
            String m_resRootUrl = location + "/dialogs/";
            XStringResourceWithLocation xResources = StringResourceWithLocation.create(m_xContext, m_resRootUrl, true, getLocation(), "GuiPanel", "", null);
            String[] ids = xResources.getResourceIDs();
            for (int i = 0; i < ids.length; i++) {
                if(ids[i].contains(name))
                    result = xResources.resolveString(ids[i]);
            }
        } catch (NoSuchElementException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        }
        return result;
    }
    
    public void removeShape() throws WrappedTargetException, IndexOutOfBoundsException, IllegalArgumentException {
        XShape xShape = null;
        XShape xControlCircle = null;
        String name = "";
        int iTopID = 0;
        XShape xSelectedShape = m_VennDiagram.getSelectedShape();
        String selectedShapeName = getShapeName(xSelectedShape);
        String selectedShapeType = "";
        if(selectedShapeName.contains("EllipseShape"))
            selectedShapeType = "EllipseShape";
        if(selectedShapeName.contains("RectangleShape"))
            selectedShapeType = "RectangleShape";   
        //user can't remove on the controlPanel the control shapes
        if(!selectedShapeType.equals("") && !selectedShapeName.endsWith("GroupShape") && !selectedShapeName.endsWith("EllipseShape0")){    
            
            String sNum = getNumer(selectedShapeName);
            int iNum = parseInt(sNum);
            int iNumBefore = 0;
            int iShapeNumber = 0;
            XShape xRectangleShape = null;
            m_xShapes.remove(xSelectedShape);
            XShape xPreviousShape = null;
            XShape xLastShape = null;     
            //seek shapes
            for(int i=0; i < m_xShapes.getCount(); i++){
                xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                name = getShapeName(xShape);
                //seek the controlCircle
                if(name.endsWith("EllipseShape0"))
                       xControlCircle = xShape;
                if(name.contains(selectedShapeType)){
                    //seek the last shape
                    iShapeNumber = parseInt(getNumer(name));
                    if(iShapeNumber>iTopID){
                        iTopID = iShapeNumber;
                        xLastShape = xShape;
                    }
                    //seek the previous shape
                    if(iShapeNumber< iNum){
                        if(iShapeNumber>iNumBefore){
                            iNumBefore = iShapeNumber;
                            xPreviousShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i)); 
                        }
                    } 
                }
                //if shape is circle, remove the corresponding rectangle (pair of circle)
                if(selectedShapeType.equals("EllipseShape")){
                    if(getNumer(name).equals(sNum) && name.contains("RectangleShape")){
                        xRectangleShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i)); 
                    }  
                }
            }
            //remove the rectangle too
            if(xRectangleShape != null)
                m_xShapes.remove(xRectangleShape);
            //and select the previous shape
            if(xPreviousShape != null){
                m_VennDiagram.setSelectedShape((Object)xPreviousShape);
            }else{
                if(xLastShape != null){
                    m_VennDiagram.setSelectedShape((Object)xLastShape);
                }else{
                    m_VennDiagram.setSelectedShape((Object)xControlCircle); 
                }
            }
        } 
    }
    
    public XShape createShape(int x, int y, int width, int height, String type){
        XShape xShape = null;
        try {
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing." + type ));
            xShape.setPosition(new Point(x, y));
            xShape.setSize(new Size(width, height));
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
            xNamed.setName("VennDiagram-" +m_id+ type + 0);
        } catch (PropertyVetoException ex) {
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return xShape;
    }
    
    public XShape createShape(int num, int width, int height, String type, int color) {
        XShape xShape = null;
        try {
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing." + type ));
            xShape.setSize(new Size(width, height));
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
            xNamed.setName("VennDiagram" + m_id + "-" + type + num);
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillColor", new Integer(color));
            setShapeProperties(xShape, type);
        } catch (UnknownPropertyException ex) {
            ex.printStackTrace();
        } catch (IllegalArgumentException ex) {
            ex.printStackTrace();
        } catch (WrappedTargetException ex) {
            ex.printStackTrace();
        } catch (PropertyVetoException ex) {
            ex.printStackTrace();
        } catch (Exception ex) {
            ex.printStackTrace();
        } 
        return xShape;    
    }
    
    public void setShapeProperties(XShape xShape, String type) {
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if( style == DEFAULT ){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if (type.equals("RectangleShape")) {
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                }
            }else if( style == MONOGRAPHIC){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.SOLID);
                xProp.setPropertyValue("LineColor", new Integer(0x000000));
                xProp.setPropertyValue("LineTransparence", new Integer(10));
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                }
            }else if( style == ROUNDED){
                xProp.setPropertyValue("FillTransparence", new Integer(20));
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(500));
                }
            } else if( style == WITHOUTFRAME){
                xProp.setPropertyValue("LineStyle", LineStyle.NONE);
                if(type.equals("EllipseShape")){
                    xProp.setPropertyValue("FillTransparence", new Integer(20));
                }
                if(type.equals("RectangleShape")){
                    xProp.setPropertyValue("CornerRadius", new Integer(0));
                    xProp.setPropertyValue("FillTransparence", new Integer(100));
                }
            }    
         } catch (UnknownPropertyException ex) {
             ex.printStackTrace();
         } catch (PropertyVetoException ex) {
             ex.printStackTrace();
         } catch (IllegalArgumentException ex) {
             ex.printStackTrace();
         } catch (WrappedTargetException ex) {
             ex.printStackTrace();
         }   
    }
    
    public void refreshShapeProperties() throws IndexOutOfBoundsException, WrappedTargetException, Exception {
        XShape xShape = null;
        String name = "";
        for(int i=0; i < m_xShapes.getCount(); i++){
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
            name = getShapeName(xShape);
            if(name.contains("EllipseShape")&& !name.endsWith("EllipseShape0"))
                setShapeProperties(xShape,"EllipseShape");
            if(name.contains("RectangleShape"))
                setShapeProperties(xShape,"RectangleShape");
        }
    }
    
    public String getShapeName(XShape xShape){
        if(xShape != null){
           XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
           return xNamed.getName();
        }
        return null;
    } 
     
    //get the specific shape id
    public String getNumer(String name){
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  charName[i] != '-')
            i++;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length())
           s +=  charName[i++];
        return s;
    }
    
    public XDrawPage getCurrentPage(){
        XDrawView xDrawView = (XDrawView)UnoRuntime.queryInterface(XDrawView.class, m_xController); 
        return xDrawView.getCurrentPage();  
    } 
    
    public String getCurrentDiagramIdName() throws WrappedTargetException, IndexOutOfBoundsException {
        String name = getShapeName(m_VennDiagram.getSelectedShape());
        String s = "";
        char[] charName = name.toCharArray();
        int i = 0;
        while(i<name.length() &&  ( charName[i] < 48 || charName[i] > 57))
            i++;
        while(i<name.length() &&  charName[i] != '-')
           s +=  charName[i++];
        return s;
    }
    
    public int parseInt(String s) {
        int n = 0;
        try{
            n = Integer.parseInt(s);
        }catch(NumberFormatException ex){
             ex.printStackTrace();
        }
        return n;
    }
    
    public void changeMode(short selected) {
        style = selected;
    }
    
} 
