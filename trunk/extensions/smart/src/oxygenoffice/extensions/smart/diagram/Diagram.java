package oxygenoffice.extensions.smart.diagram;
import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.XNamed;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.TextFitToSizeType;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.text.XText;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.Exception;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChart;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagram;
import oxygenoffice.extensions.smart.diagram.relationdiagrams.RelationDiagram;
import oxygenoffice.extensions.smart.gui.Gui;



public abstract class Diagram extends DiagramProperties{


    protected   Controller              m_Controller        = null;
    protected   Gui                     m_Gui               = null;
    protected   XFrame                  m_xFrame            = null;
    protected   XModel                  m_xModel            = null;
    protected   XMultiServiceFactory    m_xMSF              = null;

    protected   XDrawPage               m_xDrawPage         = null;
    protected   int                     m_DiagramID         = -1;

    //struct witch stores Width-Height-BorderLeft-BorderRight-BorderTop-BorderBottom properties of page
    public      PageProps               m_PageProps         = null;
    
    // m_GroupSize is the side of the bigest possible rectangle in the draw page
    protected   int                     m_DrawAreaWidth     = 0;
    protected   int                     m_DrawAreaHeight    = 0;
    protected   XShape                  m_xGroupShape       = null;
    protected   XNamed                  m_xNamed            = null;
    protected   XShapes                 m_xShapes           = null;


    public Diagram(){ }
    
    public Diagram(Controller controller, Gui gui, XFrame xFrame) {
        m_Controller    = controller;
        m_Gui           = gui;
        m_xFrame        = xFrame;
        m_xModel = m_xFrame.getController().getModel();
        m_xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, m_xModel);
        setGroupSize();
    }
    
    public void setGroupShapeSizeAndPos(int width, int height, int xPos, int yPos){
        if(m_xGroupShape != null){
            try {
                m_xGroupShape.setSize(new Size(width, height));
                m_xGroupShape.setPosition(new Point(xPos, yPos));
            }  catch (Exception ex) {
                System.err.println(ex.getLocalizedMessage());
            }
        }
    }
    
    public Object getGroupShape(){
        if(m_xGroupShape != null)
            return m_xGroupShape;
        return null;
    }
    
    public Point getGroupShapePos(){
        if(m_xGroupShape != null)
            return m_xGroupShape.getPosition();
        return null;
    }
    
    public Size getGroupShapeSize(){
        if(m_xGroupShape != null)
            return m_xGroupShape.getSize();
        return null;
    }
    
    public Controller getController(){
        return m_Controller;
    }

    public Gui getGui(){
        return m_Gui;
    }

    public abstract String getDiagramTypeName();

    public int getDiagramID(){
        return m_DiagramID;
    }

    public int getDrawAreaWidth(){
        return m_DrawAreaWidth;
    }
    
    public void setDrawAreaWidth(int s){
        m_DrawAreaWidth = s;
    }

    public int getDrawAreaHeight(){
        return m_DrawAreaHeight;
    }
    
    public void setDrawAreaHeight(int s){
        m_DrawAreaHeight = s;
    }

    // determinde m_GroupSize
    public final void setGroupSize(){
        if(m_xDrawPage == null);
            m_xDrawPage = getController().getCurrentPage();
        if(m_PageProps == null)
            adjustPageProps();
        if(m_PageProps != null){
            m_DrawAreaWidth = m_PageProps.Width - m_PageProps.BorderLeft - m_PageProps.BorderRight;
            m_DrawAreaHeight = m_PageProps.Height - m_PageProps.BorderTop - m_PageProps.BorderBottom;
        }
    }

    // instantiate PageProps object
    public void adjustPageProps(){
        int width           = 0;
        int height          = 0;
        int borderLeft      = 0;
        int borderRight     = 0;
        int borderTop       = 0;
        int borderBottom    = 0;
        try {
            XPropertySet xPageProperties = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xDrawPage);
            width           = AnyConverter.toInt(xPageProperties.getPropertyValue("Width"));
            height          = AnyConverter.toInt(xPageProperties.getPropertyValue("Height"));
            borderLeft      = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderLeft"));
            borderRight     = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderRight"));
            borderTop       = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderTop"));
            borderBottom    = AnyConverter.toInt(xPageProperties.getPropertyValue("BorderBottom"));

            if(borderLeft < 1000)
                borderLeft = 1000;
            if(borderRight < 1000)
                borderRight = 1000;
            if(borderTop < 1000)
                borderTop = 1000;
            if(borderBottom < 1000)
                borderBottom = 1000;

        } catch (IllegalArgumentException ex) {
            System.err.println("IllegalArgumentException in Diagram.getPagePoperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in Diagram.getPagePoperties(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in Diagram.getPagePoperties(). Message:\n" + ex.getLocalizedMessage());
        }
        m_PageProps = new PageProps( width, height, borderLeft, borderRight, borderTop, borderBottom );
    }

    public void setFocusGroupShape(){
        XShapes xShapes = getController().getSelectedShapes();
        getController().setSelectedShape(m_xGroupShape);
        if(getController().getSelectedShapes() == null)
            getController().setSelectedShape(xShapes);
    }

    public abstract void createDiagram(DataOfDiagram datas);
    
    // set m_xDrawPage, PageProps, m_xGroupSize, m_xGroupShape and m_xShapes
    public void createDiagram(){
        try {
            m_xDrawPage = getController().getCurrentPage();
            m_DiagramID = (int) (Math.random()*10000);

            // set diagramName in the Controller object
            String diagramName = getDiagramTypeName() + m_DiagramID;
            getController().setLastDiagramName(diagramName);

            // set new PageProps object with data of page
            // width, height, borderLeft, borderRight, borderTop, borderBottom
            adjustPageProps();

            // get minimum of width and height
            setGroupSize();

            m_xGroupShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing.GroupShape"));
            m_xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, m_xGroupShape);
            m_xNamed.setName( getDiagramTypeName() + m_DiagramID + "-GroupShape" );
            m_xDrawPage.add(m_xGroupShape);
            m_xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, m_xGroupShape );
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    //initial members: m_xDrawPage, m_DiagramID, m_xShapes
    public void initDiagram(){
        try {
            XShape xCurrShape = null;
            String currShapeName = "";
            m_xDrawPage = getController().getCurrentPage();
            m_DiagramID = getController().getCurrentDiagramId();
            String sDiaramID = "" + m_DiagramID;
            for(int i=0; i < m_xDrawPage.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xDrawPage.getByIndex(i));
                currShapeName = getShapeName(xCurrShape);
                if (currShapeName.contains(sDiaramID) && currShapeName.startsWith(getDiagramTypeName()) && currShapeName.endsWith("GroupShape")) {
                    m_xShapes = (XShapes) UnoRuntime.queryInterface(XShapes.class, xCurrShape );
                    m_xGroupShape = (XShape) UnoRuntime.queryInterface(XShape.class, xCurrShape );
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public abstract void setDrawArea();
    
    public abstract void refreshDiagram();

    public abstract void addShape();

    public abstract void removeShape();

    public abstract void showPropertyDialog();

    public abstract void setShapeProperties(XShape xShape, String type);

    public void removeShapeFromGroup(XShape xShape){
        if(m_xShapes != null)
            m_xShapes.remove(xShape);
    }

    @Override
    public void setDefaultBaseColors(){
        if(getController().getGroupType() == Controller.RELATIONGROUP)
            DiagramProperties._setDefaultBaseColors("BaseColors");
        if(getController().getGroupType() == Controller.PROCESSGROUP)
            DiagramProperties._setDefaultBaseColors("RainbowColors");
    }
    
    public XShape createShape(String type, int num){
        XShape xShape = null;
        try {
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xMSF.createInstance ("com.sun.star.drawing." + type ));
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
            xNamed.setName(getDiagramTypeName() + m_DiagramID + "-" + type + num);
        }  catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xShape;
    }

    public XShape createShape(String type, int num, int width, int height){
        XShape xShape = null;
        try {
            xShape = createShape(type, num);
            xShape.setSize(new Size(width, height));
        }  catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xShape;
    }

    public XShape createShape(String type, int num, int x, int y, int width, int height){
        XShape xShape = createShape(type, num, width, height);
        xShape.setPosition(new Point(x, y));
        return xShape;
    }

    public void renameShapes(String oldDiagramName, String newDiagramName){
        XShape xShape = null;
        try {
            XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, m_xGroupShape);
            String shapeName = xNamed.getName();
            shapeName = shapeName.replace(oldDiagramName, newDiagramName);
            xNamed.setName(shapeName);
            for(int i=0; i < m_xShapes.getCount(); i++){
                xShape = (XShape)UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                if(xShape != null){
                    xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class, xShape);
                    shapeName = xNamed.getName();
                    shapeName = shapeName.replace(oldDiagramName, newDiagramName);
                    xNamed.setName(shapeName);
                }
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public boolean isInGruopShapes(XShape xShape){
        XShape xCurrShape;
        try {
            for( int i=0; i < m_xShapes.getCount(); i++ ){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                if(xCurrShape.equals(xShape))
                    return true;
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return false;
    }
  
    public void setColorOfShape(XShape xShape, int color){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillStyle", FillStyle.SOLID);
            xProp.setPropertyValue("FillColor", new Integer(color));
        }  catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setColorOfShape(XShape xShape){
        setColorOfShape(xShape, getColorProp());
    }

    public String getShapeName(XShape xShape){
        if(xShape != null){
           XNamed xNamed = (XNamed) UnoRuntime.queryInterface(XNamed.class,xShape);
           return xNamed.getName();
        }
        return null;
    }

    public abstract void setFontPropertyValues();

    public void setTextColorOfShape(XShape xShape){
        try {
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xShape);
            XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
            xTextProps.setPropertyValue("CharColor", new Integer(getTextColorProp()));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setControlShapeProps(XShape xShape){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillStyle", FillStyle.NONE);
            xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            xProp.setPropertyValue("MoveProtect", new Boolean(true));
            if(isTextFitProp())
                xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
            else
                xProp.setPropertyValue("TextFitToSize", TextFitToSizeType.NONE);
        } catch (java.lang.Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setControlShapePropsWithoutTextProps(XShape xShape){
        try {
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillStyle", FillStyle.NONE);
            xProp.setPropertyValue("LineStyle", LineStyle.NONE);
            xProp.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (java.lang.Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }
    
    public void setFontPropertiesOfShape(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            if(isTextFitProp()){
                xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.PROPORTIONAL);
            }else{
                xPropText.setPropertyValue("TextFitToSize", TextFitToSizeType.NONE);
                xPropText.setPropertyValue("CharHeight", new Float(getFontSizeProp()));
            }
            if(isTextColorChange()){
                XText xText = (XText)UnoRuntime.queryInterface(XText.class, xShape);
                XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
                xTextProps.setPropertyValue("CharColor", new Integer(getTextColorProp()));
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setMoveProtectOfShape(XShape xShape){
        try {
            XPropertySet xPropText = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xPropText.setPropertyValue("MoveProtect", new Boolean(true));
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setGradientWithAutoAngle(XShape xShape){
        if(getGradientDirectionProp() == Diagram.VERTICAL)
            setGradient(xShape);
        else
            setGradient(xShape, (short)900);
    }

    public void setGradient(XShape xShape){
        setGradient(xShape, GradientStyle.LINEAR, m_iStartColor, m_iEndColor, (short)0, (short)0, (short)0, (short)0, (short)100, (short)100);
    }
    
    public void setGradient(XShape xShape, short angle){
        setGradient(xShape, GradientStyle.LINEAR, m_iStartColor, m_iEndColor, angle, (short)0, (short)0, (short)0, (short)100, (short)100);
    }

    public void setGradient(XShape xShape, int startColor, int endColor){
        setGradient(xShape, GradientStyle.LINEAR, startColor, endColor, (short)0, (short)0, (short)0, (short)0, (short)100, (short)100);
    }

    public void setGradient(XShape xShape, GradientStyle gradientStyle, short angle, short border, short xOffset, short yOffset, short startIntensity, short endIntensity){
        setGradient(xShape, gradientStyle, m_iStartColor, m_iEndColor, angle, border, xOffset, yOffset, startIntensity, endIntensity);
    }

    public void setGradient(XShape xShape, GradientStyle gradientStyle, int startColor, int endColor, short angle, short border, short xOffset, short yOffset, short startIntensity, short endIntensity){
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("FillStyle", FillStyle.GRADIENT);
            Gradient aGradient = new Gradient();
            aGradient.Style = gradientStyle;
            aGradient.StartColor = startColor;
            aGradient.EndColor = endColor;
            aGradient.Angle = angle;
            aGradient.Border = border;
            aGradient.XOffset = xOffset;
            aGradient.YOffset = yOffset;
            aGradient.StartIntensity = startIntensity;
            aGradient.EndIntensity = endIntensity;
            aGradient.StepCount = 24;
            xProp.setPropertyValue("FillGradient", aGradient);
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setLineColorOfShape(XShape xShape){
        XPropertySet xProp = null;
        try {
            xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            xProp.setPropertyValue("LineColor", new Integer(getLineColorProp()));
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void setTextOfShape(XShape xShape, String str){
        if(xShape != null)
            ((XText) UnoRuntime.queryInterface(XText.class, xShape)).setString(str);
    }
 
    @Override
    public void setColorModeProp(short colorMode){
        super.setColorModeProp(colorMode);
        getGui().setColorModeOfImageOfControlDialog();
        if(getController().getGroupType() == Controller.ORGANIGROUP)
            setColorModeOfControlShape(((OrganizationChart)this).getDiagramTree().getControlShape(), colorMode);
        if(getController().getGroupType() == Controller.RELATIONGROUP)
            setColorModeOfControlShape(((RelationDiagram)this).getControlShape(), colorMode);
        if(getController().getGroupType() == Controller.PROCESSGROUP)
            setColorModeOfControlShape(((ProcessDiagram)this).getControlShape(), colorMode);
    }


    @Override
    public void setStyleProp(short selected) {
        super.setStyleProp(selected);
        if(getController().getGroupType() == Controller.ORGANIGROUP)
            setStyleOfControlShape(((OrganizationChart)this).getDiagramTree().getControlShape(), selected);
        if(getController().getGroupType() == Controller.RELATIONGROUP)
            setStyleOfControlShape(((RelationDiagram)this).getControlShape(), selected);
        if(getController().getGroupType() == Controller.PROCESSGROUP)
            setStyleOfControlShape(((ProcessDiagram)this).getControlShape(), selected);
    }


    public boolean isColorSchemeStyle(short style){ return false; }
    public short getColorModeOfSchemeStyle(short style){ return -1; }
    public boolean isColorThemeGradientStyle(short style){ return false; }
    public short getColorModeOfThemeGradientStyle(short style){ return -1; }
    public boolean isColorThemeStyle(short style){ return false; }
    public short getColorModeOfThemeStyle(short style){ return -1; }

    
    abstract public void initColorModeAndStyle();
            
    public void initColorModeAndStyle(XShape xControlShape){
        short colorMode = getColorModeOfControlShape(xControlShape);
        short style = getStyleOfControlShape(xControlShape);
        setColorModeProp(colorMode);
        setStyleProp(style);
    }
    
    abstract public void initProperties();
    
    public void setColorModeAndStyeOfControlShape(XShape xControlShape){
        setShortPropertyOfControlShape("ColorMode", xControlShape, getColorModeProp());
        setShortPropertyOfControlShape("Style", xControlShape, getStyleProp());
    }
    
    public void setColorModeOfControlShape(XShape xControlShape, short colorMode){
        setShortPropertyOfControlShape("ColorMode", xControlShape, colorMode);
    }
    
    public final void setStyleOfControlShape(XShape xControlShape, short style){
        setShortPropertyOfControlShape("Style", xControlShape, style);
    }

    public final void setShortPropertyOfControlShape(String propertyName, XShape xControlShape, short value){
        if(xControlShape != null){
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xControlShape);
            if(xText != null){
                String shapeText = "";
                shapeText = xText.getString();
                if(shapeText.equals("") || shapeText.indexOf(":") == -1){
                    if(shapeText.equals(""))
                        xText.setString(propertyName + ":" + value);
                    else
                        xText.setString(shapeText + ":" + propertyName + ":" + value);
                }else{
                    boolean isAllreadyDefined = false;
                    String[] aStr = shapeText.split(":");
                    for(int i = 0; i < aStr.length - 1; i++){
                        if(aStr[i].equals(propertyName)){
                            aStr[i+1] = "" + value;
                            isAllreadyDefined = true;
                        }
                    }
                    if(isAllreadyDefined){
                        shapeText = "";
                        for(int i = 0; i < aStr.length; i++){
                            shapeText += aStr[i];
                            if(i < aStr.length - 1)
                                shapeText += ":";
                        }
                    }else{
                        shapeText += ":" + propertyName + ":" + value;
                    }
                    xText.setString(shapeText);
                }
///*
                XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
                try {
                    //CharHidden property is useless with 3.3 LO api
                    xTextProps.setPropertyValue( "CharWeight", new Float(0.0));
                    xTextProps.setPropertyValue( "CharHeight", new Float(0.0));
                } catch (UnknownPropertyException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (PropertyVetoException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (IllegalArgumentException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println(ex.getLocalizedMessage());
                }
//*/
            }
        }
    }

    public short getColorModeOfControlShape(XShape xControlShape){
        return getShortPropertyOfControlShape("ColorMode", xControlShape);
    }

    public short getStyleOfControlShape(XShape xControlShape){
        return getShortPropertyOfControlShape("Style", xControlShape);
    }

    public short getShortPropertyOfControlShape(String propertyName, XShape xControlShape){
        if(xControlShape != null){
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xControlShape);
            String text = xText.getString();
            String sNumber = "";
            if(text.indexOf(":") == -1){
                return -1;
            }else{
                String[] aStr = text.split(":");
                for(int i = 0; i < aStr.length - 1; i++)
                    if(aStr[i].equals(propertyName))
                        sNumber = aStr[i+1];
            }
            if(sNumber.equals(""))
                return -1;
            else
                return Short.parseShort(sNumber);
        }
        return -1;
    }

    //LastHorLevel is defined version 0.9.4
    //it must be threat specially becouse of compatiblity with previous versions
    public final void setHorLevelOfControlShape(XShape xControlShape, short level){
        if(xControlShape != null){
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xControlShape);
            if(xText != null){
                String text = "";
                text = xText.getString();
                if(text.equals("") || text.indexOf(":") == -1){
                    xText.setString("LastHorLevel:" + level);
                }else{
                    boolean isAllreadyDefined = false;
                    String[] aStr = text.split(":");
                    for(int i = 0; i < aStr.length - 1; i++){
                        if(aStr[i].equals("LastHorLevel")){
                            aStr[i+1] = "" + level;
                            isAllreadyDefined = true;
                        }
                    }
                    if(isAllreadyDefined){
                        text = "";
                        for(int i = 0; i < aStr.length; i++){
                            text += aStr[i];
                            if(i < aStr.length - 1)
                                text += ":";
                        }
                    }else{
                        text += ":LastHorLevel:" + level;
                    }

                    xText.setString(text);
                }

                XPropertySet xTextProps = (XPropertySet)UnoRuntime.queryInterface( XPropertySet.class, xText.createTextCursor());
                try {
                    //CharHidden property is useless with 3.3 LO api
                    xTextProps.setPropertyValue( "CharWeight", new Float(0.0));
                    xTextProps.setPropertyValue( "CharHeight", new Float(0.0));
                } catch (UnknownPropertyException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (PropertyVetoException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (IllegalArgumentException ex) {
                    System.err.println(ex.getLocalizedMessage());
                } catch (WrappedTargetException ex) {
                    System.err.println(ex.getLocalizedMessage());
                }
  
            }
        }
    }

    public short getHorLevelOfControlShape(XShape xControlShape){
        if(xControlShape != null){
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xControlShape);
            String text = xText.getString();
            String sNumber = "";
            if(text.indexOf(":") == -1){
                sNumber = text;
            }else{
                String[] aStr = text.split(":");
                for(int i=0; i<aStr.length - 1; i++)
                    if(aStr[i].equals("LastHorLevel"))
                        sNumber = aStr[i+1];
            }
            if(sNumber.equals(""))
                return -1;
            else
                return Short.parseShort(sNumber);
        }
        return -1;
    }
    
    public void removeHorLevelPropsOfControlShape(XShape xControlShape){
        if(xControlShape != null){
            XText xText = (XText)UnoRuntime.queryInterface(XText.class, xControlShape);
            String text = xText.getString();
            if(!(text.indexOf("LastHorLevel:") == -1)){
                String newText = "";
                String[] aStr = text.split(":");
                for(int i = 0; i < aStr.length - 1; i += 2)
                    if(!aStr[i].equals("LastHorLevel"))
                        newText += aStr[i] + ":" + aStr[i+1] + ":";
                newText = newText.substring(0, newText.length()-1);
                xText.setString(newText);
            }
        }
           
          
    }
 
}
