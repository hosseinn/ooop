package com.example;

import com.sun.star.awt.Gradient;
import com.sun.star.awt.GradientStyle;
import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.awt.XControlContainer;
import com.sun.star.awt.XDialog;
import com.sun.star.awt.XDialogProvider2;
import com.sun.star.awt.XNumericField;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.container.NoSuchElementException;
import com.sun.star.container.XNameAccess;
import com.sun.star.deployment.XPackageInformationProvider;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.XDrawPage;
import com.sun.star.drawing.XDrawView;
import com.sun.star.drawing.XShape;
import com.sun.star.frame.XFrame;
import com.sun.star.frame.XModel;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;


class DrawStar {

    XComponentContext   m_xContext          = null;
    XFrame              m_xFrame            = null;
    XDialog             m_xStarMakerDialog  = null;

    public DrawStar(XComponentContext xContext, XFrame xFrame) {
        this.m_xContext = xContext;
        this.m_xFrame   = xFrame;
    }

    public void showStarMakerDialog(){
        if(m_xStarMakerDialog  == null)
            createStarMakerDialog();
        if(m_xStarMakerDialog  != null)
            executeStarMakerDialog();
    }

    public void createStarMakerDialog(){
        try {
            String sPackageURL = getPackageLocation();
            String sDialogURL = sPackageURL + "/dialogs/NewStar.xdl";
            XDialogProvider2 xDialogProv = getDialogProvider();
            m_xStarMakerDialog = xDialogProv.createDialog(sDialogURL);
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public String getPackageLocation(){
        String location = null;
        try {
            XNameAccess xNameAccess = (XNameAccess) UnoRuntime.queryInterface(XNameAccess.class, m_xContext );
            Object oPIP = xNameAccess.getByName("/singletons/com.sun.star.deployment.PackageInformationProvider");
            XPackageInformationProvider xPIP = (XPackageInformationProvider) UnoRuntime.queryInterface(XPackageInformationProvider.class, oPIP);
            /* you can find the project location in
             * nbproject/project-uno.properties registration.classname property
             * here: com.example.StarMaker
             */
            location =  xPIP.getPackageLocation("com.example.StarMaker");
        } catch (NoSuchElementException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return location;
    }

    public XDialogProvider2 getDialogProvider(){
        XDialogProvider2 xDialogProv = null;
        try {
            XModel xModel = m_xFrame.getController().getModel();
            XMultiComponentFactory  xMCF = m_xContext.getServiceManager();
            Object obj;
            if (xModel != null) {
                Object[] args = new Object[1];
                args[0] = xModel;
                obj = xMCF.createInstanceWithArgumentsAndContext("com.sun.star.awt.DialogProvider2", args, m_xContext);
            } else {
                obj = xMCF.createInstanceWithContext("com.sun.star.awt.DialogProvider2", m_xContext);
            }
            xDialogProv = (XDialogProvider2) UnoRuntime.queryInterface(XDialogProvider2.class, obj);
        }catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xDialogProv;
    }

    public void executeStarMakerDialog(){
        short ok = m_xStarMakerDialog.execute();
        if(ok == 1){
            XControlContainer xControlContainer = (XControlContainer) UnoRuntime.queryInterface(XControlContainer.class, m_xStarMakerDialog);
            XNumericField pointsField = (XNumericField) UnoRuntime.queryInterface(XNumericField.class, xControlContainer.getControl("PointsField"));
            int numOfPoints = (int)pointsField.getValue();
            XNumericField ratioField = (XNumericField) UnoRuntime.queryInterface(XNumericField.class, xControlContainer.getControl("RatioField"));
            double ratio = ratioField.getValue() / 100;
            XNumericField sizeField = (XNumericField) UnoRuntime.queryInterface(XNumericField.class, xControlContainer.getControl("SizeField"));
            int size = (int)sizeField.getValue();
            drawStar(numOfPoints, ratio, size);
        }
    }

    public void drawStar(int numOfPoints, double ratio, int size){
        try {
            Size sSize = new Size(size * 100, size * 100);
            Point position = new Point(1000, 1000);
            Point middlePoint = new Point((int) (position.X + sSize.Width / 2), (int) (position.Y + sSize.Height / 2));
            Point[] points = new Point[numOfPoints*2];
            Point[] endPoints = new Point[2];
            double x;
            double y;
            double angle;
            for (int i = 0; i < numOfPoints; i++) {
                angle = i * Math.PI * 2 / numOfPoints - Math.PI / 2;
                x = middlePoint.X + sSize.Width / 2 * Math.cos(angle);
                y = middlePoint.Y + sSize.Height / 2 * Math.sin(angle);
                points[i * 2] = new Point((int) x, (int)y);
                if(i == 0)
                    endPoints[0] = new Point((int) x, (int)y);
                angle = (i + 0.5) * Math.PI * 2 / numOfPoints - Math.PI / 2;
                x = middlePoint.X + ratio * sSize.Width / 2 * Math.cos(angle);
                y = middlePoint.Y + ratio * sSize.Height / 2 * Math.sin(angle);
                points[i * 2 + 1] = new Point((int) x, (int)y);
                if(i == numOfPoints - 1)
                    endPoints[1] = new Point((int) x, (int)y);
            }
            XShape xShape = createShape("PolyPolygonShape");
            XDrawPage xPage = getCurrentPage();
            xPage.add(xShape);
            XPropertySet xProp = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xShape);
            Point[][] polyPoints = {points, endPoints};
            xProp.setPropertyValue("PolyPolygon", polyPoints);
            setGradient(xShape, GradientStyle.ELLIPTICAL, 0xffff00, 0xffbb00, (short)45, (short)5, (short)50, (short)50, (short)100, (short)100);//, short angle, short border, short xOffset, short yOffset, short startIntensity, short endIntensity);
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public XShape createShape(String type){
        XShape xShape = null;
        try {
            XMultiServiceFactory xMSF = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, m_xFrame.getController().getModel());
            xShape = (XShape) UnoRuntime.queryInterface(XShape.class, xMSF.createInstance ("com.sun.star.drawing." + type ));

        }  catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xShape;
    }

    public XDrawPage getCurrentPage(){
        XDrawView xDrawView = (XDrawView)UnoRuntime.queryInterface(XDrawView.class, m_xFrame.getController());
        return xDrawView.getCurrentPage();
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

}
