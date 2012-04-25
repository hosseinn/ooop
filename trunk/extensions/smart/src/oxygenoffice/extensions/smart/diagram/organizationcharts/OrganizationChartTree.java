package oxygenoffice.extensions.smart.diagram.organizationcharts;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.lang.IndexOutOfBoundsException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.text.XText;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;


public abstract class OrganizationChartTree {


    protected OrganizationChart     m_OrgChart          = null;
    protected ArrayList<XShape>     rectangleList       = null;
    protected ArrayList<XShape>     connectorList       = null;

    protected XShapes               m_xShapes           = null;
    protected XShape                m_xControlShape     = null;
    protected XShape                m_xRootShape        = null;

    protected OrganizationChartTreeItem              m_RootItem          = null;
    protected OrganizationChartTreeItem              m_SelectedItem      = null;


    public OrganizationChartTree(OrganizationChart orgChart){
        m_OrgChart      = orgChart;
        rectangleList   = new ArrayList<XShape>();
        connectorList   = new ArrayList<XShape>();
        m_xShapes       = m_OrgChart.getShapes();
    }

    public OrganizationChartTree(OrganizationChart orgChart, OrganizationChartTree diagramTree){
        m_OrgChart      = orgChart;
        rectangleList   = diagramTree.rectangleList;
        connectorList   = diagramTree.connectorList;
        m_xShapes       = m_OrgChart.getShapes();
        m_xControlShape = diagramTree.m_xControlShape;
        XText xText = (XText)UnoRuntime.queryInterface(XText.class, m_xControlShape);
        xText.setString("");
        m_xRootShape    = diagramTree.m_xRootShape;
    }

    public abstract void initTreeItems();

    public abstract XShape getFirstChildShape(XShape xDadShape);

    public abstract XShape getLastChildShape(XShape xDadShape);

    public abstract XShape getFirstSiblingShape(XShape xBaseShape, OrganizationChartTreeItem dad);

    public abstract void refresh();

    public OrganizationChart getOrgChart(){
        return m_OrgChart;
    }

    public OrganizationChartTreeItem getRootItem(){
        return m_RootItem;
    }

    public abstract void refreshConnectorProps();

    public void setGradientColorProps(){
        if(getOrgChart().isPreDefinedGradientsProps() && getOrgChart().getStyle() != OrganizationChart.USER_DEFINE)
            m_RootItem.setGradient();
    }

    // set root Item, return number of roots (if number is not 1, than there is error
    public short setRootItem(){
        boolean isRoot;
        short numOfRoots = 0;
        //  search root shape
        for(XShape xRectangleShape : rectangleList){
            isRoot = true;
            for(XShape xConnShape : connectorList)
                if(xRectangleShape.equals(getEndShapeOfConnector(xConnShape)))
                    isRoot = false;
            if(isRoot){
                numOfRoots ++;
                if(m_xRootShape == null)
                    m_xRootShape = xRectangleShape;
                else
                    if(xRectangleShape.getPosition().Y < m_xRootShape.getPosition().Y)
                        m_xRootShape = xRectangleShape;
            }
        }
        return numOfRoots;
    }

    public void setControlShape(XShape xShape){
        m_xControlShape = xShape;
    }

    public XShape getControlShape(){
        return m_xControlShape;
    }

    public void addToRectangles(XShape xShape){
        rectangleList.add(xShape);
    }

    public void removeFromRectangles(XShape xShape){
        rectangleList.remove(xShape);
    }

    public int getRectangleListSize(){
        return rectangleList.size();
    }

    public boolean isInRectangleList(XShape xShape){
        return rectangleList.contains(xShape);
    }

    public void addToConnectors(XShape xShape){
        connectorList.add(xShape);
    }

    public void removeFromConnectors(XShape xShape){
        connectorList.remove(xShape);
    }

    public int getConnectorListSize(){
        return connectorList.size();
    }

    public boolean isInConnectorList(XShape xShape){
        return connectorList.contains(xShape);
    }

    public void setLists(){
        try {
            clearLists();
            XShape xCurrShape = null;
            String currShapeName = "";

            for(int i=0; i < m_xShapes.getCount(); i++){
                xCurrShape = (XShape) UnoRuntime.queryInterface(XShape.class, m_xShapes.getByIndex(i));
                currShapeName = getOrgChart().getShapeName(xCurrShape);
                if (currShapeName.contains("RectangleShape")) {
                    if (currShapeName.endsWith("RectangleShape0")) {
                        setControlShape(xCurrShape);
                    }else{
                        addToRectangles(xCurrShape);
                    }
                }
                if (currShapeName.contains("ConnectorShape"))
                    addToConnectors(xCurrShape);
            }
        } catch (IndexOutOfBoundsException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
    }

    public void clearLists(){
        if(rectangleList != null)
            rectangleList.clear();
        if(connectorList != null)
            connectorList.clear();
    }

    public void setTree(){
        m_xRootShape = null;
        int error = setRootItem();

        if(m_xRootShape == null || error > 1){
            String title = getOrgChart().getGui().getDialogPropertyValue("Strings", "RoutShapeError.Title");
            String message = getOrgChart().getGui().getDialogPropertyValue("Strings", "RoutShapeError.Message");
            getOrgChart().getGui().showMessageBox(title, message);
        }else{
            initTreeItems();
        }
    }

    public int getShapeID(XShape xShape){
        return getOrgChart().getController().getShapeID(getOrgChart().getShapeName(xShape));
    }

    public XShape getStartShapeOfConnector(XShape xConnShape){
        XShape xStartShape = null;
        try {
            XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
            Object object = xPropSet.getPropertyValue("StartShape");
            xStartShape = (XShape) UnoRuntime.queryInterface(XShape.class, object);
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xStartShape;
    }

    public XShape getEndShapeOfConnector(XShape xConnShape){
        XShape xEndShape = null;
        try {
            XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
            Object object = xPropSet.getPropertyValue("EndShape");
            xEndShape = (XShape) UnoRuntime.queryInterface(XShape.class, object);
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return xEndShape;
    }

    public XShape getDadConnectorShape(XShape xRectShape){
        for(XShape xConnShape : connectorList)
            if(xRectShape.equals(getEndShapeOfConnector(xConnShape)))
                return xConnShape;
        return null;
    }

    public void repairTree(){

        removeAdditionalRoots();

        XShape startShape = null;
        XShape endShape = null;

        for(XShape xConnShape : connectorList){

            startShape = getStartShapeOfConnector(xConnShape);
            endShape = getEndShapeOfConnector(xConnShape);

            if(startShape == null){
                if(!endShape.equals(m_xRootShape))
                    m_xShapes.remove(endShape);
                m_xShapes.remove(xConnShape);
            }

            if(endShape == null){
                m_xShapes.remove(xConnShape);
            }
        }
    }

    public void removeAdditionalRoots(){

        boolean isRoot;

        for(XShape xRectangleShape : rectangleList){
            isRoot = true;
            for(XShape xConnShape : connectorList)
                if(xRectangleShape.equals(getEndShapeOfConnector(xConnShape)))
                    isRoot = false;
            if(isRoot)
                if(!xRectangleShape.equals(m_xRootShape))
                    m_xShapes.remove(xRectangleShape);
        }
    }

    public boolean isValidConnectors(){

        boolean isValid = false;
        XShape startShape = null;
        XShape endShape = null;

        for(XShape xConnShape : connectorList){

            startShape = getStartShapeOfConnector(xConnShape);
            isValid = false;

            for(XShape xRectShape : rectangleList)
                if(startShape.equals(xRectShape))
                    isValid = true;
            if(!isValid)
                return false;

            endShape = getEndShapeOfConnector(xConnShape);
            isValid = false;

            for(XShape xRectShape : rectangleList)
                if(endShape.equals(xRectShape))
                    isValid = true;
            if(!isValid)
                return false;
        }
        return true;
    }

    public void setControlShapeSize(Size size){
        try {
            m_xControlShape.setSize(size);
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }

    }

    public Size getControlShapeSize(){
        if(m_xControlShape != null)
            return m_xControlShape.getSize();
        else
            return null;
    }

    public void setControlShapePos(Point point){
            m_xControlShape.setPosition(point);
    }

    public Point getControlShapePos(){
        return m_xControlShape.getPosition();
    }

    public OrganizationChartTreeItem getTreeItem(XShape xShape){
        if(m_xRootShape != null){
            if(xShape.equals(m_xRootShape))
                return m_RootItem;
            m_RootItem.searchItem(xShape);
        }
        return m_SelectedItem;
    }

    public void setSelectedItem(OrganizationChartTreeItem treeItem){
        m_SelectedItem = treeItem;
    }

    public OrganizationChartTreeItem getPreviousSibling(OrganizationChartTreeItem treeItem){
        return m_RootItem.getPreviousSibling(treeItem);
    }
    
}
