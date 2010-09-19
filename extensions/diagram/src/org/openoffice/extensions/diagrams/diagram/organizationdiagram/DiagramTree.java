
package org.openoffice.extensions.diagrams.diagram.organizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.drawing.XShapes;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import java.util.ArrayList;


public class DiagramTree {

    private OrganizationDiagram m_Organigram        = null;
    private ArrayList<XShape>   rectangleList       = null;
    private ArrayList<XShape>   connectorList       = null;
    private XShapes             m_xShapes           = null;

    private XShape              m_xControlShape     = null;
    private XShape              m_xRootShape        = null;
    private TreeItem            m_RootItem          = null;
    private TreeItem            m_SelectedItem      = null;


    public DiagramTree(OrganizationDiagram organigram){
        m_Organigram    = organigram;
        rectangleList   = new ArrayList<XShape>();
        connectorList   = new ArrayList<XShape>();
        m_xShapes       = m_Organigram.getShapeGroup();
    }

    public OrganizationDiagram getOrganigram(){
        return m_Organigram;
    }

    public void clear(){
        if(rectangleList != null)
            rectangleList.clear();
        if(connectorList != null)
            connectorList.clear();
    }

    public void setControlShape(XShape xShape){
        m_xControlShape = xShape;
    }

    public XShape getControlShape(){
        return m_xControlShape;
    }

    public void setControlShapeSize(Size size){
        try {
            m_xControlShape.setSize(size);
        } catch (PropertyVetoException ex) {
            System.err.println("PropertyVetoException in OrganizationDiagram.setControlShapeSize(). Message:\n" + ex.getLocalizedMessage());
        }
    }

    public Size getControlShapeSize(){
        return m_xControlShape.getSize();
    }

    public void setControlShapePos(Point point){
            m_xControlShape.setPosition(point);
    }

    public Point getControlShapePos(){
        return m_xControlShape.getPosition();
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

    public void printListsOfDiagramTree(){
        System.out.println("----------------------------------------");
        for(int i = 0; i < rectangleList.size(); i++ )
            System.out.println(getOrganigram().getShapeName(rectangleList.get(i)));
        for(int i = 0; i < connectorList.size(); i++ )
            System.out.println(getOrganigram().getShapeName(connectorList.get(i)));
        System.out.println("----------------------------------------");
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
                if(!startShape.equals(m_xRootShape))
                    m_xShapes.remove(startShape);
                m_xShapes.remove(xConnShape);
            }
        }
    }

    public void setTree(){

        m_xRootShape = null;
        setRootItem();
        
        if(m_xRootShape == null){
            String title = getOrganigram().getGui().getDialogPropertyValue("Strings", "RoutShapeError.Title");
            String message = getOrganigram().getGui().getDialogPropertyValue("Strings", "RoutShapeError.Message");
            getOrganigram().getGui().showMessageBox(title, message);
        }else{
            TreeItem._maxLevel = TreeItem._maxPos = TreeItem._level1MaxPos = TreeItem._level2MaxPos = -1;
            m_RootItem = new TreeItem(this, m_xRootShape, null, (short)0, (short)0);
            m_RootItem.initTreeItems();
        }
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

    public int getShapeID(XShape xShape){
        return getOrganigram().getController().getNumberOfShape(getOrganigram().getShapeName(xShape));
    }

    public XShape getStartShapeOfConnector(XShape xConnShape){
        XShape xStartShape = null;
        try {
            XPropertySet xPropSet = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
            Object object = xPropSet.getPropertyValue("StartShape");
            xStartShape = (XShape) UnoRuntime.queryInterface(XShape.class, object);
        } catch (UnknownPropertyException ex) {
            System.err.println("UnknownPropertyException in OrganizationDiagram.getStartShapeOfConnector(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.getStartShapeOfConnector(). Message:\n" + ex.getLocalizedMessage());
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
            System.err.println("UnknownPropertyException in OrganizationDiagram.getEndShapeOfConnector(). Message:\n" + ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println("WrappedTargetException in OrganizationDiagram.getEndShapeOfConnector(). Message:\n" + ex.getLocalizedMessage());
        }
        return xEndShape;
    }

     public XShape getDadConnectorShape(XShape xRectShape){
        for(XShape xConnShape : connectorList)
            if(xRectShape.equals(getEndShapeOfConnector(xConnShape)))
                return xConnShape;
        return null;
    }

    public XShape getChildConnectorShape(XShape xRectShape){
        for(XShape xConnShape : connectorList)
            if(xRectShape.equals(getStartShapeOfConnector(xConnShape)))
                return xConnShape;
        return null;
    }

    public XShape getFirstChildShape(XShape xDadShape){
        // the struct of diagram change below second level that's why we need level of shape
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xPos == -1 || xChildeShape.getPosition().X < xPos){
                        xPos = xChildeShape.getPosition().X;
                        xFirstChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y < yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xFirstChildShape = xChildeShape;
                    }
                }
            }
        }
        return xFirstChildShape;
    }

    public XShape getLastChildShape(XShape xDadShape){

        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xPos == -1 || xChildeShape.getPosition().X > xPos){
                        xPos = xChildeShape.getPosition().X;
                        xLastChildShape = xChildeShape;
                    }
                }else{
                    if( yPos == -1 || xChildeShape.getPosition().Y > yPos){
                        yPos = xChildeShape.getPosition().Y;
                        xLastChildShape = xChildeShape;
                    }

                }
            }
        }
        return xLastChildShape;
    }

    public XShape getFirstSiblingShape(XShape xBaseShape, TreeItem dad){
        if(dad == null)
            return null;
        if(dad.getRectangleShape() == null)
            return null;
        int    level                = dad.getLevel() + 1;
        XShape xDadShape            = dad.getRectangleShape();
        XShape xSiblingShape        = null;
        XShape xFirstSiblingShape   = null;
        Point  baseShapePos         = xBaseShape.getPosition();
        int    xPos                 = -1;
        int    yPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if(level <= 2){
                    if( xSiblingShape.getPosition().X > baseShapePos.X){
                        if( xPos == -1 || xSiblingShape.getPosition().X < xPos){
                            xPos = xSiblingShape.getPosition().X;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }else{
                    if( xSiblingShape.getPosition().Y > baseShapePos.Y){
                        if( yPos == -1 || xSiblingShape.getPosition().Y < yPos){
                            yPos = xSiblingShape.getPosition().Y;
                            xFirstSiblingShape = xSiblingShape;
                        }
                    }
                }
            }
        }
        return xFirstSiblingShape;
    }

    public TreeItem getTreeItem(XShape xShape){
        if(m_xRootShape != null){
            if(xShape.equals(m_xRootShape))
                return m_RootItem;
            m_RootItem.searchItem(xShape);
        }
        return m_SelectedItem;
    }

    public TreeItem getPreviousSibling(TreeItem treeItem){
        return m_RootItem.getPreviousSibling(treeItem);
    }

    public void setSelectedItem(TreeItem treeItem){
        m_SelectedItem = treeItem;
    }

    public void refresh(){
        TreeItem._maxLevel = 0; TreeItem._maxPos = TreeItem._level1MaxPos = TreeItem._level2MaxPos = -1;
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
        m_RootItem.display();
    }

    public XShapes getGroupShapes(){
        return m_xShapes;
    }

}
