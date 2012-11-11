package oxygenoffice.extensions.smart.diagram.organizationcharts;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.FillStyle;
import com.sun.star.drawing.LineStyle;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.Controller;


// TreeItems represent the rectangles of the diagram
public class OrganizationChartTreeItem {

    protected OrganizationChartTree   m_DiagramTree       = null;
    protected XShape        m_xRectangleShape   = null;
    protected String        m_sRectangleName    = "";

    protected OrganizationChartTreeItem      m_Dad               = null;
    protected OrganizationChartTreeItem      m_FirstChild        = null;
    protected OrganizationChartTreeItem      m_FirstSibling      = null;

    protected short         m_Level             = -1;
    protected double        m_Pos               = -1;

    public static short     _maxLevel           = -1;


    public OrganizationChartTreeItem(OrganizationChartTree diagramTree, OrganizationChartTreeItem dad, OrganizationChartTreeItem item){
        m_DiagramTree = diagramTree;
        m_Dad = dad;
        m_xRectangleShape = item.m_xRectangleShape;
        m_sRectangleName = m_DiagramTree.getOrgChart().getShapeName(m_xRectangleShape);
    }

    public OrganizationChartTreeItem(OrganizationChartTree diagramTree, XShape xShape, OrganizationChartTreeItem dad){
        m_DiagramTree = diagramTree;
        m_xRectangleShape = xShape;
        m_sRectangleName = m_DiagramTree.getOrgChart().getShapeName(xShape);
        m_Dad = dad;
    }

    public void hideElement(){
        try {
            XPropertySet xConnProps = null;
            XPropertySet xBaseProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xRectangleShape);
            if(getDiagramTree().getOrgChart().isHiddenRootElementProp()){
                xBaseProps.setPropertyValue("FillStyle", FillStyle.NONE);
                xBaseProps.setPropertyValue("LineStyle", LineStyle.NONE);
                if(getDiagramTree().getOrgChart().getController().getDiagramType() != Controller.TABLEHIERARCHYDIAGRAM){
                    for(XShape xConnShape : getDiagramTree().connectorList){
                        if(m_xRectangleShape.equals(getDiagramTree().getStartShapeOfConnector(xConnShape))){
                            xConnProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
                            if(xConnProps != null)
                                xConnProps.setPropertyValue("LineStyle", LineStyle.NONE);
                        }
                    }
                }
                if(getDiagramTree().getOrgChart().getController().getSelectedShape().equals(m_xRectangleShape))
                    getDiagramTree().getOrgChart().getController().setSelectedShape(getFirstChild().getRectangleShape());
            }else{
                if(getDiagramTree().getOrgChart().isAnyGradientColorMode())
                    xBaseProps.setPropertyValue("FillStyle", FillStyle.GRADIENT);
                else
                    xBaseProps.setPropertyValue("FillStyle", FillStyle.SOLID);
                if(getDiagramTree().getOrgChart().isOutlineProp())
                    xBaseProps.setPropertyValue("LineStyle", LineStyle.SOLID);
                else
                    xBaseProps.setPropertyValue("LineStyle", LineStyle.NONE);
                if(getDiagramTree().getOrgChart().getController().getDiagramType() != Controller.TABLEHIERARCHYDIAGRAM){
                    for(XShape xConnShape : getDiagramTree().connectorList){
                        if(m_xRectangleShape.equals(getDiagramTree().getStartShapeOfConnector(xConnShape))){
                            xConnProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
                            if(xConnProps != null)
                                xConnProps.setPropertyValue("LineStyle", LineStyle.SOLID);
                        }
                    }
                }
            }
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

    public boolean isHiddenElement(){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, m_xRectangleShape);
            FillStyle style = (FillStyle)xProps.getPropertyValue("FillStyle");
            if(style.getValue() == FillStyle.NONE_value)
                return true;
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return false;
    }

    public void convertTreeItems(OrganizationChartTreeItem treeItem){ }

    public void setDiagramTree(OrganizationChartTree diagramTree){
        m_DiagramTree = diagramTree;
    }

    public void initTreeItems(){ };

    public void setPositionsOfItems(){ };

    public void setPosOfRect(){ };

    public void setMeasureProps(){ };

    public boolean isDad(){
        if(m_Dad == null)
            return false;
        return true;
    }

    public OrganizationChartTreeItem getDad(){
        return m_Dad;
    }

    public void setDad(OrganizationChartTreeItem dad){
        m_Dad = dad;
    }

    public boolean isFirstChild(){
        if(m_FirstChild == null)
            return false;
        return true;
    }

    public OrganizationChartTreeItem getFirstChild(){
        return m_FirstChild;
    }

    public void setFirstChild(OrganizationChartTreeItem child){
        m_FirstChild = child;
    }

    public OrganizationChartTreeItem getFirstSibling(){
        return m_FirstSibling;
    }

    public void setFirstSibling(OrganizationChartTreeItem sibling){
        m_FirstSibling = sibling;
    }

    public boolean isFirstSibling(){
        if(m_FirstSibling == null)
            return false;
        return true;
    }

    public OrganizationChartTreeItem getLastSibling(){
        if(isFirstSibling())
            return getFirstSibling().getLastSibling();
        else
            return this;
    }

    public OrganizationChartTreeItem getLastChild(){
        if(isFirstChild())
            return getFirstChild().getLastSibling();
        else
            return null;
    }

    public XShape getRectangleShape(){
        return m_xRectangleShape;
    }

    public Point getPosition(){
        return m_xRectangleShape.getPosition();
    }
    
    public void setPosition(Point point){
        m_xRectangleShape.setPosition(point);
    }

    public Size getSize(){
        return m_xRectangleShape.getSize();
    }

    public void setSize(Size size){
        try {
            m_xRectangleShape.setSize(size);
        } catch (com.sun.star.beans.PropertyVetoException ex) {
            System.err.println(ex.getLocalizedMessage());
        }

    }

    public OrganizationChartTree getDiagramTree(){
        return m_DiagramTree;
    }

    public void setPos(double pos){ };

    public double getPos(){
        return m_Pos;
    }

    public void setLevel(short level){
        m_Level = level;
        if(m_Level > _maxLevel)
            _maxLevel = m_Level;
    }

    public short getLevel(){
        return m_Level;
    }

    public OrganizationChartTreeItem getPreviousSibling(OrganizationChartTreeItem treeItem){
        OrganizationChartTreeItem previousSibling = null;
        OrganizationChartTreeItem item = null;

        if(m_FirstChild != null){
           item = m_FirstChild.getPreviousSibling(treeItem);
           if(item != null)
               previousSibling = item;
        }
        if(getFirstSibling() == treeItem)
            previousSibling = this;

        if(m_FirstSibling != null){
            item = m_FirstSibling.getPreviousSibling(treeItem);
            if(item != null)
               previousSibling = item;
        }
        return previousSibling;
    }

    public void searchItem(XShape xShape){
        if(m_FirstChild != null)
            m_FirstChild.searchItem(xShape);
        if(xShape.equals(m_xRectangleShape))
            getDiagramTree().setSelectedItem(this);
        if(m_FirstSibling != null)
            m_FirstSibling.searchItem(xShape);
    }

    public void increasePosInBranch(double x){
        if(m_FirstChild != null)
            m_FirstChild.increasePosInBranch(x);
        setPos(getPos() + x);
        if(m_FirstSibling != null)
            m_FirstSibling.increasePosInBranch(x);
    }

    public void display(){
        if(m_FirstChild != null)
            m_FirstChild.display();
        setPosOfRect();
        if(m_FirstSibling != null)
            m_FirstSibling.display();
    }

    public void setGradient(){
        if(m_FirstChild != null)
            m_FirstChild.setGradient();
        if(!getDiagramTree().getRootItem().equals(this) || !getDiagramTree().getOrgChart().isHiddenRootElementProp())
            getDiagramTree().getOrgChart().setShapesGradientColor(getRectangleShape(), getLevel(), OrganizationChartTreeItem._maxLevel + 1);
        if(m_FirstSibling != null)
            m_FirstSibling.setGradient();
    }

    public void setProperties(){
        if(m_FirstChild != null)
            m_FirstChild.setProperties();
        getDiagramTree().getOrgChart().setShapeProperties(this.m_xRectangleShape, "RectangleShape");
        if(m_FirstSibling != null)
            m_FirstSibling.setProperties();
    }

    public void removeItems(){
        if(isFirstChild())
            m_FirstChild.removeItems();

        XShape xConnShape = getDiagramTree().getDadConnectorShape(m_xRectangleShape);
        if(xConnShape != null){
            getDiagramTree().removeFromConnectors(xConnShape);
            getDiagramTree().getOrgChart().removeShapeFromGroup(xConnShape);
        }
        getDiagramTree().removeFromRectangles(m_xRectangleShape);
        getDiagramTree().getOrgChart().removeShapeFromGroup(m_xRectangleShape);

        if(isFirstSibling())
            m_FirstSibling.removeItems();
    }

    public void increaseDescendantsPosNum(short diff){
        m_FirstChild.increasePosInBranch(diff);
    }

    public void printTree(){
        if(m_FirstChild != null)
            m_FirstChild.printTree();
        if(m_FirstSibling != null)
            m_FirstSibling.printTree();
    }
 
    public int getDeepOfTreeBranch(OrganizationChartTreeItem treeItem){
        if(treeItem.m_FirstChild == null){
            return 0;
        }else{
            int x = 0;
            OrganizationChartTreeItem item = treeItem.m_FirstChild;
            while(item != null){
                int y = getDeepOfTreeBranch(item);
                if( y > x )
                    x = y;
                item = item.m_FirstSibling;
            }
            return (x + 1);
        }
    }

    public short getDeepOfItem(){
        if(isDad())
            return (short)(getDad().getDeepOfItem() + 1);
        else
            return (short)0;
    }

    public short getNumberOfItemsInBranch(OrganizationChartTreeItem treeItem){
        if(treeItem.m_FirstChild == null){
            return 1;
        }else{
            short n = 1;
            OrganizationChartTreeItem item = treeItem.m_FirstChild;
            while(item != null){
                n += getNumberOfItemsInBranch(item);
                item = item.m_FirstSibling;
            }
            return n;
        }
    }
    
}
