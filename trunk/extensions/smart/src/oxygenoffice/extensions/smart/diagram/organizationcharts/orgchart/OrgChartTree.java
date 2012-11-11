package oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart;

import com.sun.star.awt.Point;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.drawing.XShape;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
import com.sun.star.uno.AnyConverter;
import com.sun.star.uno.UnoRuntime;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;


public class OrgChartTree extends OrganizationChartTree{

    public static short LASTHORLEVEL = 2;

    public OrgChartTree(OrgChart organigram){
        super(organigram);
    }

    OrgChartTree(OrgChart sOrganigram, XShape xControlShape, XShape xRootItemShape) {
        super(sOrganigram);
        setControlShape(xControlShape);
        OrgChartTreeItem.initStaticMembers();
        addToRectangles(xRootItemShape);
        m_RootItem = new OrgChartTreeItem(this, xRootItemShape, null, (short)0, 0.0);
    }

    public OrgChartTree(OrgChart organigram, OrganizationChartTree diagramTree) {
        super(organigram, diagramTree);
        OrgChartTreeItem.initStaticMembers();
        getOrgChart().setHorLevelOfControlShape(getControlShape(), OrgChartTree.LASTHORLEVEL);
        m_RootItem = new OrgChartTreeItem(this, null, diagramTree.getRootItem());
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.convertTreeItems(diagramTree.getRootItem());
    }

    public void setLastHorLevel(short level){
        LASTHORLEVEL = level;
        getOrgChart().setHorLevelOfControlShape(m_xControlShape, level);
    }

    @Override
    public void initTreeItems(){
        OrgChartTreeItem.initStaticMembers();
        m_RootItem = new OrgChartTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }


    @Override
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
                if(level <= OrgChartTree.LASTHORLEVEL){
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

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     level               = getTreeItem(xDadShape).getLevel() + 1;
        int     xPos                = -1;
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if(level <= OrgChartTree.LASTHORLEVEL){
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

    @Override
    public XShape getFirstSiblingShape(XShape xBaseShape, OrganizationChartTreeItem dad){
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
                if(level <= OrgChartTree.LASTHORLEVEL){
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

    @Override
    public void refresh(){
        OrgChartTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setMeasureProps();
        m_RootItem.display();
    }

    @Override
    public void refreshConnectorProps(){
        for(XShape xConnShape : connectorList){
            XShape xShape = getEndShapeOfConnector(xConnShape);
            short level = getTreeItem(xShape).getLevel();
            Integer start, end;
            start = new Integer(2);
            if(level <= OrgChartTree.LASTHORLEVEL){
                end = new Integer(0);
            } else{
                end = new Integer(3);
            }
            getOrgChart().setConnectorShapeProps(xConnShape, start, end);
        }
    }

    public int getEndGluePointIndex(XShape xConnShape){
        try {
            XPropertySet xProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xConnShape);
                return AnyConverter.toInt(xProps.getPropertyValue("EndGluePointIndex"));
        } catch (UnknownPropertyException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (WrappedTargetException ex) {
            System.err.println(ex.getLocalizedMessage());
        } catch (IllegalArgumentException ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        return -1;
    }
}