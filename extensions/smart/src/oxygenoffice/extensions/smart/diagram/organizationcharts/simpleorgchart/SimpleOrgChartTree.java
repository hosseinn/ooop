package oxygenoffice.extensions.smart.diagram.organizationcharts.simpleorgchart;

import com.sun.star.awt.Point;
import com.sun.star.drawing.XShape;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;


public class SimpleOrgChartTree extends OrganizationChartTree{

    SimpleOrgChartTree(SimpleOrgChart sOrganigram) {
        super(sOrganigram);
    }

    SimpleOrgChartTree(SimpleOrgChart sOrganigram, OrganizationChartTree diagramTree) {
        super(sOrganigram, diagramTree);
        SimpleOrgChartTreeItem.initStaticMembers();
        m_RootItem = new SimpleOrgChartTreeItem(this, null, diagramTree.getRootItem());
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.convertTreeItems(diagramTree.getRootItem());
    }

    @Override
    public void initTreeItems(){
        SimpleOrgChartTreeItem.initStaticMembers();
        m_RootItem = new SimpleOrgChartTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public XShape getFirstChildShape(XShape xDadShape){
        int     xPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( xPos == -1 || xChildeShape.getPosition().X < xPos){
                    xPos = xChildeShape.getPosition().X;
                    xFirstChildShape = xChildeShape;
                }
            }
        }
        return xFirstChildShape;
    }

    @Override
    public XShape getFirstSiblingShape(XShape xBaseShape, OrganizationChartTreeItem dad){
        if(dad == null)
            return null;
        if(dad.getRectangleShape() == null)
            return null;
        XShape xDadShape            = dad.getRectangleShape();
        XShape xSiblingShape        = null;
        XShape xFirstSiblingShape   = null;
        Point  baseShapePos         = xBaseShape.getPosition();
        int    xPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if( xSiblingShape.getPosition().X > baseShapePos.X){
                    if( xPos == -1 || xSiblingShape.getPosition().X < xPos){
                        xPos = xSiblingShape.getPosition().X;
                        xFirstSiblingShape = xSiblingShape;
                    }
                }

            }
        }
        return xFirstSiblingShape;
    }

    @Override
    public void refresh(){
        SimpleOrgChartTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
        m_RootItem.display();
    }

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     xPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( xPos == -1 || xChildeShape.getPosition().X > xPos){
                    xPos = xChildeShape.getPosition().X;
                    xLastChildShape = xChildeShape;
                }
            }
        }
        return xLastChildShape;
    }

    @Override
    public void refreshConnectorProps(){
        for(XShape xConnShape : connectorList)
            getOrgChart().setConnectorShapeProps(xConnShape, new Integer(2), new Integer(0));
    }

}
