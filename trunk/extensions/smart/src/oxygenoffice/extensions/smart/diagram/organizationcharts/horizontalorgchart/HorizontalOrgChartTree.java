package oxygenoffice.extensions.smart.diagram.organizationcharts.horizontalorgchart;

import com.sun.star.awt.Point;
import com.sun.star.drawing.XShape;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;


public class HorizontalOrgChartTree extends OrganizationChartTree{
   

    HorizontalOrgChartTree(HorizontalOrgChart hOrganigram) {
        super(hOrganigram);
    }

    HorizontalOrgChartTree(HorizontalOrgChart sOrganigram, XShape xControlShape, XShape xRootItemShape) {
        super(sOrganigram);
        setControlShape(xControlShape);
        HorizontalOrgChartTreeItem.initStaticMembers();
        addToRectangles(xRootItemShape);
        m_RootItem = new HorizontalOrgChartTreeItem(this, xRootItemShape, null, (short)0, 0.0);
    }

    HorizontalOrgChartTree(HorizontalOrgChart hOrganigram, OrganizationChartTree diagramTree) {
        super(hOrganigram, diagramTree);
        HorizontalOrgChartTreeItem.initStaticMembers();
        m_RootItem = new HorizontalOrgChartTreeItem(this, null, diagramTree.getRootItem());
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.convertTreeItems(diagramTree.getRootItem());
    }

    @Override
    public void initTreeItems(){
        HorizontalOrgChartTreeItem.initStaticMembers();
        m_RootItem = new HorizontalOrgChartTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public void refresh(){
        HorizontalOrgChartTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setMeasureProps();
        m_RootItem.display();
    }

    @Override
    public XShape getFirstChildShape(XShape xDadShape){
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xFirstChildShape    = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( yPos == -1 || xChildeShape.getPosition().Y > yPos){
                    yPos = xChildeShape.getPosition().Y;
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
        int    yPos                 = -1;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xSiblingShape = getEndShapeOfConnector(xConnShape);
                if( xSiblingShape.getPosition().Y < baseShapePos.Y){
                    if( yPos == -1 || xSiblingShape.getPosition().Y > yPos){
                        yPos = xSiblingShape.getPosition().Y;
                        xFirstSiblingShape = xSiblingShape;
                    }
                }

            }
        }
        return xFirstSiblingShape;
    }

    @Override
    public XShape getLastChildShape(XShape xDadShape){
        int     yPos                = -1;
        XShape  xChildeShape        = null;
        XShape  xLastChildShape     = null;

        for(XShape xConnShape : connectorList){
            if(xDadShape.equals(getStartShapeOfConnector(xConnShape))){
                xChildeShape = getEndShapeOfConnector(xConnShape);
                if( yPos == -1 || xChildeShape.getPosition().Y < yPos){
                    yPos = xChildeShape.getPosition().Y;
                    xLastChildShape = xChildeShape;
                }
            }
        }
        return xLastChildShape;
    }

    @Override
    public void refreshConnectorProps(){
        for(XShape xConnShape : connectorList)
            getOrgChart().setConnectorShapeProps(xConnShape, new Integer(1), new Integer(3));
    }


}
