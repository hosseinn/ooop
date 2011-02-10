package org.openoffice.extensions.diagrams.diagram.organizationcharts.horizontalorganizationdiagram;

import com.sun.star.awt.Point;
import com.sun.star.drawing.XShape;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.DiagramTree;
import org.openoffice.extensions.diagrams.diagram.organizationcharts.TreeItem;


public class HDiagramTree extends DiagramTree{
   

    HDiagramTree(HorizontalOrganizationDiagram hOrganigram) {
        super(hOrganigram);
    }

    HDiagramTree(HorizontalOrganizationDiagram hOrganigram, DiagramTree diagramTree) {
        super(hOrganigram, diagramTree);
        HTreeItem.initStaticMembers();
        m_RootItem = new HTreeItem(this, null, diagramTree.getRootItem());
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.convertTreeItems(diagramTree.getRootItem());
    }

    @Override
    public void initTreeItems(){
        HTreeItem.initStaticMembers();
        m_RootItem = new HTreeItem(this, m_xRootShape, null, (short)0, (short)0);
        m_RootItem.initTreeItems();
    }

    @Override
    public void refresh(){
        HTreeItem.initStaticMembers();
        m_RootItem.setLevel((short)0);
        m_RootItem.setPos(0.0);
        m_RootItem.setPositionsOfItems();
        m_RootItem.setProps();
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
    public XShape getFirstSiblingShape(XShape xBaseShape, TreeItem dad){
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
