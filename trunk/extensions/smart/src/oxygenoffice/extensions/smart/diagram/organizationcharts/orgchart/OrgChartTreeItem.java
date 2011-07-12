package oxygenoffice.extensions.smart.diagram.organizationcharts.orgchart;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTree;
import oxygenoffice.extensions.smart.diagram.organizationcharts.OrganizationChartTreeItem;


public class OrgChartTreeItem extends OrganizationChartTreeItem{

    protected static double[]        _maxPositions;
    protected static double[]        _maxBranchPositions;
    protected static double          _maxPos                = -1.0;

    private static int              _horSpace               = 0;
    private static int              _verSpace               = 0;
    private static int              _shapeWidth             = 0;
    private static int              _shapeHeight            = 0;
    private static int              _groupPosX              = 0;
    private static int              _groupPosY              = 0;


    public OrgChartTreeItem(OrganizationChartTree diagramTree, OrganizationChartTreeItem dad, OrganizationChartTreeItem item){
        super(diagramTree, dad, item);
    }

    public OrgChartTreeItem(OrganizationChartTree diagramTree, XShape xShape, OrganizationChartTreeItem dad, short level, double num){
        super(diagramTree, xShape, dad);
        setLevel(level);
        setPos(num);
    }

    public static void initStaticMembers(){
        _maxLevel = -1;
        _maxPos = -1.0;
        _maxPositions = new double[100];
        for(int i=0; i<_maxPositions.length; i++)
            _maxPositions[i] = -1.0;
        _maxBranchPositions = new double[100];
        for(int i=0; i<_maxBranchPositions.length; i++)
            _maxBranchPositions[i] = -1.0;
    }

    @Override
    public void convertTreeItems(OrganizationChartTreeItem treeItem){
        if(treeItem.isFirstChild()){
            m_FirstChild = new OrgChartTreeItem(getDiagramTree(), this, treeItem.getFirstChild());
            m_FirstChild.convertTreeItems(treeItem.getFirstChild());
        }
        if(treeItem.isFirstSibling()){
            m_FirstSibling = new OrgChartTreeItem(getDiagramTree(), getDad(), treeItem.getFirstSibling());
            m_FirstSibling.convertTreeItems(treeItem.getFirstSibling());
        }
    }

    @Override
    public final void setPos(double pos){
        m_Pos = pos;
        if(m_Pos > _maxPositions[m_Level])
            _maxPositions[m_Level] = m_Pos;
        if(m_Pos > _maxPos)
            _maxPos = m_Pos;
//        XText xText = (XText) UnoRuntime.queryInterface(XText.class, m_xRectangleShape);
//        xText.setString(m_Level +":"+m_Pos);
    }

    @Override
    public void initTreeItems(){
        XShape xFirstChildShape = getDiagramTree().getFirstChildShape(m_xRectangleShape);
        if(xFirstChildShape != null){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = 0.0;
            if(firstChildLevel <= OrgChartTree.LASTHORLEVEL)
                firstChildPos = _maxPositions[firstChildLevel] + 1.0;
            else
                firstChildPos = m_Pos + 0.5;
            m_FirstChild = new OrgChartTreeItem(getDiagramTree(), xFirstChildShape, this, firstChildLevel , firstChildPos);
            m_FirstChild.initTreeItems();
        }

        if(m_Level == OrgChartTree.LASTHORLEVEL){
            short deep = getNumberOfItemsInBranch(this);
            if(deep > 2){
                double maxPosInLevel = _maxBranchPositions[m_Level + deep - 1];
                if(m_Pos < maxPosInLevel + 0.5){
                    if(isFirstChild()){
                        getFirstChild().increasePosInBranch(maxPosInLevel + 0.5 - m_Pos);
                        setPos(maxPosInLevel + 0.5);
                    }
                }
            }
            setMaxPosOfBranch();
        }

        XShape xFirstSiblingShape = getDiagramTree().getFirstSiblingShape(m_xRectangleShape, m_Dad);
        if(xFirstSiblingShape != null){

            short firstSiblingLevel = m_Level;
            double firstSiblingPos = m_Pos + 1.0;

            if(firstSiblingLevel > OrgChartTree.LASTHORLEVEL){
                firstSiblingPos = m_Pos;
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }
            m_FirstSibling = new OrgChartTreeItem(getDiagramTree(), xFirstSiblingShape, m_Dad, firstSiblingLevel, firstSiblingPos);
            m_FirstSibling.initTreeItems();
        }

        if(m_Level <= OrgChartTree.LASTHORLEVEL && m_Dad != null && m_Dad.getFirstChild().equals(this)){
            double newPos = 0.0;
            if(isFirstSibling())
                newPos = (_maxPositions[m_Level] + m_Pos) / 2;
            else
                newPos = m_Pos;
            if(newPos > m_Dad.getPos())
                m_Dad.setPos(newPos);
            if(newPos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - newPos);
        }
    }


    @Override
    public void setPositionsOfItems(){
        if(m_FirstChild != null){
            short firstChildLevel = (short)(m_Level + 1);
            double firstChildPos = 0.0;
            if(firstChildLevel <= OrgChartTree.LASTHORLEVEL)
                firstChildPos = _maxPositions[firstChildLevel] + 1.0;

            if(firstChildLevel > OrgChartTree.LASTHORLEVEL){
                firstChildPos = m_Pos + 0.5;
            }
            m_FirstChild.setLevel(firstChildLevel);
            m_FirstChild.setPos(firstChildPos);
            m_FirstChild.setPositionsOfItems();

        }

        if(m_Level == OrgChartTree.LASTHORLEVEL){
            short deep = getNumberOfItemsInBranch(this);
            if(deep > 2){
                double maxPosInLevel = _maxBranchPositions[m_Level + deep - 1];
                if(m_Pos < maxPosInLevel + 0.5){
                    if(isFirstChild()){
                        getFirstChild().increasePosInBranch(maxPosInLevel + 0.5 - m_Pos);
                        setPos(maxPosInLevel + 0.5);
                    }
                }
            }
            setMaxPosOfBranch(); 
        }
        
        if(m_FirstSibling != null){
            short firstSiblingLevel = m_Level;
            double firstSiblingPos = m_Pos + 1.0;

            if(firstSiblingLevel > OrgChartTree.LASTHORLEVEL){
                firstSiblingPos = m_Pos;
                firstSiblingLevel = (short)(m_Level + getNumberOfItemsInBranch(this));
            }
            m_FirstSibling.setLevel(firstSiblingLevel);
            m_FirstSibling.setPos(firstSiblingPos);
            m_FirstSibling.setPositionsOfItems();
        }

        if(m_Level <= OrgChartTree.LASTHORLEVEL && m_Dad != null && m_Dad.getFirstChild().equals(this)){
            double newPos = 0.0;
            if(isFirstSibling())
                newPos = (_maxPositions[m_Level] + m_Pos) / 2;
            else
                newPos = m_Pos;
            if(newPos > m_Dad.getPos())
                m_Dad.setPos(newPos);
            if(newPos < m_Dad.getPos())
                increasePosInBranch(m_Dad.getPos() - newPos);
        }
    }

    public void setMaxPosOfBranch(){
        System.arraycopy(_maxPositions, 0, _maxBranchPositions, 0, _maxBranchPositions.length);
        double localMax = -1.0;
        for(int i = 0; i < _maxBranchPositions.length; i++){
            if(i > OrgChartTree.LASTHORLEVEL){
                if(_maxBranchPositions[i] > localMax)
                    localMax = _maxBranchPositions[i];
                if(_maxBranchPositions[i] < localMax)
                    _maxBranchPositions[i] = localMax;
            }
        }
    }

    // set _shapeWidth, , horSpace, _shapeHeight, verSpace, _groupPosX, _groupPosY
    @Override
    public void setProps(){
        int baseShapeWidth = _shapeWidth = getDiagramTree().getControlShapeSize().Width;
        int baseShapeHeight = _shapeHeight = getDiagramTree().getControlShapeSize().Height;
        _horSpace = _verSpace = 0;
        if(_maxPos > 0){
            int horUnit = (int)(baseShapeWidth / ( _maxPos * (getDiagramTree().getOrgChart().getShapeWidth() + getDiagramTree().getOrgChart().getHorSpace()) + getDiagramTree().getOrgChart().getShapeWidth()));
            _shapeWidth = horUnit * getDiagramTree().getOrgChart().getShapeWidth();
            _horSpace = horUnit * getDiagramTree().getOrgChart().getHorSpace();
        }
        if(_maxLevel > 0){
            int verUnit = (baseShapeHeight) / ( _maxLevel * (getDiagramTree().getOrgChart().getShapeHeight() + getDiagramTree().getOrgChart().getVerSpace()) + getDiagramTree().getOrgChart().getShapeHeight());
            _shapeHeight = verUnit * getDiagramTree().getOrgChart().getShapeHeight();
            _verSpace = verUnit * getDiagramTree().getOrgChart().getVerSpace();
        }
        _groupPosX = getDiagramTree().getControlShapePos().X;
        _groupPosY = getDiagramTree().getControlShapePos().Y;
    }

    @Override
    public void setPosOfRect(){
        int xCoord = _groupPosX + (int)((_shapeWidth + _horSpace) * getPos());
        int yCoord = _groupPosY + (_shapeHeight + _verSpace) * getLevel();
        if(m_Level > OrgChartTree.LASTHORLEVEL){
            setPosition(new Point((int)(xCoord + _shapeWidth * 0.1), yCoord));
            setSize(new Size((int)(_shapeWidth * 0.9), _shapeHeight ));
        }else{
            setPosition(new Point(xCoord, yCoord));
            setSize(new Size(_shapeWidth, _shapeHeight));
        }
    }

}