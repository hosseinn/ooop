package oxygenoffice.extensions.smart.diagram.processes.upwardarrowprocess;

import com.sun.star.awt.Point;
import com.sun.star.awt.Size;
import com.sun.star.drawing.XShape;
import oxygenoffice.extensions.smart.diagram.Color;
import oxygenoffice.extensions.smart.diagram.processes.ProcessDiagramItem;


public class UpwardArrowProcessItem extends ProcessDiagramItem {

    UpwardArrowProcessItem(UpwardArrowProcess uaProcess, int id, XShape xRectangleShape, XShape xEllipseShape) {
        super(uaProcess, id, xRectangleShape, xEllipseShape);
    }

    public void setPosition(int numOfItems, Size controlShapeSize, Point controlShapePos) {
        try {
            int diffX = controlShapeSize.Width / 20; //0;
            int diffY = controlShapeSize.Height / 10;
            int ellipseCenterX = controlShapePos.X + controlShapeSize.Width + diffX;
            int ellipseWidth = controlShapeSize.Width; // - controlShapeSize.Width / 16;
            int ellipseCenterY = controlShapePos.Y + controlShapeSize.Height + diffY;
            int ellipseHeight = controlShapeSize.Height / (UpwardArrowProcess._ArrowEndMeasure + UpwardArrowProcess._HeightMeasure) * UpwardArrowProcess._HeightMeasure;
            Point ellipsePoint = new Point(ellipseCenterX, ellipseCenterY);
            Size ellipseSize = new Size(ellipseWidth + diffX , ellipseHeight + diffY);
            double startAngle = - Math.PI + Math.PI / 6;
            double stripAngle = Math.PI / 4; //Math.PI / 2 - angleDiff * 4;
            double diff = 0;
            if(numOfItems > 1)
                diff = stripAngle / (numOfItems - 1);
            int width = ellipseWidth / 50;
            if(width > ellipseHeight / 25)
                width = ellipseHeight / 25;
            int widthDiff = width * 4 / numOfItems;
            
            if(numOfItems == 1){
                startAngle += stripAngle * 2 / 3;
                width *= 3;
            }
            if(numOfItems == 2){
                startAngle += stripAngle / 3;
                diff = stripAngle / 2;
                width *= 2;
            }
            if(numOfItems == 3){
                startAngle += stripAngle / 4;
                diff = stripAngle / 3;
                widthDiff = width * 3 / numOfItems;
                width *= 2;
            }
            
            Size size = new Size(width + (id - 1) * widthDiff, width + (id - 1) * widthDiff);
            Point point = new Point((int)(ellipsePoint.X + (ellipseSize.Width + size.Width / 2) * Math.cos(startAngle + (id -1) * diff)),(int)(ellipsePoint.Y + (ellipseSize.Height + size.Height / 2) * Math.sin(startAngle + (id -1) * diff)));
                
            if(xSecondShape != null) {
                xSecondShape.setPosition(point);
                xSecondShape.setSize(size);
            }
            if(xMainShape != null) {
                if(id < 6)
                    xMainShape.setPosition(new Point(point.X + size.Width / 2, point.Y + size.Height * 14 / 10));
                else if(id == 6)
                    xMainShape.setPosition(new Point(point.X + size.Width / 2, point.Y + size.Height * 13 / 10));
                else if(id == 7)
                    xMainShape.setPosition(new Point(point.X + size.Width / 2, point.Y + size.Height * 23 / 20));
                else
                    xMainShape.setPosition(new Point(point.X + size.Width / 2, point.Y + size.Height));
                xMainShape.setSize(new Size(ellipseSize.Width / 6, ellipseSize.Height / 20));
            }
        } catch (Exception ex) {
            System.err.println(ex.getLocalizedMessage());
        }
        
    }
    
    @Override
    public void setColor(Color oColor){
        if(getProcess() != null){
            if(oColor.isGradient()){
                getProcess().setGradient(xSecondShape, oColor.getStartColor(), oColor.getEndColor());
            }else{
                getProcess().setColorOfShape(xSecondShape, oColor.getColor());
            }
        }
    }
    
    @Override
    public void setLineColor(){
        if(getProcess() != null){
            getProcess().setLineColorOfShape(xSecondShape);
        }
    }
    
}
