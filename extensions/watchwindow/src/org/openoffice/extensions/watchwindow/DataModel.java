package org.openoffice.extensions.watchwindow;

import com.sun.star.sheet.XSpreadsheet;
import java.util.ArrayList;
import java.util.List;


public class DataModel {
    
    private Controller m_Controller;
    private List<WatchedCell>  m_list; 
    
    public DataModel(Controller controller){
        m_Controller = controller;
        m_list = new ArrayList<WatchedCell>();
    }
    
    public Controller getController(){
        return m_Controller;
    }
    
    public void addToDataList(XSpreadsheet xValidSheet, String selectedSheetName, String selectedCellName) {   
        boolean bool = true;
        short num = 0;
        for(WatchedCell item: m_list){
            if(item.getSheetName().equals(selectedSheetName) && item.getCellName().equals(selectedCellName)){
                bool = false;
                num = item.getNum();
                if(num == (short)-1){
                    item.setNum( getController().getNumer());
                    getController().increaseNumer();
                    getController().addToListBox( item.toString(), item.getNum());
                }
            }
         }
         if(bool){
            m_list.add(new WatchedCell(getController(), xValidSheet, selectedSheetName, selectedCellName, getController().getNumer()));
            getController().increaseNumer(); 
        }
    }
    
    public void removeFromDataList(int i){
        //change id of removed element ( m_num = -1 )
        // decrease the next element in the list ( m_num-- )
        short num = 0;
        for(WatchedCell item: m_list){
            num = item.getNum();
            if(num == i){
                item.setNum((short)-1);
            }
            if(num > i)
                item.setNum(--num);
        }
    }
    
    public void removeSheetItemsFromDataList( String sheetName) { 
        short db = 0;
        short num = 0;
        for(WatchedCell item: m_list){
            num = item.getNum();
            if(num > -1){
                if(item.getSheetName().equals(sheetName)){
                    getController().removeFromListBox(num, (short)1);
                    item.setNum((short)-1);
                    getController().decreaseNumer();
                    db++;
                }else{
                    num -= db;
                    item.setNum(num); 
                    getController().removeFromListBox(item.getNum(), (short)1);
                }
            }
        }
    }
    
} 