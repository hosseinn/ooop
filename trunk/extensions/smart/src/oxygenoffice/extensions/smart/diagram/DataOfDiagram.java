package oxygenoffice.extensions.smart.diagram;

import java.util.ArrayList;

public class DataOfDiagram {

    public class Data {
        
        private short level;
        private String value;
        
        public Data(short level, String value){
            this.level = level;
            this.value = value;
        }
        
        public short getLevel(){
            return level;
        }
        
        public String getValue(){
            return value;
        }
        
        public void increase(){
            level++;
        }
    }

    private ArrayList<Data> datas;

    public DataOfDiagram(){
        datas = new ArrayList<Data>();
    }

    public void add(short level, String value){
        datas.add(new Data(level, value));
    }
    
    public Data get(int i){
        return datas.get(i);
    }

    public boolean isOneFirstLevelData(){
        int m = 0;
        int size = size();
        for(int i = 0; i < size; i++){
            if(get(i).getLevel() == 0)
                ++m;
            if(m > 1)
                return false;
        }
        return true;
    }

    public void increaseLevels(){
        for(int i = 0; i < datas.size(); i++)
            datas.get(i).increase();
    }

    public int size(){
        return datas.size();
    }

    public boolean isEmpty(){
        return datas.isEmpty();
    }

    public void print(){
        if(!isEmpty()){
            int max = size();
            for(int i=0; i<max; i++)
                System.out.println(get(i).getLevel() + " " + get(i).getValue());
        }else{
            System.out.println("Datas of diagram is empty");
        }
    }

    public void clear(){
        datas.clear();
    }
}
