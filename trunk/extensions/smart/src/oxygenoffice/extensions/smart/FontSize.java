/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

package oxygenoffice.extensions.smart;

/**
 *
 * @author tibusz
 */
public class FontSize {
    
    static float _fontSizes[] = { (float)6.0,  (float)7.0, (float)8.0, (float)9.0, (float)10.0,
                                (float)10.5, (float)11.0, (float)12.0, (float)13.0, (float)14.0,
                                (float)15.0, (float)16.0, (float)18.0, (float)20.0, (float)22.0,
                                (float)24.0, (float)26.0, (float)28.0, (float)32.0, (float)36.0,
                                (float)42.0, (float)48.0, (float)54.0, (float)60.0, (float)66.0,
                                (float)72.0, (float)80.0, (float)88.0, (float)96.0};

    public static float _getFontSize(int index){
        return _fontSizes[index];
    }

    public static short _getIndexOfFontSize(float value){
        for(short i=0; i<_fontSizes.length; i++)
            if(_fontSizes[i] == value)
                return i;
        return (short)-1;
    }
}
