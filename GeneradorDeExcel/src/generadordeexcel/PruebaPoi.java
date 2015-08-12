/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package generadordeexcel;

import org.apache.poi.hssf.usermodel.*;
import java.io.FileOutputStream;
import javax.swing.JOptionPane;

/**
 *
 * @author Beto Rojas
 */
public class Pruebapoi{

    public Pruebapoi(){
        //Creamos una instancia de la clase HSSFWorkbook llamada libro
        HSSFWorkbook libro = new HSSFWorkbook();
        
        //Creamos una instancia de la clase HSSFSheet llamada hoja y la creamos
        HSSFSheet hoja = libro.createSheet();
        
        //Creamos una instancia de la clase HSSFRow llamada fila y creamos la fila con el indice 0
        HSSFRow fila = hoja.createRow(0);
        
        //Creamos las celdas
        HSSFCell celdauno = fila.createCell(0);
        HSSFCell celdados = fila.createCell(1);
        HSSFCell celdatres = fila.createCell(2);
        
        //Anotamos texto en las celdas
        String A1 = "";
        String B1 = "";
        String C1 = "";
        
        do{
            A1 = JOptionPane.showInputDialog(null, "Contenido celda A1",JOptionPane.QUESTION_MESSAGE);
            B1 = JOptionPane.showInputDialog(null, "Contenido celda B1",JOptionPane.QUESTION_MESSAGE);
            C1 = JOptionPane.showInputDialog(null, "Contenido celda C1",JOptionPane.QUESTION_MESSAGE);
            
            if(A1 == null || B1 == null || C1 == null){
                JOptionPane.showMessageDialog(null, "Te falto introducir el contenido de alguna(s) celda(s)", "Error",2);
            }
        }while(A1 == null || B1 == null || C1 == null);
        HSSFRichTextString textouno = new HSSFRichTextString(A1);
        HSSFRichTextString textodos = new HSSFRichTextString(B1);
        HSSFRichTextString textotres = new HSSFRichTextString(C1);
        celdauno.setCellValue(textouno);
        celdados.setCellValue(textodos);
        celdatres.setCellValue(textotres);
        
        //Cuando escribamos ficheros en Java, se deben encerrar en un try y catch
        try{
            String ex = "";
            
            do{
                ex = JOptionPane.showInputDialog(null, "Nombre del Excel",JOptionPane.QUESTION_MESSAGE);
                if(ex == null){
                    JOptionPane.showMessageDialog(null, "Te falto introducir el nombre del Excel", "Error",2);
                }
            }while(ex == null);
            FileOutputStream archivo = new FileOutputStream("E:/Desktop/"+ ex +".xls");
            libro.write(archivo);
            archivo.close();
            
            JOptionPane.showMessageDialog(null, "Se ha creado el Excel con exito!!!", "Aviso",1);
        }catch(Exception e){
            JOptionPane.showMessageDialog(null, "No se pudo crear el Excel :(", "Error",2);
        }
    }}