/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;

import java.time.LocalTime; //  get heure
import java.time.format.DateTimeFormatter; //   chnager fomrat heure
import java.util.Date; // Importez la classe Date pour la gestion des dates
import java.time.LocalDate;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author lazarowicz
 */
public class Musculation extends Exercice{
    public String exerciceMuscu;
    protected int poids;
    
    public static final String majMusculation = "MaJMusculation"; //sheet
    public int nombreColonneMusculation = 0;
    
    public String nomF = Excel.nomFichier;
    
    public Musculation(String nomExo, String typeExo, String groupementMuscu, String muscle) {
        super(nomExo, typeExo, groupementMuscu, muscle);
        if(nomExo.equals("init")){
            System.out.println("Initialisation ...");
        }
        
    }
    
    public int getNombreLigne(){
        int nbLigne = nombreColonneMusculation;
        return nbLigne;
    }
    
    public int getNombreLigneMusculation(){
        int ligne = 0;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
        Workbook workbook = new XSSFWorkbook(fileIn);
        Sheet sheet = workbook.getSheet(majMusculation);

        ligne = sheet.getLastRowNum();

        workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        //ligne --;
        nombreColonneMusculation = ligne;
        return ligne;
    }
    
    
    public int updateMusculation(String nomExo, String heure, String date, int serie, int repetition, int poids){
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(majMusculation);   
            nombreColonneMusculation ++;

            Row row = sheet.createRow(nombreColonneMusculation);
            row.createCell(0).setCellValue(nomExo);
            row.createCell(1).setCellValue(date);
            row.createCell(2).setCellValue(heure);
            row.createCell(3).setCellValue(serie);
            row.createCell(4).setCellValue(repetition);
            row.createCell(5).setCellValue(poids);
            
            // Écrivez le classeur dans un fichier
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            // Fermez le classeur pour libérer les ressources
            workbook.close();
            
            return nombreColonneMusculation;
        }
        catch (IOException e) {
            e.printStackTrace();
            return -1; // Indique un échec
        }
    }
    
    
    public String[] lectureDonneesMusculation(int ligne){ 
        String[] muscuData = new String[6];
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(majMusculation);
            
            Row row = sheet.getRow(ligne);
            if(row != null){
                Cell cell0 = row.getCell(0);
                muscuData[0] = (cell0 != null) ? cell0.getStringCellValue() : "";
                Cell cell1 = row.getCell(1);
                muscuData[1] = (cell1 != null) ? cell1.getStringCellValue() : "";
                Cell cell2 = row.getCell(2);
                muscuData[2] = (cell2 != null) ? cell2.getStringCellValue() : "";
                Cell cell3 = row.getCell(3);
                double data3 = (cell3 != null && cell3.getCellType() == CellType.NUMERIC) ? cell3.getNumericCellValue() : 0.0;
                muscuData[3] = Double.toString(data3);
                Cell cell4 = row.getCell(4);
                double data4 = (cell4 != null && cell4.getCellType() == CellType.NUMERIC) ? cell4.getNumericCellValue() : 0.0;
                muscuData[4] = Double.toString(data4);
                Cell cell5 = row.getCell(5);
                double data5 = (cell5 != null && cell5.getCellType() == CellType.NUMERIC) ? cell5.getNumericCellValue() : 0.0;
                muscuData[5] = Double.toString(data5);
                
                //System.out.println("Exercice : " + exerciceData[0] + "\nType : " + exerciceData[1] + "\ngroupement muscu : " + exerciceData[2] + "\nmuscle : " + exerciceData[3]);
            }
            else {
                System.out.println("Ligne inexistante dans la feuille de calcul.");
            }

            workbook.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return muscuData;
    }
    
    public String setHeure(int h, int m, int s){
        if(h<10 & m < 10 & s<10){
            String heure = "0" + h + ":" + "0" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(h<10 & m < 10){
            String heure = "0" + h + ":" + "0" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(h<10 & s < 10){
            String heure = "0" + h + ":" + m + ":" + "0" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(m<10 & s < 10){
            String heure = h + ":" + "0" +  m + ":" + "0" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(h<10){
            String heure = "0" + h + ":" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(m<10){
            String heure = h + ":" + "0" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else if(s<10){
            String heure = h + ":" + "0" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;
        }
        else{
            String heure = h + ":" + m + ":" + s;
            //System.out.println("Heure : " + heure);
            return heure;   
        }
    }
    
    public static String getHeure(){
        LocalTime heureActuelle = LocalTime.now();
        DateTimeFormatter format = DateTimeFormatter.ofPattern("HH:mm:ss");
        String heureFormatee = heureActuelle.format(format);
        return heureFormatee;
    }
    
    
    public String setDate(int j, int m, int a){
        if(j<10 & m<10){
            String date = "0" + j + "/" + "0" + m + "/" + a;
            //System.out.println("Date : " + date);
            return date;  
        }  
        else if(j<10){
            String date = "0" + j + "/" + m + "/" + a;
            //System.out.println("Date : " + date);
            return date; 
        }
        else if(m<10){
            String date = j + "/" + "0" +  m + "/" + a;
            //System.out.println("Date : " + date);
            return date; 
        }
        else 
        {
            String date = j + "/" +  m + "/" + a;
            //System.out.println("Date : " + date);
            return date;
        }
    }
  
    public static String getDate(){
        LocalDate dateActuelle = LocalDate.now();
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("dd/MM/yyyy");
        String dateFormat = dateActuelle.format(formatter);
        return dateFormat;
    }
    
}
