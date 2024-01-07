/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 *
 * @author lazarowicz
 */
public class Planning {
    // Attributs
    public static String planning = "Planning";     
    public String nomF = Excel.nomFichier;
    
    private static boolean[] jours = new boolean[7]; //  tableau de boolean représenatant j de la semaine
    private int joursPlanning;
    
    private String jour;
    private String nom;

    // Constructeur
    public Planning(String nomJP, String jourP) {
        if("init".equals(jourP)){
        //System.out.println("Initialisation ...");
        }
        else{
            this.jour = jourP;
            this.nom = nomJP;
            setPlanning(nomJP, jourP);
        }

    }

    //Méthodes :
    public void setPlanning(String nomP, String jour){
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(planning);   
            
            int ligneJ = getLigneJour(jour);
                 
            Row row = sheet.getRow(ligneJ+1);
            row.createCell(1).setCellValue(nomP);
            //System.out.println("Jour modifié");
                
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();         
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    
    
    public String getNomSeancePlanning(String jour){
        String nomS = null;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(planning);   
            
            int ligneJ = getLigneJour(jour);
            
            Row row = sheet.getRow(ligneJ+1);
            if (row != null) {
                Cell cell = row.getCell(1);
                if (cell != null) {
                    nomS = cell.getStringCellValue();
                    //System.out.println("Nom de la séance : " + nomS);
                } 
                else {
                    //System.out.println("Cellule vide pour le nom de la séance");
                    }
                } 
            else 
            {
                //System.out.println("Ligne non trouvée pour le jour : " + jour);
            }
        workbook.close(); 
        }      
        catch (IOException e) {
            e.printStackTrace();
        }
        return nomS;
    }
    
    public void removePlanning(String jour){
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(planning);   
            
            int ligneJ = getLigneJour(jour);
                 
            Row row = sheet.getRow(ligneJ+1);
            row.createCell(1).setCellValue("");
            //System.out.println("Jour supprimé");
                
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();         
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    
    public String getNomPlanning(){
        return nom;
    }
    
    public String getJourPlanning(){
        return jour;
    }
    
    public static String getNomJours(int i){
        String[] jours = {"Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"};
        //System.out.println("Jour : " + jours[i]);
        return jours[i]; 
    }
    
    
    public static int getLigneJour(String Jour){
        String jour = Jour.toLowerCase();
        switch(jour){
            case "lundi": return 0;
            case "mardi" : return 1;
            case "mercredi": return 2;
            case "jeudi": return 3;
            case "vendredi": return 4;
            case "samedi": return 5;
            case "dimanche": return 6;
            default:
                //System.out.println(jour + " : jour inconnu");
                return -1;
        }
    }
    

    
}
