/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;

//import java.util.Date;

//import static com.mycompany.projetapp4nathanlazarowicz.Exercice.exercice;
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
        System.out.println("Initialisation ...");
        }
        else{
            this.jour = jourP;
            this.nom = nomJP;
            setNomJours(jourP);
            setPlanning(nomJP, jourP);
            //modifierPlanning(nomJP, jourP);
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
            System.out.println("Jour modifié");
                
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();         
        }
        catch (IOException e) {
            e.printStackTrace();
        }
    }
    
    public void modifierPlanning(String nomP, String jour){
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(planning);   
            
            int ligneJ = getLigneJour(jour);
                 
            if(getDispoJours(jour) == true){    // ecrire si jour est dispo 
                Row row = sheet.getRow(ligneJ+1);
                row.createCell(1).setCellValue(nomP);
                System.out.println("Jour modifié");
            }
            else{
                System.out.println("ERREUR, ce jour est deja pris, use : setNomJours");
            }
            
            // Écrivez le classeur dans un fichier
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            // Fermez le classeur pour libérer les ressources
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
                    System.out.println("Nom de la séance : " + nomS);
                } 
                else {
                    System.out.println("Cellule vide pour le nom de la séance");
                    }
                } 
            else 
            {
                System.out.println("Ligne non trouvée pour le jour : " + jour);
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
            System.out.println("Jour supprimé");
                
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
        System.out.println("Jour : " + jours[i]);
        return jours[i]; 
    }
    
    public static boolean getDispoJours(String Jour) {
    int index = -1;
    String jour = Jour.toLowerCase();
    switch(jour) {
        case "lundi": index = 0; break;
        case "mardi": index = 1; break;
        case "mercredi": index = 2; break;
        case "jeudi": index = 3; break;
        case "vendredi": index = 4; break;
        case "samedi": index = 5; break;
        case "dimanche": index = 6; break;
        default:
            System.out.println(jour + " jour inconnu");
            //return false;
    }

    if (index != -1) {
        if (!jours[index]) {
            System.out.println("Jour : " + jour + " non réservé");  //jour dispo
            return false;
        } else {
            System.out.println("Jour : " + jour + " réservé");  //jour non dispo
            return true;
        }
    }

    return false; // cas par défaut
}
    
    public static int getLigneJour(String Jour){
        String jour = Jour.toLowerCase();
        switch(jour){
            case "lundi": 
                //System.out.println("Lundi ==> 0");
                return 0;
            case "mardi" : return 1;
            case "mercredi": return 2;
            case "jeudi": return 3;
            case "vendredi": return 4;
            case "samedi": return 5;
            case "dimanche": return 6;
            default:
                System.out.println(jour + " : jour inconnu");
                return -1;
        }
    }
    
    public static void setNomJours(String Jour){
        String jour = Jour.toLowerCase();
        switch(jour){
            case "lundi" : jours[0] = true; 
                break;
            case "mardi" : jours[1] = true; 
                break;
            case "mercredi" : jours[2] = true; 
                break;
            case "jeudi" : jours[3] = true; 
                break;
            case "vendredi" : jours[4] = true; 
                break;
            case "samedi" : jours[5] = true; 
                break;
            case "dimanche" : jours[6] = true; 
                break;
            default : break;
        }
    }
    
    public static void removeJour(String Jour){
        String jour = Jour.toLowerCase();
        switch(jour){
            case "lundi" : jours[0] = false; 
                break;
            case "mardi" : jours[1] = false; 
                break;
            case "mercredi" : jours[2] = false; 
                break;
            case "jeudi" : jours[3] = false; 
                break;
            case "vendredi" : jours[4] = false; 
                break;
            case "samedi" : jours[5] = false; 
                break;
            case "dimanche" : jours[6] = false; 
                break;
            default : break;
        }
    }

    
}
