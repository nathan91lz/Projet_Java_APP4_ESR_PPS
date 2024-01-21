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
 * Cette class permet d'ajouter, de modifier ou de supprimer les séances dans un planning comprennant les jours de la semaine. 
 * Cela permet de connaître les disponibilités et une vue d'enssemble de l'utilisateur.
 * Lors de son utilisation, l'utilisateur doit avoir déjà créé une séance ainsi que des exercices associés
 * <ul>
 * <li>Le nom du fichier Excel associé au planning.</li>
 * <li>Le nombre de jours dans le planning.</li>
 * <li>Le jour de la séance.</li>
 * <li>Le nom de la séance.</li>
 * </ul>
 * @author lazarowicz
 */
public class Planning {
    // Attributs
    /**
     * Nom de la feuille du fichier excel
     */
    public static String planning = "Planning";     
    private String nomF = Excel.nomFichier;

    private int joursPlanning;
    private String jour;
    private String nom;

    /**
     * Constructeur de la classe Planning.
     * @param nomJP Nom de la séance.
     * @param jourP Jour de la séance.
     */
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

    /**
     * Définit le planning en ajoutant, modifiant ou supprimant une séance.
     * @param nomP Nom de la séance.
     * @param jour Jour de la séance.
     */
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


    /**
     * Récupère le nom de la séance pour un jour donné.
     * @param jour Jour de la séance.
     * @return Le nom de la séance.
     */
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

    /**
     * Supprime une séance du planning pour un jour donné.
     * @param jour Jour de la séance à supprimer.
     */
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


    /**
     * Récupère le nom de la séance.
     * @return Le nom de la séance.
     */
    public String getNomPlanning(){
        return nom;
    }

    /**
     * Récupère le jour de la séance.
     * @return Le jour de la séance.
     */
    public String getJourPlanning(){
        return jour;
    }

    /**
     * Récupère le nom d'un jour à partir de son indice.
     * @param i Indice du jour.
     * @return Le nom du jour.
     */
    public static String getNomJours(int i){
        String[] jours = {"Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"};
        //System.out.println("Jour : " + jours[i]);
        return jours[i]; 
    }

    /**
     * Récupère l'indice d'un jour à partir de son nom.
     * @param Jour Nom du jour.
     * @return L'indice du jour.
     */
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
