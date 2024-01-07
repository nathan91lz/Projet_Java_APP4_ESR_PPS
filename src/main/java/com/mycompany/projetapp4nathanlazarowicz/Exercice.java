/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */

package com.mycompany.projetapp4nathanlazarowicz;


import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * La classe Exercice fournit des méthodes pour gérer des exercices dans le fichier Excel.
 * Elle permet d'ajouter, supprimer, lire et rechercher des informations sur les exercices.
 *
 * @author lazarowicz
 */
public class Exercice {
    //  Héritage :
    protected String typeExercice;

    public static int nombreColonneExercices = 0; //
    public static String exercice = "Exercice"; //sheet

    private String nomF = Excel.nomFichier; //appel de variable de la classe Excel 
    private JTable tableExercice = Menu.tableExercice;  //permet d'avoir M : updateTableFromExcel

    /**
     * Constructeur de la classe Exercice.
     *
     * @param nomExo Le nom de l'exercice à ajouter.
     * @param typeExo Le type d'exercice.
     * @param groupementMuscu Le groupement musculaire ciblé par l'exercice.
     * @param muscle Le muscle spécifique ciblé par l'exercice.
     */
    public Exercice(String nomExo, String typeExo, String groupementMuscu, String muscle){     //abstract ?
        if(nomExo.equals("init")){
            //System.out.println("Initialisation ...");
        }
        else{
            if(returnExercice(nomExo) == false){
                ajouterExercice(nomExo, typeExo, groupementMuscu, muscle);
                //System.out.println("Vous avez ajouté un nouveau exercice : " + nomExo);
                System.out.println("Nombre de colonne : " + nombreColonneExercices);
            }
            else{
                System.out.println(nomExo + " existe déja, utiliser removeExercice()");
            }
        }
    }


    /**
     * Ajoute un nouvel exercice dans le fichier.
     *
     * @param nomExo Le nom de l'exercice à ajouter.
     * @param typeExo Le type de l'exercice.
     * @param groupementMuscu Le groupement musculaire ciblé par l'exercice.
     * @param muscle Le muscle spécifique ciblé par l'exercice.
     * @return Le numéro de ligne où l'exercice a été ajouté, ou -1 en cas d'échec.
     */
    public int ajouterExercice(String nomExo, String typeExo, String groupementMuscu, String muscle) {
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(exercice);   
            nombreColonneExercices ++;

            Row row = sheet.createRow(nombreColonneExercices);
            row.createCell(0).setCellValue(nomExo);
            row.createCell(1).setCellValue(typeExo);
            row.createCell(2).setCellValue(groupementMuscu);
            row.createCell(3).setCellValue(muscle);

            // Ecrit dans le fichier
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            // Ferme le fichier
            workbook.close();
            //System.out.println("Vous avez ajouté un nouveau exercice : " + nomExo);
            return nombreColonneExercices;
        }
        catch (IOException e) {
            e.printStackTrace();
            return -1; // Indique un échec
        }
    }

    /**
     * Supprime un exercice spécifié dans le fichier Excel.
     *
     * @param nomExo Le nom de l'exercice à supprimer.
     */
    public void removeExercice(String nomExo){
        int ligne = 1;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            while (ligne <= sheet.getLastRowNum()) {
                Row row = sheet.getRow(ligne);

                if (row != null) { // Vérifiez si la ligne existe

                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String nom = cell.getStringCellValue();
                        if (nomExo.equalsIgnoreCase(nom)) { // comparer deux string sans prendre en compte les majuscules
                            int ligneMax = sheet.getLastRowNum();
                            sheet.removeRow(row);
                            if(ligne != ligneMax){ //ne réorganise pas la derniere ligne
                                sheet.shiftRows(ligne + 1, sheet.getLastRowNum(), -1); //   réorganiser les lignes
                            }
                            nombreColonneExercices--;
                            //System.out.println("Exercice " + nomExo + " est supprimé");
                            //break;
                        }
                    }
                }
                ligne++;
            }

            try(FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }

            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Lit les données d'un exercice depuis le fichier et les retourne sous forme de tableau de chaînes.
     *
     * @param ligne Le numéro de ligne de l'exercice dans la feuille de calcul.
     * @return Un tableau de chaînes contenant les données de l'exercice.
     */
    public String[] lectureDonneesExercice(int ligne){ 
        String[] exerciceData = new String[4];
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            Row row = sheet.getRow(ligne);
            if(row != null){
                Cell cell0 = row.getCell(0);
                exerciceData[0] = (cell0 != null) ? cell0.getStringCellValue() : "";
                Cell cell1 = row.getCell(1);
                exerciceData[1] = (cell1 != null) ? cell1.getStringCellValue() : "";
                Cell cell2 = row.getCell(2);
                exerciceData[2] = (cell2 != null) ? cell2.getStringCellValue() : "";
                Cell cell3 = row.getCell(3);
                exerciceData[3] = (cell3 != null) ? cell3.getStringCellValue() : "";

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
        return exerciceData;
    }

    /**
     * Récupère le nombre total de lignes (exercices) dans la feuille de calcul du fichier.
     *
     * @return Le nombre total de lignes.
     */
    public int getNombreLigneExercice(){
        int ligne = 0;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
        Workbook workbook = new XSSFWorkbook(fileIn);
        Sheet sheet = workbook.getSheet(exercice);

        ligne = sheet.getLastRowNum();

        workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        nombreColonneExercices = ligne;
        return ligne;
    }

    /**
     * Recherche un nom exercice dans la feuille de calcul et renvoie sa ligne.
     *
     * @param recherche Le nom de l'exercice à rechercher.
     * @return Le numéro de ligne de l'exercice trouvé, ou -1 si non trouvé.
     */
    public int rechercheNomExercice(String recherche){
        int i = 1;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            while(i <= nombreColonneExercices){
                Row row = sheet.getRow(i);

                if (row != null) {  // Vérifiez si la ligne existe

                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String nom = cell.getStringCellValue();

                        if (recherche.equalsIgnoreCase(nom)) {
                            //System.out.println("Exercice " + recherche + " trouvé à la ligne : " + i);
                        }
                    }
                }
                i++;
            }

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        return i;
    }

    /**
     * Recherche un exercice par son type dans la feuille de calcul et renvoie la ou les ligne(s).
     *
     * @param typeExo Le type de l'exercice à rechercher.
     * @return Le numéro de ligne de l'exercice trouvé, ou -1 si non trouvé.
     */
    public int rechercheTypeExercice(String typeExo){
        int i = 1;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            while(i <= nombreColonneExercices){
                Row row = sheet.getRow(i);

                if (row != null) {  // Vérifiez si la ligne existe

                    Cell cell = row.getCell(1);
                    if (cell != null) {
                        String nom = cell.getStringCellValue();

                        if (typeExo.equalsIgnoreCase(nom)) { // comparer deux string sans prendre en compte les majuscules
                            //System.out.println("Type exercice " + typeExo + " trouvé à la ligne : " + i);
                            //return i;
                        }
                    }
                }
                i++;
            }

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        return i;
    }

    /**
     * Vérifie si un exercice existe dans la feuille de calcul.
     *
     * @param nomExo Le nom de l'exercice à vérifier.
     * @return true si l'exercice existe, false sinon.
     */
    public boolean returnExercice(String nomExo){
        int i = 1;
        boolean retour = false;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            while(i <= nombreColonneExercices){
                Row row = sheet.getRow(i);

                if (row != null) {  // Vérifiez si la ligne existe

                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String nom = cell.getStringCellValue();

                        if (nomExo.equalsIgnoreCase(nom)) { // comparer deux string sans prendre en compte les majuscules
                            //System.out.println("Type exercice " + nomExo + " existe");
                            retour = true;
                            //return i;
                        }
                    }
                }
                i++;
            }

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        return retour;
    }

    /**
     * Récupère le nom d'un exercice depuis une ligne dans la feuille de calcul Exercice.
     *
     * @param ligne Le numéro de ligne de l'exercice dans la feuille de calcul.
     * @return Le nom de l'exercice, ou null si la ligne est inexistante.
     */
    public String getExercice(int ligne){
        String nomExo = null;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);

            Row row = sheet.getRow(ligne);
            if(row != null){
                Cell cell0 = row.getCell(0);
                nomExo = (cell0 != null) ? cell0.getStringCellValue() : "";

                //System.out.println("Exercice : " + nomExo);
            }
            else {
                System.out.println("Ligne inexistante dans la feuille de calcul.");
            }

            workbook.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return nomExo;
    }

    /**
     * Récupère le numéro de la ligne (exercice) dans la feuille de calcul.
     *
     * @param num Le numéro actuel de la ligne.
     * @return Le nouveau numéro de ligne.
     */
    public int getNewNumeroExercice(int num){
        num = nombreColonneExercices;
        return num;
    }



}
