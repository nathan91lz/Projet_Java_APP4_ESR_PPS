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
    
    public int getNewNumeroExercice(int num){
        num = nombreColonneExercices;
        return num;
    }
    

    
}
