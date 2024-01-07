/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;

import static com.mycompany.projetapp4nathanlazarowicz.Exercice.exercice;
import static com.mycompany.projetapp4nathanlazarowicz.Exercice.nombreColonneExercices;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Scanner; // récupérer caracteres
import org.apache.poi.ss.usermodel.Cell;

/**
 *
 * @author lazarowicz
 */
public class Seance {
    //Création
    public static String seance = "Séance";
    public static int nombreExereciceMax = 0;
    public static int nombreColonneSeance = 0;
    
    public String nomF = Excel.nomFichier;
    
    protected String type;
    protected String nom;  
    
    public Seance(String nomJP, String type, String nomExo){  //    Appel de la classe fille avec les instances de la classe mere
        if(nomJP.equals("init")){
            System.out.println("Initialisation ...");
        }
         else{
            this.nom = nomJP;
            this.type = type;
            ajouterSeance(nomJP, type, nomExo);
         }
         
         
         //ajouterExercicesSeance(nbExo);
         /*
         if(getDispoJours(jourP) == false){  //jour dispo
             Planning p1 = new Planning(nomJP, jourP);
             ajouterSeance(nomJP, jourP);
             ajouterExercicesSeance(nbExo);
         }
         else{
             ajouterSeance(nomJP, jourP);
             ajouterExercicesSeance(nbExo);
         */
    }  

   
   //Méthodes : 
   public int ajouterSeance(String nomS, String typeS, String nomExo){
       try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(seance);   
            nombreColonneSeance ++;

            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.createCell(0).setCellValue(nomS);
            row.createCell(1).setCellValue(typeS);
            row.createCell(2).setCellValue(nomExo);
            
            
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();
            
            return nombreColonneSeance;
        }
        catch (IOException e) {
            e.printStackTrace();
            return -1; // Indique un échec
        }
   }
   
   public int setExerciceSeance(String nomS, String nomExo){
       int nombreExerciceMax = 0;
       try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(seance); 
            String testNomS = null;
            
            for(int i = 1; i <= sheet.getLastRowNum(); i++){
                testNomS = getNomSeance(i);
                if(testNomS.trim().equals(nomS)){
                    Row row = sheet.getRow(i);
                    int colonneMax = row.getLastCellNum();
                    //System.out.println("nombre cell : " + colonneMax);
                    if(colonneMax < 11){
                        row.createCell(colonneMax).setCellValue(nomExo); 
                        nombreExerciceMax = row.getLastCellNum() - 2;
                        System.out.println("Ajout d'un exo");
                        try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) { //ecrit dans le fichier
                            workbook.write(fileOut);
                        }
                    }
                }
            }
        workbook.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
       return nombreExerciceMax;
   }
   
   public int getLigneSeance(String ligneS){
        int ligneM = 0;
        int ligne = 0;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);

            ligneM = sheet.getLastRowNum();
            
            for(int i = 0; i <= ligneM; i++){
                if(getNomSeance(i).equals(ligneS)){
                    ligne = i;
                }
            }
            
            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        return ligne;
   }
   
   public int getColonneSeance(int ligne){
       int colonne = 0;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);

            Row row = sheet.getRow(ligne);
            colonne = row.getLastCellNum();

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        //ligne --;
        //nombreColonneSeance = colonne;
        return colonne;
   }
   
   public int getNombreLigneSeance(){
        int ligne = 0;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);

            ligne = sheet.getLastRowNum();

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
        //ligne --;
        nombreColonneSeance = ligne;
        return ligne;
    }
   
   public void removeSeance(String nomExo){
        int ligne = 1;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);

            while (ligne <= sheet.getLastRowNum()) {
                Row row = sheet.getRow(ligne);
                if (row != null) { // Vérifiez si la ligne existe
                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String nom = cell.getStringCellValue();
                        if (nomExo.equalsIgnoreCase(nom)) { // comparer deux string sans prendre en compte les majuscules
                            int ligneMax = sheet.getLastRowNum();
                            sheet.removeRow(row);
                            if(ligne != ligneMax){
                                sheet.shiftRows(ligne + 1, sheet.getLastRowNum(), -1); //   réorganiser les lignes
                            }
                            nombreColonneSeance--;
                            System.out.println("Seance " + nomExo + " est supprimé");
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
   
   public void removeExerciceSeance(String nomS, int exo){
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);

            int ligne = 1;
            while (ligne <= sheet.getLastRowNum()) {
                Row row = sheet.getRow(ligne);
                if (row != null) {
                    Cell cell = row.getCell(0);
                    if (cell != null) {
                        String nomSeance = cell.getStringCellValue();
                        if (nomS.equalsIgnoreCase(nomSeance)) {
                            int colonne = exo + 1;
                            Cell cellExo = row.getCell(colonne);
                            if (cellExo != null) {
                                int cMax = row.getLastCellNum();
                                /*
                                for (int c = colonne + 1; c <= row.getLastCellNum(); c++) { //Décallage de cellules
                                    Cell nextCell = row.getCell(c);
                                    if (nextCell != null) {
                                        row.getCell(c - 1).setCellValue(nextCell.getStringCellValue());
                                    }
                                }
                                */
                                row.removeCell(cellExo);
                                //row.shiftCellsLeft(colonne + 1, cMax, 1); //   réorganiser les lignes
                                System.out.println("Exercice " + exo + " supprimé de la séance " + nomS);
                                //break; // Vous pouvez ajouter une logique pour gérer la suppression de plusieurs occurrences
                            }
                        }
                    }
                }
                ligne++;
            }
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }

            workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
   }
   
   public String[] lectureDonneesSeance(int ligne){ 
        String[] exerciceData = new String[12];
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);
            
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
                Cell cell4 = row.getCell(4);
                exerciceData[4] = (cell4 != null) ? cell4.getStringCellValue() : "";
                Cell cell5 = row.getCell(5);
                exerciceData[5] = (cell5 != null) ? cell5.getStringCellValue() : "";
                Cell cell6 = row.getCell(6);
                exerciceData[6] = (cell6 != null) ? cell6.getStringCellValue() : "";
                Cell cell7 = row.getCell(7);
                exerciceData[7] = (cell7 != null) ? cell7.getStringCellValue() : "";
                Cell cell8 = row.getCell(8);
                exerciceData[8] = (cell8 != null) ? cell8.getStringCellValue() : "";
                Cell cell9 = row.getCell(9);
                exerciceData[9] = (cell9 != null) ? cell9.getStringCellValue() : "";
                Cell cell10 = row.getCell(10);
                exerciceData[10] = (cell10 != null) ? cell10.getStringCellValue() : "";
                Cell cell11 = row.getCell(11);
                exerciceData[11] = (cell11 != null) ? cell11.getStringCellValue() : "";
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
   
   public String getNomSeance(int ligne){ 
        String exerciceData = null;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);
            
            Row row = sheet.getRow(ligne);
            if(row != null){
                Cell cell0 = row.getCell(0);
                exerciceData = (cell0 != null) ? cell0.getStringCellValue() : "";  
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
   
   public boolean testNomSeance(String test){
       boolean nomS = false;
       String testNomS;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(seance);
            
            int sMax = sheet.getLastRowNum();
            
            for(int i = 1; i <= sMax; i++){
                Row row = sheet.getRow(i);
                if(row != null){
                    Cell cell0 = row.getCell(0);
                    testNomS = (cell0 != null) ? cell0.getStringCellValue() : "";  
                    if(testNomS.equals(test)){
                        nomS = true;
                    }
                    else{
                        nomS = false;
                    }
                }
                else {
                    System.out.println("Ligne inexistante dans la feuille de calcul.");
                }
            }
            workbook.close();
        }
        catch (IOException e) {
            e.printStackTrace();
        }
        return nomS;
   }
   
   
   public int ajouterExercicesSeance(int nbExo){
       try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);

            Sheet sheet = workbook.getSheet(seance);   

            Row row = sheet.getRow(nombreColonneSeance);
            
            Scanner scanner = new Scanner(System.in);
            
            for(int i = 1; i <= nbExo; i++){
                System.out.print("Rentrez le nom de l'exercice n°" + i + " : "); // => methode tester exercice s'il existe deja !!!
                String nomExo = scanner.nextLine();
                
                row.createCell(1 + i).setCellValue(nomExo);  //getCell
            }
            scanner.close();
            setNombreExerciceMax(nbExo);
            
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {
                workbook.write(fileOut);
            }
            workbook.close();
            
            return nombreColonneSeance;
        }
        catch (IOException e) {
            e.printStackTrace();
            return -1; // Indique un échec
        }
   }
   
   public int getSeanceMax(){
       int sMax = 0;
       try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
        Workbook workbook = new XSSFWorkbook(fileIn);
        Sheet sheet = workbook.getSheet(seance);

        sMax = sheet.getLastRowNum();
        

        workbook.close();
        } 
        catch (IOException e) {
            e.printStackTrace();
        }
       nombreColonneSeance = sMax;
       return sMax;
   }
   
   private void setNombreExerciceMax(int nbExoMax){
       if(nbExoMax > nombreExereciceMax){
           nombreExereciceMax = nbExoMax;
       }
   }
   
   public int getNombreExerciceMax(){
       System.out.println("Nombre exercies max : " + nombreExereciceMax);
       return nombreExereciceMax;
   }
    
   public String getNomSeance(){
       return nom;
   }
    
}

