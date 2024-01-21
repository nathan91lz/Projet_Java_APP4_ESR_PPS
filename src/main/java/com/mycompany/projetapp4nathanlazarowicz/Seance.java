/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;


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
 * La classe Seance permet de créer, supprimer et manipuler les séances dans un fichier Excel.
 * Elle inclut des fonctionnalités telles que l'ajout d'une séance, la modification des exercices d'une séance,
 * la suppression d'une séance ou d'un exercice, la lecture des données d'une séance, etc.
 * Au préalable, l'utilisateur doit avoir rentré des exercices pour pouvoir créer une séance.
 * 
 * @author lazarowicz
 */
public class Seance {
    //Création
    /**
     * Nom de la feuille du fichier excel
     */
    public static final String seance = "Séance";
    private static int nombreExereciceMax = 0;
    private static int nombreColonneSeance = 0;

    private String nomF = Excel.nomFichier;

    private String type;
    private String nom;  

    /**
     * Constructeur de la classe Seance.
     * @param nomJP Nom de la séance.
     * @param type Type d'exercices de la séance.
     * @param nomExo Nom de l'exercice associé à la séance.
     */
    public Seance(String nomJP, String type, String nomExo){  //    Appel de la classe fille avec les instances de la classe mere
        if(nomJP.equals("init")){
            //System.out.println("Initialisation ...");
        }
         else{
            this.nom = nomJP;
            this.type = type;
            ajouterSeance(nomJP, type, nomExo);
         }
    }  


   //Méthodes : 

    /**
     * Ajoute une nouvelle séance dans le fichier Excel.
     * @param nomS Nom de la séance.
     * @param typeS Type de la séance.
     * @param nomExo Nom de l'exercice associé à la séance.
     * @return Le nombre de colonnes dans la séance.
     */
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

    /**
     * Modifie les exercices associés à une séance.
     * @param nomS Nom de la séance.
     * @param nomExo Nom de l'exercice à ajouter à la séance.
     * @return Le nombre d'exercices maximum dans la séance.
     */
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
                    if(colonneMax < 11){
                        row.createCell(colonneMax).setCellValue(nomExo); 
                        nombreExerciceMax = row.getLastCellNum() - 2;
                        //System.out.println("Ajout d'un exo");
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

    /**
     * Retourne la ligne associée à une séance.
     * @param ligneS Nom de la séance.
     * @return L'indice de la ligne associée à la séance.
     */
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

    /**
     * Retourne le nombre de colonnes associées à une séance.
     * @param ligne Indice de la ligne de la séance.
     * @return Le nombre de colonnes dans la séance.
     */
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
        return colonne;
   }

    /**
     * Retourne le nombre total de lignes dans la feuille de séance.
     * @return Le nombre de lignes dans la feuille de séance.
     */
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
        nombreColonneSeance = ligne;
        return ligne;
    }

    /**
     * Supprime une séance du fichier Excel.
     * Elle permet aussi de remonter les colonnes successives
     * @param nomExo Nom de la séance à supprimer.
     */
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
                            //System.out.println("Seance " + nomExo + " est supprimé");
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
     * Supprime un exercice d'une séance du fichier Excel.
     * (J'ai essayé de décaller les cellules suivantes vers la gauche pour ne pas laisser de cellules vide. A ce jour je n'ai pas réussi.)
     * @param nomS Nom de la séance.
     * @param exo Indice de l'exercice à supprimer.
     */
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
                                //System.out.println("Exercice " + exo + " supprimé de la séance " + nomS);
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

    /**
    * Lit les données de la séance et retourne sous forme de tableau de String.
    * 
    * @param ligne Le numéro de ligne à lire dans la feuille de calcul.
    * @return Un tableau de chaînes contenant les données de l'exercice, ou un tableau vide si la ligne est inexistante.
    */
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

    /**
    * Récupère le nom de la séance, basé sur le numéro de ligne spécifié.
    * 
    * @param ligne Le numéro de ligne à lire dans la feuille de calcul.
    * @return Le nom de la séance si la ligne existe, sinon une chaîne vide.
    */
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

    /**
    * Vérifie si le nom de séance spécifié existe dans le fichier Excel.
    * 
    * @param test Le nom de séance à tester.
    * @return true si le nom de séance existe, false sinon.
    */
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


    /**
    * Ajoute des exercices à une séance en modifiant le fichier.
    * 
    * @param nbExo Le nombre d'exercices à ajouter.
    * @return Le numéro de colonne de la séance après ajout des exercices, ou -1 en cas d'échec.
    */
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

    /**
    * Récupère le numéro de la dernière séance du fichier.
    * 
    * @return Le numéro de la dernière séance.
    */
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

    /**
    * Définit le nombre maximal d'exercices pour une séance.
    * 
    * @param nbExoMax Le nombre maximal d'exercices à définir.
    */
    private void setNombreExerciceMax(int nbExoMax){
       if(nbExoMax > nombreExereciceMax){
           nombreExereciceMax = nbExoMax;
       }
   }

    /**
    * Récupère le nombre maximal d'exercices pour une séance.
    * 
    * @return Le nombre maximal d'exercices.
    */
    public int getNombreExerciceMax(){
       //System.out.println("Nombre exercies max : " + nombreExereciceMax);
       return nombreExereciceMax;
   }

    /**
    * Récupère le nom de la séance actuelle.
    * 
    * @return Le nom de la séance.
    */
    public String getNomSeance(){
       return nom;
    }

}

