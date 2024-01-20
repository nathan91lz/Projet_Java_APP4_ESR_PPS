/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.projetapp4nathanlazarowicz;

//import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * La classe Excel gère la création et l'ouverture de fichiers Excel pour la gestion des séances d'exercice, 
 * y compris les séances de musculation, le planning, et les exercices.
 * Elle sert de base de donnée.
 *
 * @author lazarowicz
 */
public class Excel {
        public static String nomFichier;    //static permet de récup la variable dans autre class
        private final String majMusculation = Musculation.majMusculation; //sheet
        private String planning = Planning.planning;
        private String exercice = Exercice.exercice;
        private String seance = Seance.seance;


    /**
     * Constructeur pour la classe Excel.
     * Ouvre un fichier Excel existant ou crée un nouveau fichier si nécessaire.
     *
     * @param nomF Le nom du fichier Excel à ouvrir ou à créer.
     */
    public Excel(String nomF){
        ouvrirExcel(nomF);
    }

    /**
     * Crée un nouveau fichier Excel prédéfinit avec des feuilles pour le planning, les séances, 
     * les mises à jour de musculation, et les exercices.
     * Chaque feuille est initialisée avec des entêtes de colonne appropriées.
     *
     * @param nomF Le nom du nouveau fichier Excel à créer.
     */
    public void creerFichierExcel(String nomF){
        this.nomFichier = nomF;
        try {
            // Crée un classeur Excel (XLSX)
            Workbook workbook = new XSSFWorkbook();

            //  Feuille Planning : comporte nom des sénaces / jours
            Sheet sheetPlanning = workbook.createSheet(planning);
            Row rowJours = sheetPlanning.createRow(0);
            rowJours.createCell(0).setCellValue("Jours");
            rowJours.createCell(1).setCellValue("Séances");
            Row rowLundi = sheetPlanning.createRow(1);
            rowLundi.createCell(0).setCellValue("Lundi");
            Row rowMardi = sheetPlanning.createRow(2);
            rowMardi.createCell(0).setCellValue("Mardi");
            Row rowMercredi = sheetPlanning.createRow(3);
            rowMercredi.createCell(0).setCellValue("Mercredi");
            Row rowJeudi = sheetPlanning.createRow(4);
            rowJeudi.createCell(0).setCellValue("Jeudi");
            Row rowVendredi = sheetPlanning.createRow(5);
            rowVendredi.createCell(0).setCellValue("Vendredi");
            Row rowSamedi = sheetPlanning.createRow(6);
            rowSamedi.createCell(0).setCellValue("Samedi");
            Row rowDimanche = sheetPlanning.createRow(7);
            rowDimanche.createCell(0).setCellValue("Dimanche");

            //  Feuille Seances 
            Sheet sheetSeance = workbook.createSheet(seance);
            Row rowS = sheetSeance.createRow(0);
            rowS.createCell(0).setCellValue("Séances");
            rowS.createCell(1).setCellValue("Type de sénaces");
            rowS.createCell(2).setCellValue("Exercice 1");
            rowS.createCell(3).setCellValue("Exercice 2");
            rowS.createCell(4).setCellValue("Exercice 3");
            rowS.createCell(5).setCellValue("Exercice 4");
            rowS.createCell(6).setCellValue("Exercice 5");
            rowS.createCell(7).setCellValue("Exercice 6");
            rowS.createCell(8).setCellValue("Exercice 7");
            rowS.createCell(9).setCellValue("Exercice 8");
            rowS.createCell(10).setCellValue("Exercice 9");
            rowS.createCell(11).setCellValue("Exercice 10");
            //  ->>> comment faire N sénace ?????

            //  Feuille MaJMusculation 
            Sheet sheetMaJMusculation = workbook.createSheet(majMusculation);
            Row rowM = sheetMaJMusculation.createRow(0);
            rowM.createCell(0).setCellValue("Exercice");
            rowM.createCell(1).setCellValue("Date");
            rowM.createCell(2).setCellValue("Heure");
            rowM.createCell(3).setCellValue("Série");
            rowM.createCell(4).setCellValue("Répétition");
            rowM.createCell(5).setCellValue("Poids");

            //  Feuille exercices
            Sheet sheetExercice = workbook.createSheet(exercice);
            Row rowE = sheetExercice.createRow(0);
            rowE.createCell(0).setCellValue("Nom exercice");
            rowE.createCell(1).setCellValue("Type d'exercice");
            rowE.createCell(2).setCellValue("Groupement musculaire");
            rowE.createCell(3).setCellValue("Muscle");

            //  -> ajouter endurance et conditionnement     <-

            // Créer un style de police : caractères gras
            Font boldFont = workbook.createFont();
            boldFont.setBold(true);

            // Créer un style de cellule et appliquer la police
            CellStyle boldStyle = workbook.createCellStyle();
            boldStyle.setFont(boldFont);

            // Application du style :
            for (Sheet sheet : workbook) {
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        cell.setCellStyle(boldStyle);
                    }
                }
            }

            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {    //fichier.xlsx
                workbook.write(fileOut);
            }
            workbook.close();
            System.out.println("Fichier Excel créé avec succès, nom de votre fichier : " + nomF);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Ouvre un fichier Excel existant.
     * Utilisé pour accéder aux données dans un fichier Excel existant.
     *
     * @param nomFichier Le nom du fichier Excel à ouvrir.
     */
    public void ouvrirExcel(String nomFichier){
        this.nomFichier = nomFichier;
        try (FileInputStream fileIn = new FileInputStream(nomFichier + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);            
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    
    public void exercicesDeBase(String nomF){
        this.nomFichier = nomF;
        try (FileInputStream fileIn = new FileInputStream(nomF + ".xlsx")) {
            Workbook workbook = new XSSFWorkbook(fileIn);
            Sheet sheet = workbook.getSheet(exercice);      

            Object[][] exercicesData = {
                    {"Pompes", "Conditionnement", "Pectoraux", "Suppérieur"},
                    {"Squats", "Musculation", "Jambes", "Quadriceps"},
                    {"Curl", "Musculation", "Bras", "Biceps"},
                    {"Tractions", "Conditionnement", "Dos", "Grand dorsale"},
                    {"Dips", "Conditionnement", "Bras", "Tricpes"},
                    {"Tractions", "Conditionnement", "Dos", "Grand dorsale"},
                    {"Rowing barre", "Musculation", "Dos", "Grand dorsale"},
                    {"Soulevé de terre", "Conditionnement", "Dos", "Grand dorsale"},
                    {"Presse", "Musculation", "Jambes", "Quadriceps"},
                    {"Leg curl", "Musculation", "Jambes", "Ischio"},
                    {"Développé militaire", "Musculation", "Epaules", "Latérale"},
                    {"Shrug", "Musculation", "Dos", "Trapèze"}
            };

            int numeroLigne = sheet.getLastRowNum();
            
            for (Object[] exerciceData : exercicesData) {
                Row row = sheet.createRow(numeroLigne++);
                int colNum = 0;
                for (Object field : exerciceData) {
                    row.createCell(colNum++).setCellValue((String) field);
                }
            }
            
            try (FileOutputStream fileOut = new FileOutputStream(nomF + ".xlsx")) {    //fichier.xlsx
                workbook.write(fileOut);
            }
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}
