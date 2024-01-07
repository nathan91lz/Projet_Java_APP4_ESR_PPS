/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.projetapp4nathanlazarowicz;

/*
 *    Fichier excel : modifier .XML du projet : ajouter dans </project> :
 *        <dependencies>
 *           <!-- Apache POI pour la manipulation de fichiers Excel -->
 *            <dependency>
 *                <groupId>org.apache.poi</groupId>
 *                <artifactId>poi</artifactId>
 *                <version>5.1.0</version> <!-- Utilisez la dernière version disponible -->
 *            </dependency>
 *            <dependency>
 *                <groupId>org.apache.poi</groupId>
 *                <artifactId>poi-ooxml</artifactId>
 *                <version>5.1.0</version> <!-- Utilisez la dernière version disponible -->
 *            </dependency>
 *        </dependencies>
 */

/**
 * Classe principale du projet APP4 de Nathan Lazarowicz.
 * 
 * Cette classe initialise et teste différentes fonctionnalités liées à la gestion d'exercices, 
 * de musculation, de planning et de séances en utilisant un fichier Excel.
 * Elle intègre la bibliothèque Apache POI pour manipuler des fichiers Excel.
 * L'interface utilisateur est également lancée à partir de cette classe.
 * 
 * Les tests et initialisations des différentes classes comme Excel, Exercice, Musculation,
 * Planning, et Seance sont effectués ici. L'interface graphique est lancée à la fin.
 * 
 * Note : Les sections de tests sont commentées et peuvent être activées pour tester 
 * différentes fonctionnalités de l'application.
 *
 * @author lazarowicz
 */
public class ProjetAPP4NathanLazarowicz {

    public static void main(String[] args) {


        //  tests Excel :
        Excel f = new Excel("baseDeDonnee");    //setup du fichier 
        //f.creerFichierExcel("baseDeDonnee");


        //  tests Exercice : 
        /*
        Exercice curl = new Exercice("Curl", "Musculation", "bras", "biceps");
        curl.ajouterExercice("DC haltères", "Musculation", "pectoraux", "pectoraux");

        Exercice DC = new Exercice("DC", "Musculation", "Pec", "pec");
        //Exercice curlbis = new Exercice("Curl", "Musculation", "bras", "triceps");
        curl.removeExercice("Curl");
        //Exercice exerciceMuscu1 = new Exercice("curl");
        */
        /*
        //  tests Exercice de Musculation :
        Musculation exo1 = new Musculation("t", "dg", "d", "gd");
        exo1.updateMusculation(curl.getExercice(1), exo1.setHeure(11, 5, 8), exo1.setDate(24,03,2001), 6, 15, 40);
        //exo1.updateMusculation(curl.getExercice(2), exo1.getHeure(), exo1.getDate(), 3, 15, 20);
        */  

        //System.out.println("Heure actuelle : " + exo1.getHeure());
        //System.out.println("Date actuelle : " + a.getDate());     

        //  tests Planning :
        //Planning p1 = new Planning("seance mardi", "mardi");
        //p1.getNomSeancePlanning("Mardi");

        //  tests Seance :
        //Seance s1 = new Seance("init", "", "");
        //s1.setExerciceSeance("test", "exo2");
        //s1.setExerciceSeance("Seance", "ajout");
        //Seance s2 = new Seance("test", "test", "test");
        //Seance s3 = new Seance("test1", "test", "test");
        //s1.removeSeance("test1");
        //int i = s1.getLigneSeance("Lundi");
        //int c = s1.getColonneSeance(2);
        //System.out.println(c);
        //s1.removeExerciceSeance("bras", 2);

        //Seance s2 = new Seance("Seance du mardi", "Mardi", 1);
        //s1.getNomPlanning();
        //p1.getNomSeancePlanning("Mardi");


        //Interface :

        new Menu().setVisible(true);



    }
}
