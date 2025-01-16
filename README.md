# 📊 Calculateur-de-Moyenne / Average-Calculator

![Python](https://img.shields.io/badge/Python-3.x-blue)
![Excel](https://img.shields.io/badge/Excel-Automation-brightgreen)
![Matplotlib](https://img.shields.io/badge/Matplotlib-Graphs-orange)
![Tkinter](https://img.shields.io/badge/Tkinter-UI-lightblue)
![License](https://img.shields.io/badge/License-MIT-lightgrey)

## 🌟 Description

🇫🇷  
Ce projet est une application graphique qui calcule la moyenne des notes à partir de fichiers Excel. Les coefficients et les notes sont extraits automatiquement, puis un graphique est généré pour visualiser les résultats, avec la possibilité d’ajouter et de visualiser des notes via une interface à onglets.

🇬🇧  
This project is a graphical application that calculates the average grades from Excel files. Coefficients and grades are automatically extracted, and a graph is generated to visualize the results, with the option to add and view grades through a tabbed interface.

---

## 🧭 Fonctionnalités / Features

- 🇫🇷  
  - Extraction automatique des coefficients et notes depuis des fichiers Excel.
  - Génération d'un graphique personnalisé des résultats.
  - Flexibilité pour ajouter des modules ou des semestres supplémentaires via une interface à onglets.

- 🇬🇧  
  - Automatic extraction of coefficients and grades from Excel files.
  - Generation of a customized graph of the results.
  - Flexibility to add additional modules or semesters via a tabbed interface.

---

## 🛠️ Utilisation / How to Use

🇫🇷  
Pour utiliser le calculateur de moyenne :
1. **Modifier le chemin du dossier courant :**  
   Dans le fichier `script_V3_Fonctionnel.py`, indiquez le chemin de votre dossier contenant les fichiers Excel.
2. **Ajouter vos données :**  
   - Remplissez vos notes dans le fichier `moyenne_RT.xlsx`. Respectez la structure existante et indiquez les modules dans la colonne *Intitulé*.
   - Ajoutez les coefficients des semestres dans `semestre.xlsx`.
3. **Exécuter le script :**  
   Lancer le script `script_V3_Fonctionnel.py` pour générer le graphique.

🇬🇧  
To use the average calculator:
1. **Modify the current folder path:**  
   In the `script_V3_Fonctionnel.py` file, specify the folder path where your Excel files are located.
2. **Add your data:**  
   - Enter your grades in the `moyenne_RT.xlsx` file. Maintain the existing structure and specify the modules in the *Intitulé* column.
   - Add the semester coefficients in `semestre.xlsx`.
3. **Run the script:**  
   Execute the `script_V3_Fonctionnel.py` script to generate the graph.

---

## 📚 Fichiers Principaux / Main Files

- **script_V3_Fonctionnel.py** : Le script Python principal qui gère l'extraction et le calcul des moyennes.
- **moyenne_RT.xlsx** : Fichier Excel contenant vos notes et modules.
- **semestre.xlsx** : Fichier Excel contenant les coefficients des semestres.

---

## 🎯 Objectif du Projet / Project Goal

🇫🇷  
L'objectif de ce projet est d'automatiser le calcul des moyennes semestrielles pour des étudiants, tout en fournissant une visualisation graphique des résultats. Ce programme peut être adapté à différents cursus académiques en modifiant les fichiers Excel.

🇬🇧  
The goal of this project is to automate the calculation of semester averages for students, while providing a graphical visualization of the results. This program can be adapted to different academic curricula by adjusting the Excel files.

---

## 🔧 Pré-requis / Prerequisites

- Python 3.x
- Fichiers Excel `moyenne_RT.xlsx` et `semestre.xlsx` dans le bon format.
- Bibliothèques Python nécessaires : `openpyxl`, `matplotlib`, `tkinter`, `os`.

---

## 📄 License

Ce projet est sous licence MIT. / This project is licensed under the MIT License.

---
