import tkinter as tk, openpyxl, matplotlib.pyplot as plt, os, numpy as np
from tkinter.messagebox import *
from tkinter import ttk
from PIL import Image, ImageTk

#init de la fenetre principale 
root = tk.Tk()
root.title('Calculateur de Moyenne')
notebook=ttk.Notebook(root)
notebook.pack(pady=10, expand=True)

#init des onglets
frm1= ttk.Frame(notebook, width=400, height=280)
frm2= ttk.Frame(notebook, width=400, height=280)
frm3= ttk.Frame(notebook, width=400, height=280)
frm4= ttk.Frame(notebook, width=400, height=280)

#paramétrage des onglet : frm3 -> grille
frm1.pack(fill='both', expand=True)
frm2.pack(fill='both', expand=True)
frm3.grid_columnconfigure(0, weight=1)
frm3.grid_columnconfigure(1, weight=1)
frm4.pack(fill='both', expand=True)

#déplacement vers le dossier Prog_Moyenne
os.chdir('Prog_Moyenne')
d={}
def moyenne(): 
    #Ouverture des 2 fichiers xlsx
    wb_file=openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
    wb_note=openpyxl.load_workbook('moyenne_RT.xlsx', data_only=True, read_only=True)

    #Init de la liste finale
    UE_S=[] 

    # Recherche des UEs et calcul des moyennes pondérées
    for X, n in enumerate(wb_file.sheetnames):
        sheet_file=wb_file[n]
        for row in sheet_file.iter_rows(max_row=2):# Recherche uniquement dans les deux premières lignes
            for cell in row:
                if cell.value and cell.data_type=='s' and len(cell.value)>2 and cell.value[0]+cell.value[1] == 'UE':
                    UE=cell.value
                    col=cell.column
                    UEden, UEnom=0, 0 # Initialisation des dénominateurs et numérateurs pour le calcul de la moyenne
                    
                    for i in range(3, sheet_file.max_row+1): # Boucle sur les lignes suivantes pour calculer les coefficients et moyennes
                        cell= sheet_file.cell(row=i, column=col).value
                        Intitule=str(sheet_file.cell(row=i,column=1).value)
                        if cell !=None and Intitule!='':
                            d[Intitule]=float(cell) # Stockage du coefficient dans un dictionnaire
                            sheet_notes=wb_note.active

                            #Boucle dans le fichier de note pour récupérer la note correspondante
                            for x2 in range(1, sheet_notes.max_column+1):
                                for y2 in range(1, sheet_notes.max_row+1):
                                    notes= sheet_notes.cell(row=y2, column=x2)
                                    if Intitule==str(sheet_notes.cell(row=y2, column=1).value) and notes.data_type != 's' and notes.value!=None:
                                        UEnom+=float(notes.value)*float(cell)
                                        UEden+=float(cell)

                    # Si le dénominateur n'est pas nul, on calcule la moyenne pour l'UE
                    if UEden!=0:
                        # Incrémentation du N° du semestre, du nom de l'UE et de la moyenne de l'UE à notre liste finale
                        UE_S+=[( X, UE, UEnom/UEden)]
    
    #Fermeture des fichiers excel
    wb_file.close()
    wb_note.close()

    # Génération des graphiques
    return init_graph(UE_S)

# Initialisation du graph et du canvas
def init_graph(val):
    global canvas
    # Si canvas existant : on le supprime pour en générer un nouveau
    if canvas != None:
        canvas.destroy()
        
    plt.figure()

    # On appel la 2e partie de la création du graph
    return Diag(val)

# Valeurs -> UE_S et TF : Booleen
def Diag(valeurs, TF=False):
    # Gestion du graph pour 1 semestre
    if len(valeurs)==3 or len(valeurs)==5: # Si longueur de valeurs = 1an (en UE : 3 ou 5 UEs par an)
        plt.subplot(2,2,valeurs[0][0]) 
        plt.title("Semestre 1")
        plt.ylim(0,20)
        plt.axhline(10, color="Red", linestyle='--')
        plt.axhline(8, color="Red", linestyle='--')
        for i in range (len(valeurs)):
            #On affiche le diagramme du semestre
            if float(valeurs[i][2]) >=10 : #Vert
                plt.bar(valeurs[i][1], float((valeurs[i][2])), color="Green")
                pass
            elif float(valeurs[i][2]) >=8: #Orange
                plt.bar(valeurs[i][1], float((valeurs[i][2])), color="Orange")
                pass
            else: #Rouge
                plt.bar(valeurs[i][1], float((valeurs[i][2])), color="Red")
                pass
        #Si TF est False et n'a pas été modifié alors le graphique ne comporte qu'un semestre et est renvoyé à la 3e partie de la création du graph
        if TF==False: return savegraph()    

    else: # n égal au nombre d'année en fonction du nombre d'UE dans valeurs 
        if len(valeurs)==6:
            n=1 # 6 UEs correspond à 1an
        elif len(valeurs)==16:
            n=2 # 16 UEs correspond à 2ans
        elif len(valeurs)==26:
            n=3 # 26 UEs correspond à 3ans
        pass

    # Gestion du graph pour n année(s)
    if valeurs!=[] and (len(valeurs)==6 or len(valeurs)==16 or len(valeurs)==26): 
        # Boucle sur le nombre d'année
        for i in range(n):
            #print(f'i : {i}')
            plt.subplot(2,2,i+1)
            plt.title(f"Année {i+1}")
            plt.ylim(0,20)
            plt.axhline(10, color="Red", linestyle='--')
            plt.axhline(8, color="Red", linestyle='--')

            # init des variables temporaires
            temp=valeurs[0][0] #Semestre Initial
            temp2=temp+1 #Semestre Initial +1
            y=0 # Semestre

            # Tant que y est plus petit que la longueur de valeurs et que le semestre actuel correspond au semestre initial ou semestre initial +1
            while y < len(valeurs) and (valeurs[y][0] == temp or valeurs[y][0] == temp2):
                if float(valeurs[y][2]) >=10 : #Vert
                    plt.bar(valeurs[y][1], float((valeurs[y][2])), color="Green") 
                    pass
                elif float(valeurs[y][2]) >=8: #Orange
                    plt.bar(valeurs[y][1], float((valeurs[y][2])), color="Orange")
                    pass
                else: #Rouge
                    plt.bar(valeurs[y][1], float((valeurs[y][2])), color="Red")
                    pass    
                y+=1 # Incrémentation au prochain semestre
            
            # On retire les valeurs parcourues
            valeurs = valeurs[y:]

            # Si la liste est vide et que c'est la première fois que Diag est lancé, alors on passe à l'étape 3
            if valeurs==[] and TF==False: return savegraph()
        pass

    # Si c'est la première fois qu'on rentre dans la fonction de Diagramme (en rapport à la récursion), alors :
    elif TF==False:

        # initialisation des variables temporaire
        temp=valeurs[-1][0] # dernier semestre
        y=0

        # On cherche à partir de quel y on arrive au dernier semestre
        if y<=(len(valeurs)-1):
            while valeurs[y][0]!=temp and valeurs:
                if y<=(len(valeurs)-1):
                    y+=1 #On ajoute un semestre
                    pass

            # Sem prend le dernier semestre et valeurs2 prend les n année(s) de valeurs à l'aide de y
            Sem=valeurs[y:]
            valeurs2=valeurs[:y]

            # Récursion vers les n année(s) et le dernier semestre puis savegraph
            if valeurs2!=[]:
                Diag(valeurs2, True)
            if Sem!=[]:
                Diag(Sem, True)
                return savegraph()
            pass
        pass

# 3e partie de la création du graph
def savegraph():
    global canvas

    # save du graph
    plt.savefig('graph.png')
    plt.close()

    # gestion du canvas
    image = Image.open("graph.png") 
    photo = ImageTk.PhotoImage(image) 
    canvas = tk.Canvas(frm4, width = image.size[0], height = image.size[1]) 
    canvas.create_image(0,0, anchor = tk.NW, image=photo)
    canvas.image= photo # on sauvegarde la référence de l'image
    canvas.pack() #ajout du canvas à l'onglet 4
    print('canvas reload')
    return canvas

# init du canvas à None pour l'utiliser en variable global dans la création du graphique
canvas=None

note={}
def ajout_note():
    global row_frm3 #row_frm3 -> Nombre de ligne à ajouter dans le excel
    if askyesno('Titre', 'Êtes-vous sûr de vouloir appliquer ces modifications ?'):
        #Ajout de chaque moyenne par module dans le dictionaire note
        for y in range(0,row_frm3):
            moyenne=globals()[f'moyenne{y}'].get("1.0", "end-1c")
            module=globals()[f'module{y}'].get("1.0", "end-1c")
            note[module]=moyenne
        
        showwarning('Titre2', 'Notes ajoutées avec succès.')

        #Réinitialisation de l'onglet
        clear_frame(frm3)
        init_frm3()

    #renvoie vers une fonction ajout des notes dans le excel
    return push(note)

def push(note):
    #Ouverture du fichier de moyenne 
    wb_note=openpyxl.load_workbook('moyenne_RT.xlsx')
    ws=wb_note.active
    modules=[] #modules déjà sauvegardés
    note_s={} #notes déjà sauvegardées

    # boucle sur chaque valeurs à ajouter au fichier de moyenne
    for x in note:
        for y in range(ws.max_row): # on parcours notre fichier pour récupérer nos informations déjà enregistrées dans 'moyenne_RT.xlsx'
            modules.append(ws.cell(row=y+1, column=1).value)
            note_s[ws.cell(row=y+1, column=1).value]=ws.cell(row=y+1, column=2).value
        if x not in modules: # Si notre module ne figure pas dans nos modules déjà sauvegardés, alors :
            ws.cell(row=ws.max_row+1 , column=1).value=x #ajout du module dans le fichier
            ws.cell(row=ws.max_row, column=2).value=int(note[x]) #ajout de la note dans le fichier
        elif str(note_s[x])!='': # Si déjà dans le fichier mais pas de note enregistré alors :
            ws.cell(row=list(note_s.keys()).index(x)+1, column=2).value=int(note[x]) #ajout de la note dans le fichier
    wb_note.save('moyenne_RT.xlsx') # On sauvegarde
    wb_note.close() # On ferme

def clear_frame(frame):
    # Détruit tous les widgets enfants du frame
    for widget in frame.winfo_children():
        widget.destroy()

def newline(): #Incrémentation d'une nouvelle ligne d'ajout dans l'onglet 3
    global row_frm3
    i=row_frm3+1
    tk.Button(frm3, text='Newline', command=newline).grid(column=2, row=i)

    # On créer une variable globale pour chaque module et moyenne de chaque ligne
    globals()[f'module{row_frm3}'] =tk.Text(frm3, bg='cornsilk', width=15, height=2)
    globals()[f'module{row_frm3}'].grid(column=0, row=i)
    globals()[f'moyenne{row_frm3}'] =tk.Text(frm3, bg='cornsilk', width=15, height=2)
    globals()[f'moyenne{row_frm3}'].grid(column=1, row=i)

    # On lance une vérification de la variable row_frm3
    return verif(i)

def verif(i=1):
    #On ajoute une ligne à row_frm3 pour le prochain appel de newline
    if i!=1:
        global row_frm3
        row_frm3+=1
        return row_frm3
    else: #init de row_frm3
        return 1
    

#init de row_frm3 à 1
row_frm3=verif()

# Affichage de l'onglet 1
def affichage_moyenne():
    clear_frame(frm1) #On supprime tout les widget de l'onglet 1
    wb_note=openpyxl.load_workbook('moyenne_RT.xlsx', data_only=True, read_only=True)
    ws=wb_note.active #On s'occupe uniquement de la feuille 1 du fichier 'moyenne_RT.xlsx'

    # On boucle dans cette feuille et on ajoute chaque information du fichier dans l'onglet 1
    for i, row in enumerate(ws.iter_rows()):
        x=0
        for cell in row:
            x+=1
            if cell.value!=None:
                tk.Label(frm1, text=str(cell.value)).grid(column=x, row=i+1)
    wb_note.close() # On ferme le fichier

    tk.Button(frm1, text='reload', command=affichage_moyenne).grid(column=100, row=100) #Ajout du bouton relançant la fonction pour actualiser en fonction des modifications du fichier 'moyenne_RT.xlsx'

# Affichage de l'onglet 2
def affichage_coef():
    clear_frame(frm2) #On supprime tout les widget de l'onglet 2
    wb_file=openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
    #On boucle dans le fichier 'semestre.xlsx' feuille après feuille
    for m, n in enumerate(wb_file.sheetnames): #m commence a 0 et prend donc le numéro du semestre -1
        sheet_file=wb_file[n]
        for i, row in enumerate(sheet_file.iter_rows()): #Parcours de chaque lignes du tableau
            x=0 # initialisation de x : N° de cellule
            for cell in row: #On parcours chaque cellule de chaque ligne
                x+=1 # N° de la cellule
                if cell.value!=None: #On vérifie bien que tout soit
                    if i!=0: #Si != de la ligne 1:
                        tk.Label(frm2, text=str(cell.value)).grid(column=x+m*10, row=i+1) #ajout du texte colone après colone
                    else: #Sinon ajout du texte au dessus des 3 ou 5 UE en fonction du semestre
                        if m<=1:
                            tk.Label(frm2, text=str(cell.value)).grid(column=x+m*10, row=i+1, columnspan=3)
                        else:
                            tk.Label(frm2, text=str(cell.value)).grid(column=x+m*10, row=i+1, columnspan=5)
    wb_file.close() #On ferme le fichier
                        
    tk.Button(frm2, text='reload', command=affichage_coef).grid(column=100, row=100) #Ajout du bouton relançant la fonction pour actualiser en fonction des modifications du fichier 'semestre.xlsx'

# Noms des onglets
notebook.add(frm1, text='Moyennes')
notebook.add(frm2, text='Coefficient')
notebook.add(frm3, text='Ajout de note')
notebook.add(frm4, text='UE')

# fonction initalisation de l'onglet 3 
def init_frm3():
    global module0, moyenne0
    tk.Label(frm3, text='Module').grid(column=0, row=0)
    tk.Label(frm3, text='Moyenne').grid(column=1, row=0)
    module0=tk.Text(frm3, bg='cornsilk', width=15, height=2)
    module0.grid(column=0, row=1)
    moyenne0=tk.Text(frm3, bg='cornsilk', width=15, height=2)
    moyenne0.grid(column=1, row=1)
    tk.Button(frm3, text='Newline', command=newline).grid(column=2, row=1)
    tk.Button(frm3, text='Quit', command=root.quit).grid(column=2, row=32)
    tk.Button(frm3, text='Submit change', command=ajout_note).grid(column=1, row=32)

tk.Button(frm4, text='Générer le graphique', command=moyenne).pack()

#Premier affichage des onglets lors de l'ouverture du script
init_frm3()
affichage_coef()
affichage_moyenne()

#On boucle notre code sur la fenetre
root.mainloop()