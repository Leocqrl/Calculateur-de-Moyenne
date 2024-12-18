import os, openpyxl, matplotlib.pyplot as plt

os.chdir('DOSSIER COURANT')

def moyenne():
    wb_file=openpyxl.load_workbook('semestre.xlsx', data_only=True, read_only=True)
    wb_note=openpyxl.load_workbook('notes_RT.xlsx', data_only=True, read_only=True)
    UE_S=[]
    for n in wb_file.sheetnames:
        sheet_file=wb_file[n]
        print()
        for x in range(1, sheet_file.max_column+1):
            for y in range(1, sheet_file.max_row+1):
                cell= sheet_file.cell(row=y, column=x)
                if cell.value != None and cell.data_type=='s':
                    if len(cell.value)>2:
                        C=cell.value[0]+cell.value[1]
                        if C == 'UE':
                            UE=cell.value
                            col=cell.column
                            UEden=0
                            UEnom=0
                            for i in range(3, sheet_file.max_row+1):
                                cell= sheet_file.cell(row=i, column=col)
                                Intitule=str(sheet_file.cell(row=i,column=1))
                                if cell.value !=None and Intitule!='':
                                    sheet_notes=wb_note.active
                                    for x2 in range(1, sheet_notes.max_column+1):
                                        for y2 in range(1, sheet_notes.max_row+1):
                                            notes= sheet_notes.cell(row=y2, column=x2)
                                            if Intitule==str(sheet_notes.cell(row=y2, column=1)) and notes.data_type != 's'  and notes.value!=None:
                                                UEnom+=int(notes.value)*int(cell.value)
                                                UEden+=int(cell.value)
                            if UEden!=0:
                                UE_S+=[( UE, UEnom/UEden)]
    wb_file.close()
    wb_note.close()
    return UE_S
    
print(moyenne())


def Diag(valeurs):
    plt.figure()
    plt.title("Résultats du BUT RT")
    plt.ylim(0,20)
    plt.axhline(10, color="Red")
    plt.axhline(8, color="Red")
    for n in range (len(valeurs)): 
        if int(valeurs[n][1]) >=10 : #Vert
            plt.bar(valeurs[n][0], int(valeurs[n][1]), color="Green")
        elif int(valeurs[n][1]) >=8: #Orange
            plt.bar(valeurs[n][0], int(valeurs[n][1]), color="Orange")
        else: #Rouge
            plt.bar(valeurs[n][0], int(valeurs[n][1]), color="Red")
    plt.show()

Diag(moyenne())
    
