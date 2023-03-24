import sys
import openpyxl
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

# Chargement du fichier Excel
wb = openpyxl.load_workbook('Inventaire.xlsx')

# Accéder à la feuille "Critères"
ws_criteria = wb['Critères']

# Récupérer tous les critères dans la colonne A
criteria_set = set()
for cell in ws_criteria['A']:
    criteria_set.add(cell.value)
criterias = list(criteria_set)

#Variable en read pour les subwindows
selectedSerialNumber = "Should be overwritten"

def is_available(item):
    # Open the workbook and select the sheet
    workbook = openpyxl.load_workbook("Inventaire.xlsx")
    sheet = workbook["Liste"]

    # Loop through the rows to find the matching item
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[1] == item:
            # Check the availability
            if row[2] == "Dispo":
                return True
            else:
                return False


class AnotherWindow2(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Information")
        self.setGeometry(100, 100, 400, 200)

        self.InputInitiales = QLineEdit()
        self.InputInitiales.setPlaceholderText("Initiales")

        self.InputLocalisation = QLineEdit()
        self.InputLocalisation.setPlaceholderText("Localisation")

        self.DateToday = QDate.currentDate()
        self.DateToday = self.DateToday.toString("dd/MM/yyyy")

        self.InputManip = QLineEdit()
        self.InputManip.setPlaceholderText("Manip / stock")

        self.ConfirmButton = QPushButton("Confirmer")
        self.ConfirmButton.clicked.connect(self.ConfirmButtonFonction)

        layout = QVBoxLayout()
        layout.addWidget(self.InputInitiales)
        layout.addWidget(self.InputLocalisation)
        layout.addWidget(self.InputManip)
        layout.addWidget(QLabel(self.DateToday))
        layout.addWidget(self.ConfirmButton)

        self.setLayout(layout)

        self.show()

    def ConfirmButtonFonction(self):
        # Open the inventory workbook
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]
        print("Workbook opened")


# Nouvelle fenetre créer quand on clique sur un objet dans la liste
class AnotherWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Information")
        self.setGeometry(100, 100, 500, 600)
        self.PetitTableau = QTableWidget()
        self.PetitTableau.setRowCount(15)
        self.PetitTableau.setColumnCount(3)

        # Bouton Test
        self.Test = QPushButton("Test")
        self.Test.clicked.connect(self.TestFonction)

        # Bouton "Modifier"
        self.Modifier = QPushButton("Modifier")
        self.Modifier.clicked.connect(self.ModifierFonction)

        # Appel de la fonction qui copie le tableau excel dans le tableau de la fenetre
        self.copyExcelTable()

        layout = QVBoxLayout()
        layout.addWidget(self.PetitTableau)
        layout.addWidget(self.Modifier)
        layout.addWidget(self.Test)

        self.setLayout(layout)

        self.show()

    def TestFonction(self):
        print("Selected serial number in another class", selectedSerialNumber)


    def ModifierFonction(self):
        # Create a new window
        self.w = AnotherWindow2()
        self.w.show()
        print(selectedSerialNumber)

    # Fonction qui copie un tableau excel dans un tableau de la fenetre
    def copyExcelTable(self):

        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("test.xlsx", data_only=True)
        sheet = workbook["Résumé"]

        # Ajout des titres des colonnes
        self.PetitTableau.setHorizontalHeaderLabels(["Critères", "Valeur"])

        # Copy the Excel table in the new window
        for row in range(1, 16):
            for col in range(1, 4):
                # Copie de la cellule dans le tableau de la fenetre
                self.PetitTableau.setItem(row - 1, col - 1, QTableWidgetItem(str(sheet.cell(row, col).value)))


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Trouve ta gauge !")

        self.setGeometry(100, 100, 800, 600)

        layout = QVBoxLayout()

        # Create a search bar
        qsearch = QLineEdit()
        qsearch.setPlaceholderText("Search")
        qsearch.textChanged.connect(self.SearchSerialNumber)
        layout.addWidget(qsearch)

        qbox = QComboBox()
        for crit in criterias:
            qbox.addItem(str(crit))

        layout.addWidget(qbox)

        self.listView = QListView()
        self.model = QStandardItemModel()
        self.listView.setModel(self.model)
        self.listView.setObjectName("listView-1")
        self.listView.clicked.connect(self.ClickOnObject)

        layout.addWidget(self.listView)

        search_button = QPushButton("Search", self)
        search_button.clicked.connect(lambda: self.displayObjectsRegardingCriteria())
        layout.addWidget(search_button)

        widget = QWidget()
        widget.setLayout(layout)

        # Set the central widget of the Window. Widget will expand
        # to take up all the space in the window by default.
        self.setCentralWidget(widget)

        # Vérification de la version du logiciel
        wb = openpyxl.load_workbook('Inventaire.xlsx')
        sh_ver = wb['Version']
        version = sh_ver['A2'].value
        if version != "Alpha 1.0":
            print("Version du logiciel incompatible")
            QMessageBox.critical(None, "Erreur",
                                 "Version du logiciel incompatible, Merci de télécharger la dernière version")
            sys.exit()

    def SearchSerialNumber(self, serial_number):
        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Clear ListView
        self.model.clear()

        # Loop through the rows to find the matching item even partially
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # if the serial number in search bar is in the row, it will add the item to the listview
            if str(serial_number).lower() in str(row[3]).lower():
                # Add the item to the listview
                self.model.appendRow(QStandardItem(str(row[1])))
                # Applying the color function
                self.color(self.model.item(0))
            else:  # if the serial number is not in the row, it will not add the item to the listview
                pass

    # fonction qui choisi une couleur en fonction de la disponibilité
    def color(self, item):
        # Loop through results to find if they are available and display with different colors
        for i in range(0, self.model.rowCount()):
            item = self.model.item(i)
            if is_available(item.text()):
                item.setForeground(QColor(0, 255, 0))
            else:
                item.setForeground(QColor(255, 0, 0))

    def search(self, criteria):
        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Clear ListView
        self.model.clear()

        # Loop through the rows to find the matching items
        results = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if str(row[0]).lower() == str(criteria).lower():
                self.model.appendRow(QStandardItem(str(row[1])))
                # Application couleur
                self.color(row[1])
        if self.model.rowCount() == 0:
            self.model.appendRow(QStandardItem("Aucun résultat"))
        # Return the results
        return results

    # Create a new window
    def show_new_window(self, checked):
        self.w = AnotherWindow()
        self.w.show()

    def AssociateSerialNumber(self, AssociateSerialNumber):
        # Open the workbook and select the sheet
        global selectedSerialNumber
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Associate the serial number of the clicked item to a variable
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[1] == self.listView.currentIndex().data():
                AssociateSerialNumber = row[3]
                print("globally selected serial number before", selectedSerialNumber)
                selectedSerialNumber = AssociateSerialNumber
                print("globally selected serial number after", selectedSerialNumber)

    def ClickOnObject(self):
        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]
        AssociateSerialNumber = ''

        # Call he function to associate the serial number of the clicked item to a variable
        self.AssociateSerialNumber(AssociateSerialNumber)

        # Open the new window
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[1] == self.listView.currentIndex().data():
                # Create a new window
                self.show_new_window(True)

    def displayObjectsRegardingCriteria(self):

        criteria = ''

        for child in self.window().children()[1].children():
            if type(child) == QComboBox:
                print(child.currentText())
                criteria = child.currentText()

        objects = self.search(criteria)

        u = 0
        for object in objects:

            children = self.window().children()[1].children()
            print(object)
            print(children)
            self.layout().addWidget(QLabel(object[1]))
            for i in range(u, len(children)):
                if type(children[i]) == QLabel:
                    u = i
                    print(children[i].setText(object[1]))

        return ''


app = QApplication(sys.argv)
window = MainWindow()
window.show()

app.exec()
