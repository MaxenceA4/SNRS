import sys
import openpyxl
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

# Chargement du fichier Excel
wb = openpyxl.load_workbook('C:/Users/Maxence/Desktop/Logiciel_test/Inventaire.xlsx')

# Accéder à la feuille "Utilisateurs"
ws_users = wb['Utilisateurs']

# Récupérer tous les noms d'utilisateurs dans la colonne A
users = []
for cell in ws_users['A']:
    users.append(cell.value)

# Accéder à la feuille "Critères"
ws_criteria = wb['Critères']

# Récupérer tous les critères dans la colonne A
criteria_set = set()
for cell in ws_criteria['A']:
    criteria_set.add(cell.value)
criterias = list(criteria_set)


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


# Nouvelle fenetre créer quand on clique sur un objet dans la liste
class AnotherWindow(QWidget):
    def __init__(self):
        super().__init__()
        print("Window created")

        self.setWindowTitle("Information")
        self.setGeometry(100, 100, 500, 600)
        self.PetitTableau = QTableWidget()
        self.PetitTableau.setRowCount(15)
        self.PetitTableau.setColumnCount(3)

        # Appel de la fonction qui copie le tableau excel dans le tableau de la fenetre
        self.copyExcelTable()

        layout = QVBoxLayout()
        layout.addWidget(self.PetitTableau)
        self.setLayout(layout)

        self.show()

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
                # self.PetitTableau.setItem(row - 1, col - 1, QTableWidgetItem(sheet.cell(row, col).value))


# Subclass QMainWindow to customize your application's main window
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

    def SearchSerialNumber(self, serial_number):
        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Clear ListView
        self.model.clear()

        # Loop through the rows to find the matching item even partially
        for row in sheet.iter_rows(min_row=2, values_only=True):
            # any serial number starting with the same letters as the search bar
            if str(serial_number).lower() in str(row[3]).lower():
                self.model.appendRow(QStandardItem(str(row[1])))
                #Applying the color function
                self.color(self.model.item(0))
            # if the viewlist is empty, it will display "Aucun résultat"
            elif self.model.rowCount() == 0:
                self.model.appendRow(QStandardItem("Aucun résultat"))
            # if it doesnt include the number, it will delete the row
            elif str(serial_number).lower() not in str(row[3]).lower():
                self.model.removeRow(0)
            # if the search bar is empty, it will display nothing
            elif str(serial_number).lower() == "":
                self.model.clear()

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
        print("Creating new window...")
        self.w = AnotherWindow()
        self.w.show()

    def ClickOnObject(self):
        # Open the workbook and select the sheet
        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Get information in row 4
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if row[1] == self.listView.currentIndex().data():
                # Create a MessageBox
                # QMessageBox.about(self, "Coefficient", "Coefficient : " + str(row[3]))
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
