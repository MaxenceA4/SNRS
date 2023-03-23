import sys
import openpyxl
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

# Chargement du fichier Excel
wb = openpyxl.load_workbook('D:/Vrai_bureau/Logiciel_test/Inventaire.xlsx')

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

    # Return False if the item is not found
    return False


# Subclass QMainWindow to customize your application's main window
class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("Widgets App")

        self.setGeometry(500, 500, 200, 200)

        layout = QVBoxLayout()

        qbox = QComboBox()
        for crit in criterias:
            qbox.addItem(str(crit))

        for i in range(10):
            layout.addWidget(QLabel("Row {}".format(i)))

        layout.addWidget(qbox)

        selected_criteria = qbox.currentText()
        print(selected_criteria)

        search_button = QPushButton("Search", self)
        search_button.clicked.connect(lambda: self.displayObjectsRegardingCriteria())
        layout.addWidget(search_button)

        widget = QWidget()
        widget.setLayout(layout)

        # Set the central widget of the Window. Widget will expand
        # to take up all the space in the window by default.
        self.setCentralWidget(widget)

    def search(self, criteria):
        # Open the workbook and select the sheet

        workbook = openpyxl.load_workbook("Inventaire.xlsx")
        sheet = workbook["Liste"]

        # Loop through the rows to find the matching items
        results = []
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if str(row[0]).lower() == str(criteria).lower():
                results.append(row)

        return (results)

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
