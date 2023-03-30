class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()

        # Create the PetitTableau widget
        self.PetitTableau = QtWidgets.QTableWidget(self)
        self.PetitTableau.setColumnCount(3)
        self.PetitTableau.setRowCount(3)
        self.PetitTableau.setItem(0, 0, QtWidgets.QTableWidgetItem("Lorem ipsum dolor sit amet"))
        self.PetitTableau.setItem(0, 1, QtWidgets.QTableWidgetItem("consectetur adipiscing elit"))
        self.PetitTableau.setItem(0, 2, QtWidgets.QTableWidgetItem("sed do eiusmod tempor incididunt ut labore et dolore magna aliqua"))
        # ...

        # Adjust the column widths based on the contents of the table
        header = self.PetitTableau.horizontalHeader()
        header.setStretchLastSection(True)
        self.PetitTableau.resizeColumnsToContents()

        # Resize AnotherWindow to fit the width of the table
        table_width = self.PetitTableau.horizontalHeader().length()
        padding = 20
        width_with_padding = table_width + padding
        another_window.resize(width_with_padding, another_window.height())
