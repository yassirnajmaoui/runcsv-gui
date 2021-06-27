from os import set_blocking
import sys
import gc

from PyQt5.QtCore import(
	Qt,
	QModelIndex,
	QRect,
	QItemSelectionModel
)
from PyQt5.QtGui import (
	QStandardItemModel,
	QColor
)
from PyQt5.QtWidgets import (
	QApplication,
	QMainWindow,
	QDesktopWidget,
	QVBoxLayout,
	QPushButton,
	QWidget,
	QTableWidget,
	QLineEdit,
	QTableWidgetItem,
	QTableWidgetSelectionRange,
	QMenu,
	QMenuBar,
	QToolBar,
	QFileDialog
)

import csv
from runcsv import *

# Priorities:
# TODO ASAP: Save file and Open file
# TODO: Fancy selection using "Shift"
# TODO: Quality of life:
#	- Tab button has to be fixed someday


class MainWindow(QMainWindow):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("A spreadsheet better than Microsoft Incel")
		self.ci=0
		self.cj=0

		# Window geometry
		self.setGeometry(100, 100, 1024, 640)
		qtRectangle = self.frameGeometry()
		avGeo = QDesktopWidget().availableGeometry()
		centerPoint = avGeo.center()
		qtRectangle.moveCenter(centerPoint)
		self.move(qtRectangle.topLeft())
		#self.setGeometry(avGeo)

		# Central widget
		self.centralWidget = QWidget()

		# Menu bar
		self.menuBar = QMenuBar(self)
		self.fileMenu = self.menuBar.addMenu("&File")
		self.openAction = self.fileMenu.addAction("Open")
		self.openAction.triggered.connect(self.openFile)
		self.saveAction = self.fileMenu.addAction("Save")
		self.saveAction.triggered.connect(self.saveFile)
		self.editMenu = self.menuBar.addMenu("&Edit")
		self.helpMenu = self.menuBar.addMenu("&Help")

		# Toolbar
		self.toolbar = QToolBar()
		copy_range_action = self.toolbar.addAction("Copy range selector")
		copy_range_action.triggered.connect(self.copyRange)
		toggle_equals_action = self.toolbar.addAction("Toggle \"=\" to beginning of the cell")
		toggle_equals_action.triggered.connect(self.addEquals)
		add_row_action = self.toolbar.addAction("Add row")
		add_row_action.triggered.connect(self.addRow)
		add_column_action = self.toolbar.addAction("Add column")
		add_column_action.triggered.connect(self.addColumn)

		# Formula editor
		self.formulaEdit = FormulaEdit(self, self)
		self.formulaEdit.textChanged.connect(self.onFormulaChange)
		self.formulaEdit.returnPressed.connect(self.formulaReturnPress)

		# Table
		self.tableWidget = QTableWidget(self)
		self.tableWidget.setRowCount(s.shape[0])
		self.tableWidget.setColumnCount(s.shape[1])
		self.tableWidget.setHorizontalHeaderLabels([str(i) for i in range(s.shape[0])])
		self.tableWidget.setVerticalHeaderLabels([str(i) for i in range(s.shape[1])])
		self.tableWidget.itemSelectionChanged.connect(self.onCellChanged)

		# Run button
		self.runBtn = QPushButton("Run")
		self.runBtn.clicked.connect(self.runSheet)

		# Create a QHBoxLayout instance
		layout = QVBoxLayout()
		# Add widgets to the layout
		layout.addWidget(self.toolbar)
		layout.addWidget(self.formulaEdit)
		layout.addWidget(self.tableWidget)
		layout.addWidget(self.runBtn, 2)
		# Set the layout on the application's window
		self.setMenuBar(self.menuBar)
		self.centralWidget.setLayout(layout)
		self.setCentralWidget(self.centralWidget)
	
	def onCellChanged(self):
		# TODO: There should be a way to press Escape and return back to where we were
		self.affectCell(self.ci, self.cj) # Maybe not the most efficient way to do it
		self.ci = self.tableWidget.currentRow()
		self.cj = self.tableWidget.currentColumn()
		if len(self.tableWidget.selectedItems()) > 1:
			self.formulaEdit.clearFocus()
		else:
			self.formulaEdit.setFocus()
			self.formulaEdit.setText(s[self.ci][self.cj])
	
	def onFormulaChange(self):
		s[self.ci][self.cj] = self.formulaEdit.text()
	
	def down(self):
		if(self.ci!=s.shape[0]-1):
			self.tableWidget.setCurrentCell(self.ci+1,self.cj)

	def up(self):
		if(self.ci!=0):
			self.tableWidget.setCurrentCell(self.ci-1,self.cj)

	def right(self):
		if(self.cj!=s.shape[1]-1):
			self.tableWidget.setCurrentCell(self.ci,self.cj+1)

	def left(self):
		if(self.cj!=0):
			self.tableWidget.setCurrentCell(self.ci,self.cj-1)

	def formulaReturnPress(self, key=None):
		self.down()
	
	def runSheet(self):
		for i in np.arange(s.shape[0]):
			for j in np.arange(s.shape[1]):
				self.affectCell(i,j)
	
	def affectCell(self,i,j):
		try:
			process_cell(i,j)
			newitem = QTableWidgetItem(f[i][j])
			self.tableWidget.setItem(i,j, newitem)
			if(s[i][j] != "" and s[i][j][0] == "="):
				self.tableWidget.item(i,j).setBackground(QColor(100,200,200))
		except ValueError as e:
			print("There was an error in cell {"+str(i)+","+str(j)+"}:")
			print(e)

	def keyPressEvent(self, event):
		key = event.key()
		if(key == Qt.Key_Delete):
			for modelIndex in self.tableWidget.selectedIndexes():
				row = modelIndex.row()
				column = modelIndex.column()
				s[row][column] = ""
				self.affectCell(row,column)
		QWidget.keyPressEvent(self,event)

	def copyRange(self):
		selectionMode = "o"
		# gets the selector string and adds it to clipboard
		selectionString = selectionMode+self.getRange()
		QApplication.clipboard().setText(selectionString)
	
	def getRange(self):
		# Returns a string from the table selection and selection mode
		selectionRanges = self.tableWidget.selectedRanges()
		selectionRange = selectionRanges[0]
		min_i = selectionRange.topRow()
		max_i = selectionRange.bottomRow()+1
		min_j = selectionRange.leftColumn()
		max_j = selectionRange.rightColumn()+1
		selectionString = "["+str(min_i)+":"+str(max_i)+","+str(min_j)+":"+str(max_j)+"]"
		return selectionString

	def addEquals(self):
		for selectionRange in self.tableWidget.selectedRanges():
			min_i = selectionRange.topRow()
			max_i = selectionRange.bottomRow()+1
			min_j = selectionRange.leftColumn()
			max_j = selectionRange.rightColumn()+1
			for current_i in range(min_i,max_i):
				for current_j in range(min_j,max_j):
					if(s[current_i][current_j][0] != "="):
						s[current_i][current_j] = "="+s[current_i][current_j]
					else:
						s[current_i][current_j] = s[current_i][current_j][1:]
					self.affectCell(current_i,current_j)

	def saveFile(self):
		fileDialog = QFileDialog()
		name = fileDialog.getSaveFileName(self, 'Save File',"file.csv")
		if(name[0] is not None and name[0] != ""):
			with open(name[0], 'w') as file:
				writer = csv.writer(file)
				writer.writerows(s)
		else:
			print("No file specified, exiting...")

	def openFile(self):
		fileDialog = QFileDialog()
		name = fileDialog.getOpenFileName(self, 'Open File')
		if(name[0] is not None and name[0] != ""):
			s_l = []
			with open(name[0], newline='') as file:
				reader = csv.reader(file)
				s_l = list(reader)

			s_l_arr = np.array(s_l, dtype=object)
			s.resize(s_l_arr.shape,refcheck=False)
			s[:,:] = np.copy(s_l_arr)
			p.resize(s.shape,refcheck=False)
			o.resize(s.shape,refcheck=False)
			f.resize(s.shape,refcheck=False)
			gc.collect()
			self.refreshTableSize()
			self.runSheet()
		else:
			print("No file specified, exiting...")

	def refreshTableSize(self):
		self.tableWidget.setRowCount(s.shape[0])
		self.tableWidget.setColumnCount(s.shape[1])
		self.tableWidget.setHorizontalHeaderLabels([str(i) for i in range(0,s.shape[1])])
		self.tableWidget.setVerticalHeaderLabels([str(i) for i in range(0,s.shape[0])])
		print(s.shape)
		print(p.shape)
		print(o.shape)
		print(f.shape)
	
	def addRow(self):
		s.resize((s.shape[0]+1, s.shape[1]),refcheck=False)
		p.resize((s.shape[0], s.shape[1]),refcheck=False)
		o.resize((s.shape[0], s.shape[1]),refcheck=False)
		f.resize((s.shape[0], s.shape[1]),refcheck=False)
		s[s.shape[0]-1,:] = ""
		p[s.shape[0]-1,:] = ""
		o[s.shape[0]-1,:] = None
		f[s.shape[0]-1,:] = ""
		self.refreshTableSize()

	def addColumn(self):
		s.resize((s.shape[0], s.shape[1]+1),refcheck=False)
		p.resize(s.shape,refcheck=False)
		o.resize(s.shape,refcheck=False)
		f.resize(s.shape,refcheck=False)
		s[:,-1] = ""
		p[:,-1] = ""
		o[:,-1] = None
		f[:,-1] = ""
		self.refreshTableSize()


# TODO: Add rows and columns
# TODO: Add Tab function?

class FormulaEdit(QLineEdit):
	def __init__(self, editor : MainWindow, parent):
		QLineEdit.__init__(self, parent)
		self.editor = editor
	
	def keyPressEvent(self, event):
		key = event.key()
		if key == Qt.Key_Up:
			self.editor.up()
		elif key == Qt.Key_Down:
			self.editor.down()
		elif key == Qt.Key_Left and self.editor.formulaEdit.cursorPosition() == 0:
			self.editor.left()
		elif key == Qt.Key_Right and self.editor.formulaEdit.cursorPosition() == len(self.editor.formulaEdit.text()):
			self.editor.right()
		QLineEdit.keyPressEvent(self, event)


if __name__ == "__main__":
	app = QApplication(sys.argv)
	window = MainWindow()
	window.show()
	sys.exit(app.exec_())

