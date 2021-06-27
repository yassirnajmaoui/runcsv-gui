import numpy as np
import re
import csv
from matplotlib import pyplot as plt

parse_refs_regex = re.compile("(\*)?{(.*?),(.*?)}")
# Initial size 10x10
s = np.ndarray([10,10], dtype=object)
s.fill("")
p = np.ndarray(s.shape, dtype=object)
o = np.ndarray(s.shape, dtype=object)
f = np.ndarray(s.shape, dtype=object)

def check_pos(i,j):
	if(i<0 or j<0 or i>s.shape[0] or j>s.shape[1]):
		return False
	return True

def parse_cell(i, j):
	if(not check_pos(i,j)):
		return
	current_s = s[i][j][1:] # remove the "=" in the string
	while True:
		delta = 0
		matches = list() # List of matches
		for match in re.finditer(parse_refs_regex, current_s):
			matches.append(match)
		if len(matches) == 0:
			break
		for match in matches:
			ref_groups = match.groups()
			referenced_i_str = ref_groups[-2]
			referenced_j_str = ref_groups[-1]
			resulting_val = ""
			resulting_i = 0
			resulting_j = 0
			# This part could be made to be parsed by another engine instead of using "exec"
			resulting_i = eval(referenced_i_str)
			resulting_j = eval(referenced_j_str)
			if(ref_groups[0] == '*'):
				resulting_val = p[resulting_i][resulting_j]
			else:
				resulting_val = str(f[resulting_i][resulting_j])
			matched_span = match.span()
			current_s = current_s[:matched_span[0]+delta] + resulting_val + current_s[matched_span[1]+delta:]
			# Correct position for next replacement (Because you cannot reverse a callable_terator)
			delta = len(resulting_val) - (matched_span[1]-matched_span[0])
	return current_s


def process_cell(i,j):
	item = s[i][j]
	if item != "" and item[0] == '=':
		p[i][j] = parse_cell(i,j)
		o[i][j] = eval(p[i][j]) # Execution of the generated command
		f[i][j] = str(o[i][j])
	else:
		p[i][j] = s[i][j]
		o[i][j] = p[i][j]
		f[i][j] = str(o[i][j])


# END OF RUNCSV CODE (Had to be pasted instead of imported to prevent global variables mistakes in Python)

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


# Priorities:
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
		self.formulaEdit.textEdited.connect(self.onFormulaChange)
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
		global s,p,o,f
		self.affectCell(self.ci, self.cj) # Maybe not the most efficient way to do it
		self.ci = self.tableWidget.currentRow()
		self.cj = self.tableWidget.currentColumn()
		if len(self.tableWidget.selectedItems()) > 1:
			self.formulaEdit.clearFocus()
		else:
			self.formulaEdit.setFocus()
			if(s[self.ci][self.cj]=="" or s[self.ci][self.cj]==0):
				self.formulaEdit.setText("")
			else:
				self.formulaEdit.setText(s[self.ci][self.cj])
			
	def onFormulaChange(self):
		global s,p,o,f
		capturedText = self.formulaEdit.text()
		s[self.ci][self.cj] = capturedText
	
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
		global s,p,o,f
		try:
			process_cell(i,j)
			newitem = QTableWidgetItem(f[i][j])
			self.tableWidget.setItem(i,j, newitem)
			if(s[i][j] != "" and s[i][j][0] == "="):
				self.tableWidget.item(i,j).setBackground(QColor(100,200,100))
		except TypeError:
			print("There was an error in cell {"+str(i)+","+str(j)+"}:")

	def keyPressEvent(self, event):
		global s,p,o,f
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
		global s,p,o,f
		# Returns a string from the table selection and selection mode
		selectionRanges = self.tableWidget.selectedRanges()
		selectionRange = selectionRanges[0]
		min_i = selectionRange.topRow()
		max_i = selectionRange.bottomRow()+1
		min_j = selectionRange.leftColumn()
		max_j = selectionRange.rightColumn()+1
		if(max_i-1==min_i):
			i_range_string = str(min_i)
		else:
			i_range_string = str(min_i)+":"+str(max_i)
		if(max_j-1==min_j):
			j_range_string = str(min_j)
		else:
			j_range_string = str(min_j)+":"+str(max_j)
		selectionString = "["+i_range_string+","+j_range_string+"]"
		return selectionString

	def addEquals(self):
		global s,p,o,f
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
		global s,p,o,f
		fileDialog = QFileDialog()
		name = fileDialog.getSaveFileName(self, 'Save File',"file.csv")
		if(name[0] is not None and name[0] != ""):
			with open(name[0], 'w') as file:
				writer = csv.writer(file)
				writer.writerows(s)
		else:
			print("No file specified, exiting...")

	def openFile(self):
		global s,p,o,f
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
			#gc.collect()
			self.refreshTableSize()
			self.runSheet()
		else:
			print("No file specified, exiting...")

	def refreshTableSize(self):
		self.tableWidget.setRowCount(s.shape[0])
		self.tableWidget.setColumnCount(s.shape[1])
		self.tableWidget.setHorizontalHeaderLabels([str(i) for i in range(0,s.shape[1])])
		self.tableWidget.setVerticalHeaderLabels([str(i) for i in range(0,s.shape[0])])
		#self.runSheet()
	
	def addRow(self):
		global s,p,o,f
		s = np.insert(s, (s.shape[0]), np.array([""]*(s.shape[1]),dtype=object), 0)
		p = np.insert(p, (p.shape[0]), np.array([""]*(p.shape[1]),dtype=object), 0)
		o = np.insert(o, (o.shape[0]), np.array([None]*(o.shape[1]),dtype=object), 0)
		f = np.insert(f, (f.shape[0]), np.array([""]*(f.shape[1]),dtype=object), 0)
		self.refreshTableSize()

	def addColumn(self):
		global s,p,o,f
		s = np.insert(s, (s.shape[1]), np.array([""]*(s.shape[0]),dtype=object), 1)
		p = np.insert(p, (p.shape[1]), np.array([""]*(p.shape[0]),dtype=object), 1)
		o = np.insert(o, (o.shape[1]), np.array([None]*(o.shape[0]),dtype=object), 1)
		f = np.insert(f, (f.shape[1]), np.array([""]*(f.shape[0]),dtype=object), 1)
		self.refreshTableSize()


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

