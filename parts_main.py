import sys
import partsdialog
import os
import pyodbc
import PyQt5
import datetime
import re
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QGridLayout, QDesktopWidget, QLCDNumber
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSlot, QTimer, QTime

# Main class for initiating the barcoding application.
# I should probably rename this to something more specific as
# multiple barcoding applications will be created. i.e. parts, labor, etc.

class BarcodeApp(QMainWindow, partsdialog.Ui_MainWindow):

	# Init main.
	def __init__(self, parent=None):

		# Start.
		super(BarcodeApp, self).__init__(parent)
		self.setupUi(self)

		# Code for sizing the program and centering it on screen.
		qtRectangle = self.frameGeometry()
		print(qtRectangle)
		centerPoint = QDesktopWidget().availableGeometry().center()
		qtRectangle.moveCenter(centerPoint)
		self.move(qtRectangle.topLeft())

		# What we do after WOID is scanned.
		self.txtWOBOMID.returnPressed.connect(self.wobomid_after_update)
		# What we do after Part is scanned.

		# What we do after clock # is scanned.
		self.txtClockID.returnPressed.connect(self.clockid_after_update)

		# Event for enter btn click.
		self.btnEnter.clicked.connect(self.on_click)

		# Clear values from field. Reset focus to WOID.
		self.btnClear.clicked.connect(self.clearForm)

		# On load, we're setting focus to the clockID
		self.txtClockID.setFocus()

		self.lblQty.setText('')

		# Needed to intially define the message on load to be used later.
		BarcodeApp.message = ''

		# Call clock at start of program.
		self.clock()

		self.btn1.clicked.connect(self.btn1Click)
		self.btn2.clicked.connect(self.btn2Click)
		self.btn3.clicked.connect(self.btn3Click)
		self.btn4.clicked.connect(self.btn4Click)
		self.btn5.clicked.connect(self.btn5Click)
		self.btn6.clicked.connect(self.btn6Click)
		self.btn7.clicked.connect(self.btn7Click)
		self.btn8.clicked.connect(self.btn8Click)
		self.btn9.clicked.connect(self.btn9Click)
		self.btn0.clicked.connect(self.btn0Click)
		self.btnClearQty.clicked.connect(self.clearqtyclick)


	# Function for connecting to SQL server.
	def connect(self):

		global cnxn
		global cursor
		# I'll probably replace the UID and PWD with something more secure.
		cnxn = pyodbc.connect("DSN=sqlserver;UID=FSUser;PWD=1@freeman")
		cursor = cnxn.cursor()

	# Function for closing cursor and disconnecting.
	def disconnect(self):
		cursor = cnxn.cursor()
		cursor.close()
		del cursor
		cnxn.close()

	def clockid_after_update(self):

		global ClockID

		ClockID = self.txtClockID.text()

		# Here we take the clock ID and return the Employee ID.
		# This is because the employee is familiar with their clock ID,
		# But for storing data we want to use their Employee ID.
		if self.get_employee_info() is False:
			self.clearForm()
			BarcodeApp.message = "You are not on the list of current users."
			return

		self.txtWOBOMID.setFocus()

	def get_employee_info(self):

		global EmpID, FirstName

		self.connect()

		# This is the select stmt for returning EmpID from ClockID.
		cursor.execute("SELECT ID, [First Name] AS FirstName, ClockID, [Status ID] " 
			"FROM Employees "
			"WHERE [Status ID] = 1 AND ClockID = ?", (ClockID))
		row = cursor.fetchone()

		if row:
			FirstName = row.FirstName
			EmpID = row.ID
		else:
			return False

		self.disconnect()

	# Function for moving to Part textbox after WOID has been scanned.
	# Return WOID information so the user can confirm.
	def wobomid_after_update(self):

		global WOID, ProductID, ProductCode, WOBOMID

		WOBOMID = self.txtWOBOMID.text()

		self.connect()

		cursor.execute("SELECT [Work Order BOM].[ID] AS WOBOMID, [Work Order].[WOID], [Work Order].[WO Status], "
			"[Products].[Product Code] AS WOName, [Products_1].[Product Code] AS MaterialName, [Products_1].[ID] AS ProductID "
			"FROM [Work Order] INNER JOIN [Products] ON [Work Order].[ProductID] = [Products].[ID] INNER JOIN "
			"[Work Order BOM] ON [Work Order].[WOID] = [Work Order BOM].[WOID] INNER JOIN Products AS Products_1 "
			"ON [Work Order BOM].[ProductID] = [Products_1].[ID] "
			"WHERE ([Work Order].[WO Status] = 2 OR [Work Order].[WO Status] = 3) AND ([Work Order BOM].[ID] = ?)", (WOBOMID))

		row = cursor.fetchone()

		if row:
			self.lblWOIDReturn.setText(row.WOName)
			self.lblPartReturn.setText(row.MaterialName)
			ProductID = int(row.ProductID)
			WOID = int(row.WOID)
			ProductCode = row.MaterialName
		else:
			self.txtWOBOMID.clear()
			self.txtWOBOMID.setFocus()
			self.lblWOIDReturn.setText('')
			return

		self.lblQty.setFocus()

		self.disconnect()

	# Click event for entering data to server.
	# First we validate data. This will become more extensive based on parameters.
	# Then we create a connection string. This will need to be more secure.
	# Following this will load variables from controls, then create INSERT statement.
	# Lastly, we reset controls to form load values.

	@pyqtSlot()

	def on_click(self):

		global rv

		if self.validate() == False:
			return

		#Quantity = self.slQty.value()
		#Quantity = int(self.lblQty.text())

		self.connect()

		# SQL string for stored procedure with a return value.
		sql = """\
		DECLARE @outRV int;
		EXEC @outRV = [dbo].[sprocFIFOIssueInventory] @pProductID = ?, @pQtyToIssue = ?, @pWOBOMID = ?, @pEmpID = ?;
		SELECT @outRV AS RV;
		"""

		# Parameters to feed into the stored procedure.
		params = (ProductID, Quantity, WOBOMID, EmpID, )
		try:
			cursor.execute(sql, params)
		except Exception as e:
			print(e)

		# Fetch the return value (either 1, 2, or 3)
		return_value = cursor.fetchone()

		# Turn it into a string.
		return_value = str(return_value)

		# Get rid of extra stuff and put it into a list.
		rv = [int(s) for s in re.findall(r'\d+', return_value)]

		# Pull it out of the list.
		rv = rv[0]

		# Commit the SPROC. Not doing this locks up the tables.
		cnxn.commit()

		# Clear form.
		self.clearForm()

		# Evaluate RV and determine message to be returned.
		self.check_return_value()

		# Close cursor and disconnect.
		self.disconnect()

	# Evaluate for the return value. These numbers correspond to the RETURNs on the sproc.
	# Create appropriate messageboxes. If it failed it didn't go through already.
	# We don't have to worry about RETURN here, or exiting early.
	def check_return_value(self):
		if rv == 1:
			BarcodeApp.message = "Thanks " + FirstName + "! You issued " + str(Quantity) + \
				" units of " + ProductCode + " to Work Order: " + str(WOID)
			self.call_msg_timer()
		elif rv == 2:
			BarcodeApp.message = "Unable to issue. Inventory shows none remaining. Please contact the tool-room."
			self.call_msg_timer()
		elif rv == 3:
			BarcodeApp.message = "This item is not part of the Work Order's BOM. Please contact the tool-room."
			self.call_msg_timer()

	# Create a clock and call start	
	def clock(self):
		timer = QTimer(self)
		timer.timeout.connect(self.showTime)
		timer.start(1000)
		self.showTime

	# Time to be shown on clock, using system time.
	def showTime(self):
		time = QTime.currentTime()
		text = time.toString('hh:mm')
		if (time.second() % 2) == 0:
			text = text[:2] + ' ' + text[3:]

		self.lcdTime.display(text)

	def btn1Click(self):

		QtyClicked = '1'
		self.show_lbl_qty(QtyClicked)

	def btn2Click(self):

		QtyClicked = '2'
		self.show_lbl_qty(QtyClicked)

	def btn3Click(self):

		QtyClicked = '3'
		self.show_lbl_qty(QtyClicked)

	def btn4Click(self):

		QtyClicked = '4'
		self.show_lbl_qty(QtyClicked)

	def btn5Click(self):

		QtyClicked = '5'
		self.show_lbl_qty(QtyClicked)

	def btn6Click(self):

		QtyClicked = '6'
		self.show_lbl_qty(QtyClicked)

	def btn7Click(self):

		QtyClicked = '7'
		self.show_lbl_qty(QtyClicked)

	def btn8Click(self):

		QtyClicked = '8'
		self.show_lbl_qty(QtyClicked)

	def btn9Click(self):

		QtyClicked = '9'
		self.show_lbl_qty(QtyClicked)

	def btn0Click(self):

		QtyClicked = '0'
		self.show_lbl_qty(str(QtyClicked))

	# Function for showing label value on change.
	def show_lbl_qty(self, QtyClicked):
		global Quantity
		newNumber = self.lblQty.text() + str(QtyClicked)
		self.lblQty.setText(newNumber)
		Quantity = int(self.lblQty.text())

	def clearqtyclick(self):
		self.lblQty.setText('')

	# Function for validating data.
	def validate(self):

		if self.txtWOBOMID.text() == "":
			BarcodeApp.message = "Error: Please re-scan WOBOMID."
			self.call_msg_timer()
			self.txtWOBOMID.clear()
			self.txtWOBOMID.setFocus()
			return False
		elif self.lblQty.text() == "":
			BarcodeApp.message = "Error: Please select a quantity, 30 or fewer."
			self.call_msg_timer()
			self.lblQty.clear()
			self.lblQty.setFocus()
			return False
		elif Quantity > 30 or Quantity is None:
			BarcodeApp.message = "Error: Please select a quantity, 30 or fewer."
			self.call_msg_timer()
			self.lblQty.clear()
			self.lblQty.setFocus()
			return False

	# This method calls the class from any other method using it's
	# inherited message. Essentially it connects the originating 
	# method and the messagebox timer class.
	def call_msg_timer(self):
		msgBox = TimerMessageBox(10, self)
		msgBox.exec()

	# Function for clearing form and resetting focus to on load parameters.
	def clearForm(self):
		self.txtWOBOMID.clear()
		self.txtClockID.clear()
		self.txtClockID.setFocus()
		self.lblWOIDReturn.setText('')
		self.lblPartReturn.setText('')
		self.lblQty.setText('')
			
# New class for a messagebox timer, so employees don't have to click the OK button.
# I display any message for 10 seconds by default.
# In the methods above, I define the message then call the msgbox timer method,
# which in turn calls this class to initiate.
class TimerMessageBox(QMessageBox):

	def __init__(self, timeout=10, parent=None):
		super(TimerMessageBox, self).__init__(parent)
		self.setWindowTitle('Message')
		self.time_to_wait = timeout
		self.setText(BarcodeApp.message)
		self.setStandardButtons(QMessageBox.Ok)
		self.timer = QTimer(self)
		self.timer.setInterval(1000)
		self.timer.timeout.connect(self.change_timer)
		self.change_timer()
		self.timer.start()

	def change_timer(self):
		self.time_to_wait -= 1
		if self.time_to_wait <= 0:
			self.close()

	def closeEvent(self, event):
		self.timer.stop()
		event.accept()

# This part is what calls the main form. I don't entirely understand how this works.
def main():

	app = QApplication(sys.argv)
	form = BarcodeApp()
	form.show()
	app.exec_()

if __name__== '__main__':

	main()