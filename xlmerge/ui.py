# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'xlmerge.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets

class Ui_Dialog(object):
    def setupUi(self, Dialog):
        Dialog.setObjectName("Dialog")
        Dialog.resize(862, 520)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(Dialog.sizePolicy().hasHeightForWidth())
        Dialog.setSizePolicy(sizePolicy)
        Dialog.setMinimumSize(QtCore.QSize(862, 520))
        Dialog.setMaximumSize(QtCore.QSize(862, 520))
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        Dialog.setFont(font)
        Dialog.setFocusPolicy(QtCore.Qt.WheelFocus)
        Dialog.setToolTip("")
        Dialog.setSizeGripEnabled(False)
        self.layoutWidget = QtWidgets.QWidget(Dialog)
        self.layoutWidget.setGeometry(QtCore.QRect(15, 6, 831, 491))
        self.layoutWidget.setObjectName("layoutWidget")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.layoutWidget)
        self.verticalLayout.setContentsMargins(0, 0, 0, 0)
        self.verticalLayout.setSpacing(12)
        self.verticalLayout.setObjectName("verticalLayout")
        self.labelBox = QtWidgets.QHBoxLayout()
        self.labelBox.setSpacing(0)
        self.labelBox.setObjectName("labelBox")
        self.label = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label.sizePolicy().hasHeightForWidth())
        self.label.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.labelBox.addWidget(self.label)
        self.labelFile = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.labelFile.sizePolicy().hasHeightForWidth())
        self.labelFile.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        self.labelFile.setFont(font)
        self.labelFile.setText("")
        self.labelFile.setObjectName("labelFile")
        self.labelBox.addWidget(self.labelFile)
        self.label_3 = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.label_3.sizePolicy().hasHeightForWidth())
        self.label_3.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.labelBox.addWidget(self.label_3)
        self.labelSheet = QtWidgets.QLabel(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.labelSheet.sizePolicy().hasHeightForWidth())
        self.labelSheet.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        self.labelSheet.setFont(font)
        self.labelSheet.setText("")
        self.labelSheet.setObjectName("labelSheet")
        self.labelBox.addWidget(self.labelSheet)
        self.verticalLayout.addLayout(self.labelBox)
        self.progressBar = QtWidgets.QProgressBar(self.layoutWidget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Fixed)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.progressBar.sizePolicy().hasHeightForWidth())
        self.progressBar.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        self.progressBar.setFont(font)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setInvertedAppearance(False)
        self.progressBar.setObjectName("progressBar")
        self.verticalLayout.addWidget(self.progressBar)
        self.table = QtWidgets.QTableWidget(self.layoutWidget)
        self.table.setMinimumSize(QtCore.QSize(0, 0))
        font = QtGui.QFont()
        font.setPointSize(10)
        self.table.setFont(font)
        self.table.setStyleSheet("QHeaderView::section { \n"
"    /* set the bottom border of the header, in order to set a single \n"
"    border you must declare a generic border first or set all other \n"
"    borders */\n"
"    border: none;\n"
"    border-bottom: 1px solid gray; \n"
"    border-right: 1px solid grey;\n"
"    border-radius: 3px;\n"
"    font-family: \"맑은 고딕\";\n"
"    font-size: 15px;\n"
"    font-weight: bold;\n"
"    min-height: 28px;\n"
"}\n"
"")
        self.table.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.table.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.table.setTabKeyNavigation(False)
        self.table.setProperty("showDropIndicator", False)
        self.table.setDragDropOverwriteMode(False)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.table.setRowCount(0)
        self.table.setColumnCount(0)
        self.table.setObjectName("table")
        self.table.horizontalHeader().setDefaultSectionSize(120)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(30)
        self.table.verticalHeader().setHighlightSections(False)
        self.verticalLayout.addWidget(self.table)
        self.buttonBox = QtWidgets.QHBoxLayout()
        self.buttonBox.setContentsMargins(0, 10, 50, -1)
        self.buttonBox.setSpacing(50)
        self.buttonBox.setObjectName("buttonBox")
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.buttonBox.addItem(spacerItem)
        self.btnNext = QtWidgets.QPushButton(self.layoutWidget)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.btnNext.setFont(font)
        self.btnNext.setObjectName("btnNext")
        self.buttonBox.addWidget(self.btnNext)
        self.btnFinish = QtWidgets.QPushButton(self.layoutWidget)
        font = QtGui.QFont()
        font.setFamily("맑은 고딕")
        font.setPointSize(10)
        font.setBold(False)
        font.setWeight(50)
        self.btnFinish.setFont(font)
        self.btnFinish.setObjectName("btnFinish")
        self.buttonBox.addWidget(self.btnFinish)
        self.verticalLayout.addLayout(self.buttonBox)

        self.retranslateUi(Dialog)
        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        _translate = QtCore.QCoreApplication.translate
        Dialog.setWindowTitle(_translate("Dialog", "xlmerge - 머리글 행 선택"))
        self.label.setText(_translate("Dialog", "File: "))
        self.label_3.setText(_translate("Dialog", "Sheet: "))
        self.btnNext.setText(_translate("Dialog", "다음"))
        self.btnFinish.setText(_translate("Dialog", "마침"))
