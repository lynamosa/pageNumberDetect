# -*- coding: utf-8 -*-
"""
Created on Fri Jun 28 09:31:42 2024

@author: zeromons
"""

from PyQt5 import uic
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QPushButton,
    QRadioButton, QLineEdit, QSpinBox, QComboBox, QCheckBox, QHeaderView,
    QMenu, QFileDialog, QAbstractItemView, QMessageBox, QInputDialog
)

import fitz  # PyMuPDF
import sys
import os
from math import ceil
import numpy as np
from PyQt5.QtCore import pyqtSignal, QObject, QPoint, Qt, QSettings
# import inspect

# def print_params(func):
#     def wrapper(*args, **kwargs):
#         # Lấy tên của các tham số và giá trị của chúng
#         func_args = inspect.signature(func).parameters
#         bound_args = inspect.signature(func).bind(*args, **kwargs)
#         bound_args.apply_defaults()

#         print(f"Calling {func.__name__} with:")
#         for name, value in bound_args.arguments.items():
#             print(f"    {name} = {value}")
#         return func(*args, **kwargs)
#     return wrapper

pt2mm = 595/210
paperSizes = {
    'A0': [2384, 3370],
    'A1': [1684, 2384],
    'A2': [1191, 1684],
    'A3': [842, 1191],
    'A4': [595, 842],
    'A5': [420, 595],
    'A6': [298, 420]
}
typeCopies = ['by_1', 'by_fill', 'by_column', 'by_row']
# pagelog = open('F:/users/desktops/test/log.txt', 'w', encoding='utf8')


def show_message_box(parent, title, message):
    msg_box = QMessageBox(parent)
    msg_box.setWindowTitle(title)
    msg_box.setText(message)
    msg_box.setIcon(QMessageBox.Information)
    msg_box.setStandardButtons(QMessageBox.Ok)
    msg_box.exec_()

class FileListModel(QObject):
    dataChanged = pyqtSignal()

    def __init__(self, fileLists):
        super().__init__()
        self._fileLists = fileLists

    def add_file(self, path, name):
        self._fileLists.append([path, name])
        self.dataChanged.emit()

    def remove_file(self, index):
        if 0 <= index < len(self._fileLists):
            self._fileLists.pop(index)
            self.dataChanged.emit()

    def update_file(self, index, path, name):
        if 0 <= index < len(self._fileLists):
            self._fileLists[index] = [path, name]
            self.dataChanged.emit()
            
    def clear_files(self):
        self._fileLists.clear()
        self.dataChanged.emit()
        
    def get_file_lists(self):
        return self._fileLists
  
def tbl_aray(numPages, shuff_type, columns, rows, horizontal, duplet):
    if shuff_type == 'by_fill':
        return np.repeat(np.arange(numPages, dtype=np.int32), columns * rows)
        # return pagelist.reshape((numPages, rows, columns))

    tbl_z = ceil(numPages / (columns * rows * (2 if duplet else 1)))
    # tbl_z = (tbl_z + 1) // 2 * 2 if duplet else tbl_z
    tbl_z = tbl_z * 2 if duplet else tbl_z

    pagelist = np.arange(tbl_z * columns * rows, dtype=np.int32)

    if shuff_type == 'by_column' or shuff_type == 'by_row':
        horizontal = 'ltr'

    if horizontal == 'ltr':
        if shuff_type == 'by_1':
            indices = pagelist.reshape((rows, columns, tbl_z)).transpose(2, 0, 1)
        elif shuff_type == 'by_column':
            indices = pagelist.reshape((rows, 1, columns * tbl_z)).transpose(2, 0, 1)
            indices = np.repeat(indices, columns, axis=2)
        elif shuff_type == 'by_row':
            indices = pagelist.reshape((1, columns, rows * tbl_z)).transpose(2, 0, 1)
            indices = np.repeat(indices, rows, axis=1)
    else:
        indices = pagelist.reshape((columns, rows, tbl_z)).transpose(2, 1, 0)

    if duplet:
        indices[1::2] = np.flip(indices[1::2], axis=2)

    cells = np.where(indices >= numPages, None, indices)
    return cells

def page_orientation(width, height, rotation=0):
    ori = width < height
    if rotation== 90 or rotation == 270:
        ori = not ori
    return "portrait" if ori else "landscape"

class MergePagesApp(QMainWindow):
    def __init__(self, file_model):
        
        super(MergePagesApp, self).__init__()
        uic.loadUi('merge_page.ui', self)
        self.setWindowFlags(self.windowFlags() | 262144) #Qt.WindowStaysOnTopHint
        self.setAcceptDrops(True)
        
        self.file_model = file_model
        
        #Action control
        self.listFile = self.findChild(QTableWidget, 'listFile')
        self.listFile.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.listFile.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.listFile.setContextMenuPolicy(Qt.CustomContextMenu)
        self.listFile.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.listFile.customContextMenuRequested.connect(self.open_context_menu)
        self.update_table()
        self.file_model.dataChanged.connect(self.update_table)
        
        self.previewButton = self.findChild(QPushButton, 'previewButton')
        self.addStyle = self.findChild(QPushButton, 'btnAddStyle')
        self.removeStyle = self.findChild(QPushButton, 'btnRemoveStyle')
        self.okButton = self.findChild(QPushButton, 'okButton')
        
        self.previewButton.clicked.connect(self.preview_setup)
        self.okButton.clicked.connect(self.merge_pages)
        self.addStyle.clicked.connect(self.add_style)
        self.removeStyle.clicked.connect(self.remove_style)
        
        
        #Radio control
        self.radioPrefix = self.findChild(QRadioButton, 'radioSavePrefix')
        self.radioSuffix = self.findChild(QRadioButton, 'radioSaveSuffix')
        self.radioSubFolder = self.findChild(QRadioButton, 'radioSaveSubFolder')
        self.radioPath = self.findChild(QRadioButton, 'radioSavePath')
        
        self.pgAutoSize = self.findChild(QRadioButton, 'pgAutoSize')
        self.pgStandardSize = self.findChild(QRadioButton, 'pgStandardSize')
        self.pgCustomSize = self.findChild(QRadioButton, 'pgCustomSize')

        self.savePath = 'subFolder'
        self.pageSize = [595,842]

        self.radioPrefix.clicked.connect(self.update_save_path)
        self.radioSuffix.clicked.connect(self.update_save_path)
        self.radioSubFolder.clicked.connect(self.update_save_path)
        self.radioPath.clicked.connect(self.update_save_path)
        
        self.pgStandardSize.toggled.connect(self.update_page_size)
        self.pgCustomSize.toggled.connect(self.update_page_size)
        
        #Value data
        # self.savePrefix = self.findChild(QLineEdit, 'txtSavePre')
        self.saveSub = self.findChild(QLineEdit, 'txtSaveSubFolder')
        self.savePathtxt = self.findChild(QLineEdit, 'txtSavePath')
        self.gridColumns = self.findChild(QSpinBox, 'gridColumns')
        self.gridRows = self.findChild(QSpinBox, 'gridRows')
        self.copies = self.findChild(QComboBox, 'copies')
        self.direction = self.findChild(QComboBox, 'gridDirection')
        self.leftMargin = self.findChild(QSpinBox, 'leftMargin')
        self.rightMargin = self.findChild(QSpinBox, 'rightMargin')
        self.topMargin = self.findChild(QSpinBox, 'topMargin')
        self.bottomMargin = self.findChild(QSpinBox, 'bottomMargin')

        self.horizontalSpacing = self.findChild(QSpinBox, 'hSpace')
        self.verticalSpacing = self.findChild(QSpinBox, 'vSpace')
        
        self.sizeStandard = self.findChild(QComboBox, 'StandardSize')
        self.sizeStandardOri = self.findChild(QComboBox, 'StandardOri')
        self.sizeCustomW = self.findChild(QSpinBox, 'customWidth')
        self.sizeCustomH = self.findChild(QSpinBox, 'customHeight')
        
        self.duplet = self.findChild(QCheckBox, 'dupletCheckbox')
        self.rotate = self.findChild(QCheckBox, 'rotateCheckbox')
        self.border = self.findChild(QCheckBox, 'borderCheckbox')
        self.ratio = self.findChild(QCheckBox, 'ratioCheckbox')
        self.styleList = self.findChild(QComboBox, 'comboboxStyle')
        self.styleList.currentIndexChanged.connect(self.load_style)
        
        self.load_styles()
        self.load_style()
        
        
        
    def update_table(self):
        fileLists = self.file_model.get_file_lists()
        self.listFile.setRowCount(len(fileLists))
        for row, file_info in enumerate(fileLists):
            self.listFile.setItem(row, 0, QTableWidgetItem(file_info[0]))
            self.listFile.setItem(row, 1, QTableWidgetItem(file_info[1]))
    
    def remove_file(self):
        selected_rows = sorted(set(index.row() for index in self.listFile.selectedIndexes()), reverse=True)
        for row in selected_rows:
            self.file_model.remove_file(row)
            
    def open_context_menu(self, position: QPoint):
        context_menu = QMenu(self)
        
        add_action = context_menu.addAction("Add file")
        addFolder_action = context_menu.addAction("Add Folder")
        remove_action = context_menu.addAction("Remove")
        clear_action = context_menu.addAction("Clear")
        
        action = context_menu.exec_(self.listFile.viewport().mapToGlobal(position))
        
        if action == add_action:
            self.open_file_dialog()
        elif action == addFolder_action:
            self.open_folder_dialog()
        elif action == remove_action:
            self.remove_file()
        elif action == clear_action:
            self.file_model.clear_files()
            
    def merge_pages(self, event):
        self.update_save_path()
        self.update_page_size()
        for path, fileName in fileLists:
            input_path = os.path.join(path, fileName)
            if self.savePath == 'prefix':
                output_path = os.path.join(path, 'layout_' + fileName)
            elif self.savePath == 'suffix':
                output_path = os.path.join(path, f'{fileName[:-4]}_{self.gridColumns.value()}x{self.gridRows.value()}.pdf')
            elif self.savePath == 'subFolder':
                output_dir = os.path.join(path, self.saveSub.text())
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                output_path = os.path.join(output_dir, fileName)
            elif self.savePath == 'folder':
                output_dir = self.savePathtxt.text()
                if not os.path.exists(output_dir):
                    os.makedirs(output_dir)
                output_path = os.path.join(output_dir, fileName)
            else:
                print("OUTPUT NOT FOUND", self.savePath)
                return None
            
            page_size = self.pageSize  # A4 size in points (width, height)
            margins = [x*pt2mm for x in [self.topMargin.value(),self.bottomMargin.value(),self.leftMargin.value(),self.rightMargin.value()]]  # margins (top, bottom, left, right) in points
            spacing = [x*pt2mm for x in [self.horizontalSpacing.value(), self.verticalSpacing.value()]]  # spacing (horizontal, vertical) in points
            grid_size = [self.gridColumns.value(), self.gridRows.value()]  # number of pages (columns, rows)
            isDuplet = self.duplet.isChecked()
            isBorder = self.border.isChecked()
            isRotate = self.rotate.isChecked()
            isRatio = self.ratio.isChecked()
            self.create_nup_pdf(input_path, output_path, page_size, margins, spacing, grid_size, isDuplet, isBorder, isRotate, isRatio)
        self.show_message_box("FINISH", 'File had been N-Up completed')
            
    def show_message_box(self, title, message):
        show_message_box(self, title, message)
        
    def create_nup_pdf(self, input_path, output_path, page_size, margins, spacing, grid_size, duplet, border, rotate, ratio):
        input_pdf = fitz.open(input_path)
        output_pdf = fitz.open()

        page_width, page_height = page_size
        margin_top, margin_bottom, margin_left, margin_right = margins
        spacing_x, spacing_y = spacing
        num_cols, num_rows = grid_size

        # Kích thước mỗi ô trang nhỏ
        cell_width = (page_width - margin_left - margin_right - (num_cols - 1) * spacing_x) / num_cols
        cell_height = (page_height - margin_top - margin_bottom - (num_rows - 1) * spacing_y) / num_rows
        cell_orientation = page_orientation(cell_width, cell_height)

        num_pages = input_pdf.page_count

        # Tạo các trang PDF mới với cấu hình N-Up
        copies = typeCopies[self.copies.currentIndex()]
        direction = 'ltr' if self.direction.currentIndex()==0 else 'ttb'
        pageList = tbl_aray(num_pages, copies, num_cols, num_rows, direction, True)
        # print(pageList, sep='\n', end='\n================')
        pageList = np.ravel(pageList)
        pageList = pageList.tolist()

        for i in range(0, len(pageList), num_cols * num_rows):
            new_page = output_pdf.new_page(width=page_width, height=page_height)
            mLeft = margin_left if (duplet==True and (i//(num_cols * num_rows))%2 ==0) else margin_right
            for j in range(num_cols * num_rows):
                if i + j < len(pageList):
                    # Tính toán vị trí trong lưới N-Up
                    row = j // num_cols
                    col = j % num_cols
                    x0 = mLeft + col * (cell_width + spacing_x)
                    y0 = margin_top + row * (cell_height + spacing_y)
                    x1 = x0 + cell_width
                    y1 = y0 + cell_height

                    # Định nghĩa vùng đích trên trang mới
                    target_rect = fitz.Rect(x0, y0, x1, y1)

                    # Đặt trang PDF vào vùng đích trên trang mới
                    if pageList[i + j]!=None:
                        pg = input_pdf[pageList[i + j]]
                        pg_orientation = page_orientation(pg.mediabox_size.x, pg.mediabox_size.y, pg.rotation)
                        rotate = 90 #pg.rotation if cell_orientation == pg_orientation else (pg.rotation+90)%360
                        # print(i+j+1, cell_orientation, pg_orientation, int(pg.mediabox_size.x), int(pg.mediabox_size.y), pg.rotation, rotate)
                        new_page.show_pdf_page(target_rect, input_pdf, pageList[i + j], keep_proportion=ratio, rotate=rotate)
                            

                    # Vẽ đường viền cho vùng ghép
                    if border==True:
                        border_rect = fitz.Rect(x0, y0, x1, y1)
                        new_page.draw_rect(border_rect, color=(1, 0, 0), width=0.5)  # Đường viền màu đỏ, độ rộng 0.5

        # Lưu PDF đầu ra
        output_pdf.save(output_path)
        output_pdf.close()
        input_pdf.close()
        
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        for url in urls:
            if url.isLocalFile() and url.toLocalFile().endswith('.pdf'):
                directory, file_name = os.path.split(url.toLocalFile())
                self.file_model.add_file(directory, file_name)
        event.acceptProposedAction()
    
    def open_file_dialog(self):
        files, x = QFileDialog.getOpenFileNames(self, "Select PDF Files", "", "PDF Files (*.pdf)")
        for file in files:
            directory, file_name = os.path.split(file)
            self.file_model.add_file(directory, file_name)

    def open_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            for filename in os.listdir(folder):
                if filename.endswith(".pdf"):
                    self.file_model.add_file(folder, filename)
    
    def preview_setup(self):
        pass

    def update_save_path(self):
        if self.radioPrefix.isChecked():
            self.savePath = 'prefix'
        if self.radioSuffix.isChecked():
            self.savePath = 'suffix'
        elif self.radioSubFolder.isChecked():
            self.savePath = 'subFolder'
        elif self.radioPath.isChecked():
            self.savePath = 'folder'

    def update_page_size(self):
        if self.pgStandardSize.isChecked():
            self.pageSize = 2
            self.pageSize = paperSizes[self.sizeStandard.currentText()]
            if self.sizeStandardOri.currentIndex() == 1:
                self.pageSize=list(reversed(self.pageSize))
        elif self.pgCustomSize.isChecked():
            self.pageSize = [self.sizeCustomW.value()*pt2mm,self.sizeCustomH.value()*pt2mm]
    
    def load_styles(self):
        settings = QSettings("config.ini", QSettings.IniFormat)
        self.styleList.clear()
        styles = settings.childGroups()
        self.styleList.addItems(styles)

    def add_style(self):
        style_name, ok = QInputDialog.getText(self, 'Input Dialog', 'Enter style name:')
        if ok:
            self.save_style(style_name)
            self.load_styles()
    
    def load_style(self):
        style_name = self.styleList.currentText()
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.beginGroup(style_name)
        if settings.value('radioSavePrefix', type=bool)==True:
            self.radioPrefix.setChecked(True)
        elif settings.value('radioSaveSuffix', type=bool)==True:
            self.radioSuffix.setChecked(True)
        elif settings.value('radioSaveSubFolder', type=bool)==True:
            self.radioSubFolder.setChecked(True)
        elif settings.value('radioSavePath', type=bool)==True:
            self.radioPath.setChecked(True)
        
        if settings.value('pgAutoSize', type=bool)==True:
            self.pgAutoSize.setChecked(True)
        elif settings.value('pgStandardSize', type=bool)==True:
            self.pgStandardSize.setChecked(True)
        
        self.pgCustomSize.setChecked(settings.value('pgCustomSize', type=bool))
        # self.savePrefix.setText(settings.value('txtSavePre', type=str))
        self.saveSub.setText(settings.value('txtSaveSubFolder', type=str))
        self.savePathtxt.setText(settings.value('txtSavePath', type=str))
        self.gridColumns.setValue(settings.value('gridColumns', type=int))
        self.gridRows.setValue(settings.value('gridRows', type=int))
        self.copies.setCurrentIndex(settings.value('copies', type=int))
        self.direction.setCurrentIndex(settings.value('gridDirection', type=int))
        self.leftMargin.setValue(settings.value('leftMargin', type=int))
        self.rightMargin.setValue(settings.value('rightMargin', type=int))
        self.topMargin.setValue(settings.value('topMargin', type=int))
        self.bottomMargin.setValue(settings.value('bottomMargin', type=int))
        self.horizontalSpacing.setValue(settings.value('hSpace', type=int))
        self.verticalSpacing.setValue(settings.value('vSpace', type=int))
        self.sizeStandard.setCurrentIndex(settings.value('StandardSize', type=int))
        self.sizeStandardOri.setCurrentIndex(settings.value('StandardOri', type=int))
        self.sizeCustomW.setValue(settings.value('customWidth', type=int))
        self.sizeCustomH.setValue(settings.value('customHeight', type=int))
        self.duplet.setChecked(settings.value('dupletCheckbox', type=bool))
        self.rotate.setChecked(settings.value('rotateCheckbox', type=bool))
        self.border.setChecked(settings.value('borderCheckbox', type=bool))
        self.ratio.setChecked(settings.value('ratioCheckbox', type=bool))
        settings.endGroup()
    
    def save_style(self, style_name):
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.beginGroup(style_name)
        settings.setValue('radioSavePrefix', self.radioPrefix.isChecked())
        settings.setValue('radioSaveSuffix', self.radioSuffix.isChecked())
        settings.setValue('radioSaveSubFolder', self.radioSubFolder.isChecked())
        settings.setValue('radioSavePath', self.radioPath.isChecked())
        if self.pgStandardSize.isChecked():
            settings.setValue('pgStandardSize', True)
            settings.setValue('pgAutoSize', False)
        else:
            settings.setValue('pgStandardSize', False)
            settings.setValue('pgAutoSize', True)
            
        settings.setValue('pgCustomSize', self.pgCustomSize.isChecked())
        # settings.setValue('txtSavePre', self.savePrefix.text())
        settings.setValue('txtSaveSubFolder', self.saveSub.text())
        settings.setValue('txtSavePath', self.savePathtxt.text())
        settings.setValue('gridColumns', self.gridColumns.value())
        settings.setValue('gridRows', self.gridRows.value())
        settings.setValue('copies', self.copies.currentIndex())
        settings.setValue('gridDirection', self.direction.currentIndex())
        settings.setValue('leftMargin', self.leftMargin.value())
        settings.setValue('rightMargin', self.rightMargin.value())
        settings.setValue('topMargin', self.topMargin.value())
        settings.setValue('bottomMargin', self.bottomMargin.value())
        settings.setValue('hSpace', self.horizontalSpacing.value())
        settings.setValue('vSpace', self.verticalSpacing.value())
        settings.setValue('StandardSize', self.sizeStandard.currentIndex())
        settings.setValue('StandardOri', self.sizeStandardOri.currentIndex())
        settings.setValue('customWidth', self.sizeCustomW.value())
        settings.setValue('customHeight', self.sizeCustomH.value())
        settings.setValue('dupletCheckbox', self.duplet.isChecked())
        settings.setValue('rotateCheckbox', self.rotate.isChecked())
        settings.setValue('borderCheckbox', self.border.isChecked())
        settings.setValue('ratioCheckbox', self.ratio.isChecked())
        settings.endGroup()

    def remove_style(self):
        style_name = self.styleList.currentText()
        settings = QSettings("config.ini", QSettings.IniFormat)
        settings.beginGroup(style_name)
        settings.remove("")
        settings.endGroup()
        self.load_styles()
        
if __name__ == "__main__":
    fileLists = []

    app = QApplication(sys.argv)
    file_model = FileListModel(fileLists)
    main_window = MergePagesApp(file_model)
    main_window.show()
    sys.exit(app.exec_())

