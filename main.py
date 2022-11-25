from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog
from ui import Ui_MainWindow 
from scripts import script_1, script_2, script_3, script_4
from glob import glob

import sys
import os.path
import traceback
import time


class Window(QMainWindow):

    def __init__(self):
        #подключение графического интерфейса из файла ui.py 
        #ui.py генерируется автоматически через qt-designer
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        #подключение обработчика нажатия на кнопки
        self.add_funcs()
        #установка значения по умолчаннию в полях ввода
        self.set_start_values()

    #подключение функций-обработчиков нажатия на кнопки
    def add_funcs(self):
        self.ui.btn_table.clicked.connect(lambda: self.get_file_url(self.ui.lineEdit_table))
        self.ui.btn_doc.clicked.connect(lambda: self.get_file_url(self.ui.lineEdit_doc))
        self.ui.btn_output.clicked.connect(lambda: self.get_folder_url(self.ui.lineEdit_output))
        self.ui.btn_start.clicked.connect(self.start)


    def set_start_values(self):
        #значения по умолчанию (ищутся автоматически в папке с приложением)
        self.path = os.path.abspath('')
        self.excel_table = glob('*.xlsx')
        self.word_doc = glob('*.docx')
        self.key_word = ''

        #проверка наличия таблицы в папке с приложением и установка значения в поле ввода
        if self.excel_table != []:
            self.ui.lineEdit_table.setText(self.excel_table[0])
        else:
            self.ui.lineEdit_table.setPlaceholderText('None')
        #проверка наличия ворд-документа в папке с приложением и установка значения в поле ввода
        if self.word_doc != []:
            self.ui.lineEdit_doc.setText(self.word_doc[0])
        else: 
            self.ui.lineEdit_doc.setPlaceholderText('None')

        #установка значений по умолчанию в полях ввода (url папки вывода и фамилия преподавателя)
        self.ui.lineEdit_output.setText(f'{self.path}/')
        self.ui.lineEdit_key.setText(self.key_word)

    def get_file_url(self, object):
        #получение url файлов (таблицы и ворд-документа)
        file_url = QFileDialog.getOpenFileName(self)[0]
        #установка его в поле ввода
        object.setText(file_url)
    
    def get_folder_url(self, object):
        #получение url папки, в которую будет сохранен результат
        file_url = QFileDialog.getExistingDirectory(self)
        #установка его в поле ввода
        object.setText(file_url)

    def start(self):
        #значения из полей ввода
        #адрес таблицы, из которой будет взято расписание
        excel_table = self.ui.lineEdit_table.text()
        #адрес ворд-документа, из которого будут взяты замены
        word_doc = self.ui.lineEdit_doc.text()
        #фамилия преподавателя
        key_word = self.ui.lineEdit_key.text()
        #адрес папки, в которую будет сохранен результат
        output = self.ui.lineEdit_output.text()

        try:
            start_time = time.time()
            #подключение скриптов
            script_1(src='resources/base.xlsx', dst=f'{output}/raspisanie.xlsx')
            script_2(file_name=excel_table, key_word=key_word, output_table=f'{output}/raspisanie.xlsx')
            script_3(document_name=word_doc, key_word=key_word, excel_table_name=f'{output}/raspisanie.xlsx')
            script_4(excel_table_name=f'{output}/raspisanie.xlsx')

            self.ui.btn_start.setText('Done')
            self.ui.btn_start.setStyleSheet('background-color: rgb(77, 255, 121)')
            #убирает возможность нажать на кнопку "START" несколько раз подряд
            #self.ui.btn_start.blockSignals(True)
            print(time.time()-start_time)

        except:
            #установка стиля и значения для кпопки "START" при возникновении ошибки в процессе работы программы
            self.ui.btn_start.setText('Failed')
            self.ui.btn_start.setStyleSheet('background-color: rgb(255, 77, 77)')
            
            with open('logs.txt', 'w', encoding='UTF-8') as file:
                file.write(traceback.format_exc())

def app():
    app = QApplication(sys.argv)
    #создание окна и отображение его на экране
    window = Window()
    window.setWindowTitle('')
    window.show()
    #добавление возможности выхода из приложения
    sys.exit(app.exec_())

#запуск приложения
if __name__ == '__main__':
    app()