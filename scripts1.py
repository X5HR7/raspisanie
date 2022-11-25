from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from docx import Document

import shutil


#копируем базовую таблицу и вставляем по указанному адресу
def script_1(src: str, dst: str) -> None:
    shutil.copy2(src=src, dst=dst)


async def script_2(file_name: str, key_word: str, output_table: str) -> None:
    #открытие таблицы с расписанием
    woorkbook_input = load_workbook(filename=file_name)
    ws1 = woorkbook_input.active
    #открытие таблицы, в которую будет записан результат
    woorkbook_base = load_workbook(filename=output_table)
    ws2 = woorkbook_base.active

    #перебираем все номера строк с таблице
    for row in range(1, ws1.max_row+1):
        #перебираем все номера столбцов в таблице
        async for col in range(1, ws1.max_column+1):
            #получение ячейки по адресу: i, c; row-№строки, col-№столбца
            cell = ws1._get_cell(row, col)
            #получение значения ячейки
            value = str(cell.value)

            #проверка наличия ключевого слова в ячейке (фамилия преподавателя)
            if key_word in value.split():
                #время пары
                time = del_all_spaces(ws1._get_cell(row=cell.row, column=2).value)
                #название группы
                group = get_group(del_spaces(ws1._get_cell(row=cell.row-1-offset[time], column=cell.column).value))
                #день недели
                day = del_spaces(ws1._get_cell(row=cell.row-offset[time], column=1).value)

                #base_days_1[day]: из словаря по ключу "day" получаем букву столбца таблицы
                #base_times[time]: из словаря по ключу "time" получаем номер строки
                #ws2[f'{base_days[day]}{base_times[time]}'] - ячейка в таблице вывода
                if ws2[f'{base_days_1[day]}{base_times[time]}'].value == None:
                    #если ячейка пустая, устанавливаем значение: группа
                    ws2[f'{base_days_1[day]}{base_times[time]}'] = f'{group}\n'
                    ws2[f'{base_days_2[day]}{base_times[time]}'] = f'{group}\n'
                else:
                    #если в ячейке уже есть группа, добавляем еще одну
                    ws2[f'{base_days_1[day]}{base_times[time]}'] = ws2[f'{base_days_1[day]}{base_times[time]}'].value+f'{group}\n'
                    ws2[f'{base_days_2[day]}{base_times[time]}'] = ws2[f'{base_days_2[day]}{base_times[time]}'].value+f'{group}\n'
                #сохраняем таблицу
                woorkbook_base.save(filename=output_table)
    #закрываем соединение с таблицами
    woorkbook_input.close()
    woorkbook_base.close()

#удаление лишних пробелов из строки
def del_spaces(string: str) -> str:
    if string != None:
        return ' '.join([str(i) for i in string.split()])

#удаление всех пробелов из строки
def del_all_spaces(string: str) -> str:
    if string != None:
        return string.replace(' ', '')

#удаление лишних символов из названия группы (из таблицы с расписанием)
def get_group(group_name: str) -> str:
    group = group_name.split('(')[-1]
    return group[:-1]


def script_3(document_name: str, key_word: str, excel_table_name: str) -> None:
    #создание экземпляра документа
    doc = Document(document_name)
    #получение списка всех таблиц из документа
    tables = doc.tables
    #заводим счетчик, по значению которого можно будет отпределить день недели (букву столбца в талице вывода)
    #используя значение счетчика, как ключ в словаре "doc_days"
    counter = 0

    #перебираем все таблицы в документе
    for table in tables:
        counter += 1
        #перебираем все строки в таблице
        for row in table.rows:
            #перебираем все ячейки в строке
            for cell in row.cells:
                #проверка наличия ключевого слова в ячейке
                if key_word in cell.text.split():
                    #список всех элементов строки
                    changes = [element.text for element in row.cells]

                    #открытие таблицы вывода
                    wb = load_workbook(excel_table_name)
                    ws = wb.active

                    for element in changes:
                        
                        if key_word in element.split():
                            #убирает пару из расписания
                            if changes.index(element) == 3 and changes[3] != changes[5]:
                                #changes[0] - номер пары
                                #changes[0]+1 - номер строки в таблице вывода
                                #doc_days[counter] - день недели
                                #ws[f'{doc_days[counter]}{int(changes[0])+1}'] - ячейка в таблице вывода со 
                                #столбцом doc_days[counter] и строкой changes[0]+1

                                #задает стиль ячейки (красный)
                                font = Font(color='52181b', bold=True)
                                fill = PatternFill(patternType='solid', fgColor='ff707a')
                                ws[f'{doc_days[counter]}{int(changes[0])+1}'].font = font
                                ws[f'{doc_days[counter]}{int(changes[0])+1}'].fill = fill
                                #try:
                                #    group_curr = f'{changes[1].split()[0]} - {changes[1].split()[1]}'
                                #    group_list = ws[f'{doc_days[counter]}{int(changes[0])+1}'].value.split('\n')
                                #except:
                                #    print(ws[f'{doc_days[counter]}{int(changes[0])+1}'].value, ws[f'{doc_days[counter]}{int(changes[0])+1}'].column_letter, ws[f'{doc_days[counter]}{int(changes[0])+1}'].row)
                                group_curr = f'{changes[1].split()[0]} - {changes[1].split()[1]}'
                                group_list = ws[f'{doc_days[counter]}{int(changes[0])+1}'].value
                                if group_list != None:
                                    group_list = ws[f'{doc_days[counter]}{int(changes[0])+1}'].value.split('\n')
                                else:
                                    group_list = ''
                                if group_curr in group_list:
                                    group_list.remove(group_curr)
                                    text = '\n'.join(group for group in group_list)
                                    ws[f'{doc_days[counter]}{int(changes[0])+1}'] = text

                                #сохранение изменений в таблице
                                wb.save(excel_table_name)

                            #добавляет пару в расписание
                            elif changes.index(element) == 5:
                                #задает стиль ячейки (зеленый)
                                font = Font(color='1b4715', bold=True)
                                fill = PatternFill(patternType='solid', fgColor='86d980')
                                ws[f'{doc_days[counter]}{int(changes[0])+1}'].font = font
                                ws[f'{doc_days[counter]}{int(changes[0])+1}'].fill = fill

                                #добавляет пару в рассписание
                                text_old = ws[f'{doc_days[counter]}{int(changes[0])+1}'].value
                                text_new = changes[1]
                                if text_old != None:
                                    ws[f'{doc_days[counter]}{int(changes[0])+1}'] = text_old+f'{text_new} \n'
                                else:
                                    ws[f'{doc_days[counter]}{int(changes[0])+1}'] = f'{text_new} \n'
                            
                                #сохранение изменений в таблице
                                wb.save(excel_table_name)

                    #закрываем соединение с таблицей                       
                    wb.close()

def script_4(excel_table_name: str) -> None:
    #открытие таблицы вывода
    wb = load_workbook(excel_table_name)
    ws = wb.active

    for col in range(3, 12, 2):
        for row in range(2, 8):
            #значение текущей ячейки
            cell_text = ws._get_cell(row=row, column=col).value
            #значение предыдущей ячейки
            prev_cell_text = ws._get_cell(row=row, column=col-1).value

            if cell_text != None and prev_cell_text != None:
                #список групп в текущей ячейке
                cell_list = cell_text.split('\n')
                #список групп в предыдущей ячейке
                prev_cell_list = prev_cell_text.split('\n')

                for group in prev_cell_list:
                    if group not in cell_list:
                        #задает стиль ячейки (синий)
                        font = Font(color='000000', bold=True)
                        fill = PatternFill(patternType='solid', fgColor='427bf5')
                        ws._get_cell(row=row, column=col).font = font
                        ws._get_cell(row=row, column=col).fill = fill
                        wb.save(excel_table_name)

    #закрываем соединение с таблицей 
    wb.close()


offset = {'08.45-10.15': 0, '10.30-12.00': 1, '12.40-14.10': 2, '14.20-15.50': 3, '16.00-17.30': 4, '17.40-19.10': 5}
#используется для определения столбца по названию дня
base_days_1 = {'понедельник': 'B', 'вторник': 'D', 'среда': 'F', 'четверг': 'H', 'пятница': 'J'}
base_days_2 = {'понедельник': 'C', 'вторник': 'E', 'среда': 'G', 'четверг': 'I', 'пятница': 'K'}
#используется для определения номера строки в таблице вывода по времени пары
base_times = {'08.45-10.15': 2, '10.30-12.00': 3, '12.40-14.10': 4, '14.20-15.50': 5, '16.00-17.30': 6, '17.40-19.10': 7}
#писпользуется для определения столбца по номеру таблице в ворд-документе
doc_days = {1: 'C', 2: 'E', 3: 'G', 4: 'I', 5: 'K'}