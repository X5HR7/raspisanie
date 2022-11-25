from docx import Document

def get_names_list(doc_name: str):
    doc = Document(doc_name)
    tables = doc.tables

    teachers_names = set()

    for table in tables:
        for row in table.rows:
            string = [element.text for element in row.cells]
            name_1 = list(string[3])
            name_2 = list(string[5])
            
            if name_1 != []:
                if '-' not in name_1 and name_1[-3] == '.':
                    teachers_names.add(rename(string[3]))
            if name_2 != []:
                if '-' not in name_2 and name_2[-3] == '.':
                    teachers_names.add(rename(string[5]))
                    
    return sorted(list(teachers_names))
        
def rename(string: str):
    if string[-1] == '.' and string[-3] == '.':
        return string
    elif string[-1] == ',' and string[-3] == '.':
        string = list(string)
        string[-1] = '.'
        return ''.join([s for s in string]) 

#choose word document path
print(get_names_list('docs/31.docx'))