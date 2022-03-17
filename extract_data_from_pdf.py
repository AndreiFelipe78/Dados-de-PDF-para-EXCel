import fitz
import os 
import openpyxl 

files_path = r'/home/andrei/Downloads/tickets'
excel_path = r'/home/andrei/Documents/data_from_pdf/controle de ticket.xlsx'

def extract_data(file):
    with fitz.open(file) as pdf:
        text = ""
        for page in pdf:
            text += f'{page.get_text()}\n'
    
    data_for_mtr = list(text.split('\n'))
    my_list = []
    index_needed = [3, 5, 13, 17, 21, 31, 34, 36]
    for i in index_needed:
        my_list.append(data_for_mtr[i].strip())
    
    return my_list

def store_data(data):
    book = openpyxl.load_workbook(excel_path)
    #selecting a page to work with  
    ticket_page = book['Ticket de pesagem']
    #entering data on the page
    ticket_page.append(data)    
    book.save('controle de ticket.xlsx')


def list_file(dir):
    file_names = os.listdir(dir)
    for item in file_names:
        if item.endswith('pdf'):
            abs =  os.path.abspath(os.path.join(dir, item))
            std_exit = extract_data(abs)
            store_data(std_exit)
        


if __name__ == '__main__':
    list_file(files_path)