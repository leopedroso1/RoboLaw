# -*- coding: utf-8 -*-
"""
Created on Sun Aug  2 16:40:26 2020

@author: Leonardo

University of Oxford - Department of Computer Science - AE03

"""

# NEXT FEATURES! 
# Retry system. If something fails or is not doing correctly. Stop and retry.
# Verification of existing processes to be closed before start.
# GUI
# Install external components automatically at first execution

import pip
import os
import pyautogui # pip install selenium
import subprocess
import time
import re
import glob # Read file names with RegEx
import shutil # Move files easily
import ctypes # Mesage box and alerts

from selenium import webdriver # pip install selenium
from selenium.webdriver.support.ui import Select # Support for user interface
from selenium.webdriver.support.ui import WebDriverWait # Wait until the page is complete
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from datetime import date

tjsp_cadernos = 'https://dje.tjsp.jus.br/cdje/index.do;jsessionid=904E03097EDCF133765EE6CB860DFD6B.cdje2'
esaj_consulta_processos = 'https://esaj.tjsp.jus.br/sajcas/login?service=https%3A%2F%2Fesaj.tjsp.jus.br%2Fesaj%2Fj_spring_cas_security_check'

adobePath = r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"
files_path = os.path.expanduser('~\Downloads')
robot_path = os.path.expanduser('~\Desktop')

full_text = []

time_control = 5

key = "ADD YOUR ID (BRAZIL CPF) OR KEY FOR ESAJ PORTAL HERE"
raisepass = "ADD YOUR PASSWORD HERE"

execution_control = [0, 0, 0, 0, 0, 0, 0]

def import_or_install(package):

    try:

        __import__(package)

    except ImportError:

        pip.main(['install', package])  


# 1.Setup environment --> Check if the environment is ok
def setup_environment():

    print("Checking system requirements")
    
#    msgbox_val = ctypes.windll.user32.MessageBoxW(0, "", "Robo Law - version 1.2.5", 0)
    
    execution_control[0] = 1
    
    return 1

# 2. Download pdf file from justice court web page
def collect_data():
    
    chrome_driver_path = os.path.dirname(os.path.abspath(__file__)) + '\chromedriver_84'
    browser = webdriver.Chrome(executable_path = chrome_driver_path) 
    browser.get(tjsp_cadernos)
    
    search_elem = WebDriverWait(browser, 3).until(EC.presence_of_element_located((By.ID, 'consultar')))

    # Select the notebooks to be downloaded. We will download the entire file for performance purposes        
    select_cadastro = Select(browser.find_element_by_id('cadernosCad'))
    select_cadastro.select_by_value("12") 

    elem_consulta = browser.find_element_by_id('download') 
    elem_consulta.click()

    # If the download is not completed, wait, otherwise Go!   
    while not os.path.exists(files_path + '\caderno3-Judicial-1ªInstancia-Capital.pdf'):
   
        time.sleep(1)

    if os.path.isfile(files_path + '\caderno3-Judicial-1ªInstancia-Capital.pdf'):

        browser.quit()

        execution_control[1] = 1        

    else:

        print("1x00: Download failed. Please reboot the process")
        
        browser.quit()
        
    return 1

# 3. Converts PDF to TXT using adobe reader
def process_pdf_to_txt(file_name, file_type):
        
    # Open adobe reader and use it to save as .txt
    
    if file_type == 1: # 1 stands for 'one heavy' processing files

        subprocess.Popen("%s %s" % (adobePath, files_path + file_name + ".pdf"))
        
        time.sleep(10)
        
        pyautogui.keyDown('alt')
    
        time.sleep(5)
    
        pyautogui.press('q')
    
        time.sleep(5)
    
        pyautogui.press('v')
        
        time.sleep(5)
        
        pyautogui.keyUp('alt')
        
        time.sleep(5)
        
        print(files_path + file_name + '.txt')
        
        pyautogui.write(files_path + file_name + '.txt') 
        
        time.sleep(5)
        
        pyautogui.press('enter')
        
        time.sleep(6)
        
        # Check if the transformation for txt is complete
        current_size = os.stat(files_path + '\caderno3-Judicial-1Instancia-Capital' + '.txt').st_size
        previous_size = -1        
    
        width = int((pyautogui.size()[0]) / 2)
        height = int((pyautogui.size()[1]) / 2) 
            
        pyautogui.moveTo(width,height)
        
        while current_size != previous_size:
            
            previous_size = current_size
        
            pyautogui.click()
            
            time.sleep(60)
            
            current_size = os.stat(files_path + '\caderno3-Judicial-1Instancia-Capital' + '.txt').st_size
        
        pyautogui.press('enter')
        
        pyautogui.press('enter')
    
        time.sleep(5)
        
        pyautogui.keyDown('alt')   
        
        pyautogui.press('f4') 
        
        pyautogui.keyUp('alt')
 
        time.sleep(1)
        
        # Finish the process after conversion
        os.system("taskkill /f /im " + "AcroRd32.exe")

 
    else: 
        
        pdf_files_path = []
    
        
        for file in glob.glob(os.getcwd() +'\data' + file_name):
    
            pdf_files_path.append(file)    
        
        print('Amount of pdf files to be converted: ', len(pdf_files_path))
        print(pdf_files_path)
       
        for file in pdf_files_path:
            
            subprocess.Popen("%s" % (adobePath))

            time.sleep(5)
                
            pyautogui.keyDown('alt')

            time.sleep(5)            
            
            pyautogui.press('q')
            
            time.sleep(5)
            
            pyautogui.press('a')
            
            pyautogui.keyUp('alt')
            
            time.sleep(5)
            
            pyautogui.write(file)

            time.sleep(5)
            
            pyautogui.keyDown('alt')
            
            pyautogui.press('a')
            
            pyautogui.keyUp('alt')

            time.sleep(5)
            
            pyautogui.keyDown('alt')
    
            time.sleep(5)
    
            pyautogui.press('q')
    
            time.sleep(5)
    
            pyautogui.press('v')
        
            time.sleep(5)
        
            pyautogui.keyUp('alt')
        
            time.sleep(5)
            
            pyautogui.press('enter')            

            time.sleep(12)

            os.system("taskkill /f /im " + "AcroRd32.exe")
    
    execution_control[2] = 1
     
    return 1
 
# 4. Process files
def get_cases_from_txt(file_name):

    cases = []
    cases_details = []
    
    time.sleep(5)
    
    # Reading the txt file into a list
    file_reader = open(files_path + file_name + '.txt', mode='r', encoding='utf-8', errors='ignore')

    # Cleaning before Reading strategy!
    for line in file_reader:
            
        if line.rstrip('\n') != "": # Remove empty indexes from the list
            
            full_text.append(line.rstrip("\n")) # Remove new lines
       
# The same thing below, but with different syntax
    
    for index, item in enumerate(full_text):
        
        if item.find('CLASSE :BUSCA E APREENSÃO EM ALIENAÇÃO FIDUCIÁRIA') != -1 and len(full_text[index - 1].replace('PROCESSO :',"")) == 26 and (full_text[index + 1].replace('REQTE :',"").replace('REQTE ',"").count('.')) == 0:
            
            if str(full_text[index + 3]).find('REQTE') == 0:
                
                req = ""
                
            else:
                
                req = full_text[index + 3].replace('REQDO :',"").replace('REQDA :',"").replace('REQDA ',"").replace('REQDO ',"")
                
            cases.append(full_text[index - 1].replace('PROCESSO :',""))        
            cases_details.append({'REQTE': full_text[index + 1].replace('REQTE :',"") , 
                                  'REQDO':  req if len(full_text[index + 3]) > 5 else ""
                                 })

    print("Amount of cases to be searched:", len(cases))
    
    execution_control[3] = 1

    return cases, dict(zip(cases, cases_details))

# 5. Check the process on web
# Potential Exception: SessionNotCreatedException:
def check_process(cases):

    cases_file_map = []  
    
    try: 
        chrome_driver_path = os.path.dirname(os.path.abspath(__file__)) + '\chromedriver_84'
        browser = webdriver.Chrome(executable_path = chrome_driver_path) 
        browser.maximize_window()
        browser.get(esaj_consulta_processos)
    
        root_window = browser.window_handles[0]
    
        time.sleep(time_control)   
        
        browser.find_element_by_id('usernameForm').send_keys(key)
    
        time.sleep(time_control)   
    
        browser.find_element_by_id('passwordForm').send_keys(raisepass)
        
        time.sleep(time_control)   
      
        submit = browser.find_element_by_id('pbEntrar') 
        submit.click()
      
        time.sleep(time_control)
        
        search_process = browser.find_element_by_xpath('//*[@id="esajConteudoHome"]/table[2]/tbody/tr/td[2]/a')
        search_process.click()
    
        time.sleep(time_control)
        
        for case in cases:
            
            process_number_p1 = case.replace('.','').replace('-','')[0:13]
            process_number_p2 = case.replace('.','').replace('-','')[-4:]
        
            search_process_1level = browser.find_element_by_xpath('//*[@id="esajConteudoHome"]/table[1]/tbody/tr/td[2]/a')
            search_process_1level.click()
        
            time.sleep(time_control)    
        
            browser.find_element_by_id('numeroDigitoAnoUnificado').send_keys(process_number_p1)
        
            time.sleep(time_control)    
        
            browser.find_element_by_id('foroNumeroUnificado').send_keys(process_number_p2)
        
            time.sleep(time_control)    
        
            submit = browser.find_element_by_id('pbEnviar') # download for full download (website not working properly)
            submit.click()
        
            time.sleep(time_control)    
            
            access_elem = browser.find_element_by_xpath('//*[@id="linkPasta"]')
            access_elem.click()
        
            time.sleep(time_control)    
        
            tab_window = browser.window_handles[1]
            browser.switch_to_window(tab_window)
            browser.maximize_window()
        
            time.sleep(time_control)   
        
            checkbox = browser.find_element_by_xpath('//*[@id="pagina_1_cont_0_anchor"]')
            checkbox.click()
             
            time.sleep(time_control) 
            
            pyautogui.press('enter')
        
            time.sleep(time_control) 
            
            width = int((pyautogui.size()[0]) / 2)
            height = int((pyautogui.size()[1]) / 2) 
                    
            pyautogui.moveTo(width,height)
            pyautogui.click()
        
            time.sleep(time_control) 
            
            # Download button
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
            pyautogui.press('tab')
    
            time.sleep(time_control) 
            
            pyautogui.press('enter')
        
            time.sleep(6) 
        
            pyautogui.keyDown('alt')
            pyautogui.press('f4')
            pyautogui.keyUp('alt')
            
            browser.switch_to_window(root_window)
            
            time.sleep(time_control)    
            
            back_search = browser.find_element_by_xpath('/html/body/div/table[2]/tbody/tr/td[2]/table/tbody/tr[3]/td/a[3]')
            back_search.click()
            
            time.sleep(7)
            
            # Save to build a common df            
            for name in glob.glob(files_path + '\doc_*.pdf'):
                
                filename = os.path.basename(name)
                cases_file_map.append(filename[-12:-4])
                shutil.move(name , os.getcwd() + '\data')
                            
    except:

        pass
    
    browser.quit()
    
    execution_control[4] = 1
    
    print('Amount cases:', len(cases))
    print('Amount petitions:', len(cases_file_map))
    
    return list(zip(cases, cases_file_map))

# 6. Convert initial petition to txt and detect informations 
def extract_useful_info():
          
    files = []
    vehicle_roi = []    
    modelo = []
    chassi = []
    placa = []
    renavam = []
    word_bank = ['modelo', 'chassi', 'renavam', 'placa']
    
    # 1. Get all files to be processed
    for file in glob.glob(os.getcwd() + '\data' + '\doc_*.txt'):
    
        file_reader = open(file, mode='r', encoding='utf-8', errors='ignore') 
        file_content = file_reader.readlines()               
        file_reader.close()
        
        files.append(file[-12:-4])
                
        # 2. Natural Language Processing >> ROI strategy: Search the region of interest for vehicle details
        for statement in file_content:
                   
            for matcher in word_bank:
                
                if matcher in statement.lower():
                                        
                    # Key: Document + Elegible text                    
                    vehicle_roi.append([file[-12:-4], statement.strip('\n')])


        for statement in file_content:
            
            if 'modelo' in statement.lower():
                
                modelo.append([file[-12:-4], statement.lower().strip('\n').replace("modelo","").replace(":","").replace("marca","").replace('/'," ")])

            if 'chassi' in statement.lower():
                
                chassi.append([file[-12:-4], statement.lower().strip('\n').replace("chassi","").replace(":","")])

            if 'renavam' in statement.lower():
                
                text_found = statement.lower().strip('\n').replace("renavam","").rstrip().replace(":","")
                
                if text_found.isdigit():
                    
                    renavam.append([file[-12:-4], text_found])

            if 'placa' in statement.lower():
                
                placa.append([file[-12:-4], statement.lower().strip('\n').replace("placa","").replace(":","")])
                    
    #vehicle_information = []
    #vehicle_information = extract_text(files, vehicle_roi)

    execution_control[5] = 1
    
    return files, modelo, chassi, renavam, placa

# 7. Natural Language Processing

def extract_text(files, textList):

    # Controlers >> For each file to be processed, extract the following words from the textList (ROIs)   
    idx_list = 0
    idx_file = 0

    # Words to be extracted
    marca = ""
    modelo = ""
    chassi = ""
    renavam = ""
    placa = ""
    
    vehicle_information = []
   
    while idx_file < len(files):
    
        while idx_list < len(textList):
            
            if files[idx_file] == textList[idx_list][0]:
                
                text = textList[idx_list][1].lower().replace('.', "").replace(":","").split(" ")
                
                for i, elem in enumerate(text):
                    
                    if elem == 'marca':
                        
                        marca = text[i + 1]
                        
                        break
                        
                    if elem == 'modelo':
                        
                        modelo = text[i + 1]
                        
                        break
    
                    if elem == 'chassi':
                        
                        chassi = text[i + 1]
                        
                        break
                    
                    if elem == 'renavam':
                        
                        if str(text[i + 1]).isdigit():
                            
                            renavam = text[i + 1]
    
                        else:
    
                            pass
                        
                        break
                        
                    if elem == 'placa':
                        
                        placa = text[i + 1]                    
                        
                        break
    
            idx_list += 1
       
        vehicle_information.append({'FILE': files[idx_file],
                                'MARCA': marca,
                                'MODELO': modelo,
                                'CHASSI': chassi,
                                'RENAVAM': renavam,
                                'PLACA': placa
                                 })                      
        idx_file += 1
        idx_list = 0
    
    return vehicle_information    

def get_info(key, cases_files_map):

    idx_elem = 0
    
    while idx_elem < len(cases_files_map):        

        if key == cases_files_map[idx_elem][0]:
            
            file_name = (cases_files_map[idx_elem][1])
    
            return file_name            

            break
                
        idx_elem += 1
    
    return ' '


# 8. Build final letter
def letter_builder(process_info, cases_files_map, models, chassi, car_id, plate):
     
    for key, item in process_info.items():
        
        fname = get_info(key, cases_files_map)
        
        idx = 0
        target = ['|PROCESSO|','|BANCO|','|NOME|','|ENDERECO|','|VEICULO|']
            
        os.startfile(robot_path + '\RoboLaw\model.docx')
        
        time.sleep(10)
    
        pyautogui.keyDown('ctrl')
        
        pyautogui.press('u')
        
        pyautogui.keyUp('ctrl')
    
        while idx < len(target):
    
            time.sleep(3)
        
            pyautogui.write(target[idx])
        
            time.sleep(5)
        
            pyautogui.press('tab')
        
            time.sleep(5)
            
            if idx == 0:
                
                pyautogui.write(key)
                
            elif idx == 1:

                pyautogui.write(list(item.values())[0])
                
            elif idx == 2: 
                
                pyautogui.write(list(item.values())[1])
                
            elif idx == 3:
            
                pyautogui.write('EM DESENVOLVIMENTO')
                
            elif idx == 4:

                file_model = get_info(fname, models)
                file_chassi = get_info(fname, chassi)
                file_carid = get_info(fname, car_id)
                file_plate = get_info(fname, plate)

                text = 'Modelo-Marca:' + str(file_model) + ' Chassi:' + str(file_chassi) + ' Renavam:' + str(file_carid) + ' Placa:' + str(file_plate)
                pyautogui.write(text)
        
            time.sleep(5)
        
            pyautogui.keyDown('alt')
        
            pyautogui.press('i')
        
            pyautogui.keyUp('alt')
        
            time.sleep(5)        

            pyautogui.press('enter')
        
            idx += 1
    
        pyautogui.press('esc')
    
        time.sleep(5)
    
        pyautogui.press('f12') # Salvar como
    
        time.sleep(5)

        pyautogui.write(robot_path + '\RoboLaw\letters\\'+ 'Carta_'+ key + "_" + fname) # Substituir aqui para o numero do processo com loop

        time.sleep(5)
    
        pyautogui.press('enter')

        time.sleep(5)
    
        pyautogui.press('enter')
    
        time.sleep(5)
        
        pyautogui.keyDown('alt')   
        
        pyautogui.press('f4') 
        
        pyautogui.keyUp('alt')

        time.sleep(5)
    
#        os.system("taskkill /f /im " + "WINWORD.EXE") # passar para alt + f4

def attempt_controller():
    
    print('Under Construction')
    
    
def extract_address(cases_info):
    
    print('Under Construction')
    
def clean_environment():
    
    folder_name = str(date.today())
    
    # Check if the folder already exists. Otherwise create a current date folder
    if os.path.exists(robot_path + "\RoboLaw\hist\\" + folder_name) == False:
            
        os.mkdir(robot_path + "\RoboLaw\hist\\" + folder_name)

    # Move Caderno .txt and remove .pdf
    for name in glob.glob(files_path + '\caderno3-Judicial*.txt'):
        
        if os.path.exists(name):

            filename = os.path.basename(name)
            shutil.move(name , robot_path + "\RoboLaw\hist\\" + folder_name)

    for name in glob.glob(files_path + '\caderno3-Judicial*.pdf'):
        
        if os.path.exists(name):

            filename = os.path.basename(name)
            os.remove(name)
         
    
    # Move .pdf files for hist
    for name in glob.glob(robot_path + '\RoboLaw\data' + '\doc_*.pdf'):
        
        if os.path.exists(name):

            filename = os.path.basename(name)
            shutil.move(name , robot_path + "\RoboLaw\hist\\" + folder_name)


    # Delete .txt files 
    for name in glob.glob(robot_path + '\RoboLaw\data' + '\doc_*.txt'):
                        
        if os.path.exists(name):

            filename = os.path.basename(name)
            os.remove(name)


if __name__ == "__main__":
    
    print("Loading...")

#    ctypes.windll.user32.MessageBoxW(0, "Hello! =) I am a simple box", "Robo", 1)
           
#    setup_environment()   
    collect_data()
    process_pdf_to_txt(file_name= '\caderno3-Judicial-1ªInstancia-Capital', file_type= 1) 
    cases_list, cases_info = get_cases_from_txt(file_name= '\caderno3-Judicial-1Instancia-Capital')            
    cases_files_map = check_process(cases_list)                
    process_pdf_to_txt(file_name= '\doc_*', file_type= 0)   
    processed_files, models, chassi, car_id, plate = extract_useful_info() # NLP technique 
    address_details = extract_address(cases_info) # NLP technique    
    print('Cases Info length: ', len(cases_info)) # Let's check if the lenght is really the same
    print('Mapping Cases x File length:', len(cases_files_map)) # Let's check if the lenght is really the same
    letter_builder(cases_info, cases_files_map)
    clean_environment()

# THIS SOFTWARE IS AT ALPHA VERSION -- BUG FIXES AND TESTING ARE NEEDED EXTENSIVELY   
# cases_list >> List of all process individually
# cases_info >> Dictionary with process + individuals details (REQTE + REQDO)
# cases_files_map >> Dictionary with process + petition index files. Maps files to processes as index


# ON CODE VERSION CONTROL:
# v.1.2.1 >> Minor improvement: return map cases as List and not dictionary
# v.1.2.5 >> Transform to PDF Critical Update: Check each minute (60s) if the big file transformation is completed
# v.1.2.5 >> Add Vehicles information to be printed
# v.1.2.5 >> Improvements on software stability for 5s 
# v.1.2.5 >> Added the petition file name at the end of the letter
# v.1.2.5 >> Added new get_info() function. This function runs all over the dictionary looking for specific information given the petition file name