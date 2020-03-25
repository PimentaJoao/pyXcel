# -*- coding: utf-8 -*-

import xlrd
import os
import shutil
from termcolor import colored

os.system('')

def verifiesInputFolderExists():
      pastaInput = './INPUT'
      try:
            print '[SISTEMA]Tentando criar pasta de INPUT...'
            os.mkdir(pastaInput, 0755)
            return True
      except:
            print '[SISTEMA]Pasta INPUT ja existe.'
            return False
            
def verifiesOutputFolderExists():
      pastaOutput = './OUTPUT'
      try:
            print '[SISTEMA]Tentando criar pasta de OUTPUT...'
            os.mkdir(pastaOutput, 0755)
      except:
            print '[SISTEMA]Pasta OUTPUT ja existe.\n'
            
def collectsNewName(person_directory, person_file):
      file_location = './INPUT/' + person_directory + '/' + person_file[:-4]

      workbook = xlrd.open_workbook(file_location)

      sheet = workbook.sheet_by_index(0)

      info = sheet.cell_value(1, 2)

      return info 

# I/O
didICreateTheInputFolderJustNow = verifiesInputFolderExists()
verifiesOutputFolderExists()
if didICreateTheInputFolderJustNow == False:
      for person_directory in os.listdir('./INPUT'):
            try:
                  new_folder = './OUTPUT/' + person_directory
                  os.mkdir(new_folder, 0755)

                  for person_file in os.listdir('./INPUT/' + person_directory):
                        new_name = './OUTPUT/' + person_directory + '/' + person_file[:-5] + ' ' +  collectsNewName(person_directory, person_file + 'xlsx') + '.xlsx'
                        shutil.copy2('./INPUT/' + person_directory + '/' + person_file, new_name)
                  
                  print colored('[SUCESSO]: Pasta ' + person_directory + ' processada.', 'green')

            except:
                  print colored('[ALERTA]: A pasta ' + person_directory + ' ja existe em OUTPUT e foi ignorada.', 'yellow')

      print colored('\n[SISTEMA]: PROCESSO FINALIZADO, TODAS PASTAS PROCESSADAS.\n', 'green')
      raw_input('Pressione Enter para fechar...')
