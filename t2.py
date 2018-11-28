# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from itertools import islice
import sys
import pandas as pd


def GetInOutDataFrame(path):
  wb = load_workbook(filename=path, read_only=True)
  ws = wb['MOUVEMENTS']

  expected_initial_header = 'Nom'
  header_row_index = 3
  row_index = header_row_index
  if ws.cell(row_index, 1).value != expected_initial_header:
    raise NameError('La valeur attendue de la cellule ' + ws.cell(row_index, 1).coordinate + ' est "' + expected_initial_header + '".')
  while ws.cell(row_index, 1).value:
    row_index = row_index + 1

  column_limit = 'Observations'
  column_max = 20
  column_index = 1
  columns = []
  while ws.cell(header_row_index, column_index).value != column_limit and column_index < column_max:
    columns.append(ws.cell(header_row_index, column_index).value)
    column_index = column_index + 1
  columns.append(ws.cell(header_row_index, column_index).value)

  if column_index == column_max:
    raise NameError('Il n\'apparait pas une colonne "' + column_limit + '" dans les 20 premières colonnes.')

  data = list(ws.values)[header_row_index:]
  return pd.DataFrame(data, columns=columns)


def main():
  #moves = GetInOutDataFrame(mouvements_path)
  #moves.to_csv('test.csv', index=False, encoding='utf-8')


if __name__ == '__main__':
  mouvements_path = '/home/thomas/Documents/Beta.gouv.fr/RH/2018-11/MOUVEMENTS DINSIC.xlsm'
  
  sys.exit(main())
