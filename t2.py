# -*- coding: utf-8 -*-

from openpyxl import load_workbook
from itertools import islice
import sys
import pandas as pd


start = '2018-01-01'
end = '2018-12-31'


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
  raw_df = pd.DataFrame(data, columns=columns)

  columns = ['Nom', u'Prénom', u'Date d\'entrée', 'Date de sortie', 'Observations']
  new_columns = ['NOM', 'PRENOM', 'ENTREE', 'SORTIE', 'ATTRIBUTION']
  df = raw_df[columns]

  df.columns = new_columns

  df = df[-df.ENTREE.isnull()].copy()

  for col in df.columns:
    if df[col].dtype == 'object':
      df[col] = df[col].str.strip()

  df.loc[df.SORTIE.isnull(), 'SORTIE'] = end
  df['MATCH'] = 1

  df.ENTREE = pd.to_datetime(df.ENTREE)
  df.SORTIE = pd.to_datetime(df.SORTIE)

  return df


def getDailyMoves(df):
  days = pd.date_range(start=start, end=end, freq='B')
  day_df = pd.DataFrame({ 'MATCH': 1, 'JOUR':days})

  daily_df = df.merge(day_df, on='MATCH')
  idx = (daily_df.ENTREE <= daily_df.JOUR) & (daily_df.JOUR <= daily_df.SORTIE)
  filterd_df = daily_df[idx].copy()

  filterd_df.loc[:, 'PERIODE'] = filterd_df.JOUR.dt.month
  return filterd_df.set_index(['NOM', 'PRENOM', 'PERIODE'])


def getT2(path):
  wb = load_workbook(filename=path, read_only=True)
  ws = wb['Feuil1']
  data = ws.values
  columns = next(data)
  data = list(data)
  raw_df = pd.DataFrame(data, columns=columns)

  period_columns = [i for i in range(13) if i in raw_df.columns]
  if not len(period_columns):
    raise KeyError('Aucune période n\'est présente dans le fichier.', raw_df.columns)

  required_columns = ['NOM', 'PRENOM']
  preserved_columns = required_columns + period_columns

  missing = [required for required in required_columns if required not in preserved_columns]
  if len(missing):
    raise KeyError('Des champs obligatoires sont manquants : ' + missing)

  limited_df = raw_df[preserved_columns]
  for col in limited_df.columns:
    if limited_df[col].dtype == 'object':
      limited_df.loc[:,col] = limited_df[col].str.strip()

  long_df = limited_df.melt(id_vars=required_columns, var_name='PERIODE', value_name='MONTANT')
  print(long_df.shape)
  long_df = (long_df[-long_df.MONTANT.isnull()]).copy()

  print(long_df.shape)
  return long_df.groupby([] + required_columns + ['PERIODE']).sum()


def getAttribution(moves, t2):
  distribution = moves.groupby(moves.index.names).ATTRIBUTION.count()

  t2_daily = pd.concat([t2, distribution], axis=1)
  t2_daily['TJ'] = t2_daily.MONTANT / t2_daily.ATTRIBUTION

  moves_montant = pd.merge(moves.reset_index(), t2_daily.TJ.reset_index(), how='outer')

  return moves_montant


def main():
  mouvements_path = '/home/thomas/Documents/Beta.gouv.fr/RH/2018-11/MOUVEMENTS DINSIC.xlsm'
  moves = GetInOutDataFrame(mouvements_path)
  daily_moves = getDailyMoves(moves)
  daily_moves.to_csv('moves-test.csv', encoding='utf-8', index=False)

  p = '/home/thomas/Documents/Beta.gouv.fr/RH/2018-11/T2.xlsx'
  df = getT2(p)
  df.to_csv('t2-test.csv', encoding='utf-8')

  attrib = getAttribution(daily_moves, df)
  attrib.to_csv('attrib-test.csv', encoding='utf-8')

  montants = attrib.groupby('ATTRIBUTION').TJ.sum()
  montants.to_csv('montants-test.csv', encoding='utf-8')

  control_groups_df = attrib.groupby(['NOM', 'PRENOM'])
  control_df = pd.DataFrame({ 'MONTANT': control_groups_df.TJ.sum(), 'JOURS': control_groups_df.ENTREE.count() })
  control_df.to_csv('control-full-test.csv', encoding='utf-8')

  control_df[(control_df.MONTANT == 0) | (control_df.JOURS == 0)].to_csv('control-test.csv', encoding='utf-8')

if __name__ == '__main__':
  sys.exit(main())
