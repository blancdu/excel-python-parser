import os
import re
import string
import sqlite3
import logging
import pandas as pd
import sqlalchemy as sa
import multiprocessing

if os.path.isfile('log.log'):
    os.remove('log.log')

logging.basicConfig(filename="log.log",
                    level=logging.WARNING, filemode='a', format='%(message)s')


def createFolder(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print('Error: Creating directory. ' + directory)


def excel_2_sql(filename: str, sheet_idx:int=0) -> None:
    xlsx = pd.ExcelFile(filename)
    sheet_name = xlsx.sheet_names[sheet_idx]
    df = pd.read_excel(filename, engine='openpyxl', sheet_name=sheet_name, usecols=[0, 1, 2, 3])
    engine = sa.create_engine(
        'sqlite:///db.db', echo=False, use_insertmanyvalues=True)
    sqlite_connection = engine.connect()
    filelist = os.listdir('src')
    for filename in filelist:
        if filename[-5:] != '.xlsx':
            continue
        table_name = filename[:-5]
        partial_df = df[df.iloc[:, 0] == table_name].drop(
            df.columns[0], axis=1)
        coltype = {'Description': sa.types.String,
                   'Source_CN': sa.types.String, 'Target_KR': sa.types.String}
        partial_df.to_sql(name=table_name, con=sqlite_connection, dtype=coltype,
                          if_exists='replace', index=False, index_label='Description')
    sqlite_connection.close()


def translate(filename: str) -> bool:
    def coord_2_idx(coord: str) -> tuple[int]:
        if not coord:
            return None
        col, row, *_ = re.split('(\d+)', coord)
        num = 0
        for c in col:
            if c in string.ascii_letters:
                num = num * 26 + (ord(c.upper()) - ord('A')) + 1
        col = num - 1
        row = int(row) - 1
        return (row, col)

    def validation(table_name: str, row: list, df: pd.DataFrame):
        def advanced_strip(s):
            s = s.strip()
            if len(s) >= 7 and s[-7:].lower() == '_x000d_':
                s = s[:-7]
            return s

        coord, cn, kr = row
        r, c = coord_2_idx(coord)
        c = c*2
        try:
            target_value = df.iat[r, c]
            if target_value is pd.NA:
                target_value = None
        except IndexError:
            logging.warning('coordinate is out of excel sheet')
            logging.warning(f'[{table_name}]-통합시트: {row}')
            return False

        # advanced strip
        if isinstance(target_value, str):
            target_value = advanced_strip(target_value)
        if isinstance(cn, str):
            cn = advanced_strip(cn)

        # if cell value is Null
        if not target_value:
            target_value = ''
        if not cn:
            cn = ''
        if not kr:
            return False
            kr = ''
        if not all([target_value, cn, kr]):
            if not any([target_value, cn, kr]):
                return False
            logging.warning('missing cell value')
            logging.warning(f'[{table_name}]-통합시트{coord} "{cn}" "{kr}"')
            logging.warning(f'[{table_name}.xlsx]{coord}: "{target_value}"')
            return False

        # mismatch string
        if target_value != cn:
            logging.warning('mismatch')
            logging.warning(f'[{table_name}]-통합시트{coord} "{cn}" "{kr}"')
            logging.warning(f'[{table_name}.xlsx]{coord}: "{target_value}"')
            return False

        return True

    xlsx = pd.ExcelFile('src/'+filename)
    sheet_name = xlsx.sheet_names[0]
    df = pd.read_excel(xlsx, engine='openpyxl',
                       header=None, dtype='string', sheet_name=sheet_name)

    # dst columns duplication
    column_size = len(df.columns)
    for i in range(column_size):
        df.insert(i*2+1, (i*2+1)/2, "")
    con = sqlite3.connect('db.db')
    cur = con.cursor()

    table_name = filename[:-5]
    rows = cur.execute(f'SELECT * FROM {table_name}').fetchall()
    cur.close()

    for row in rows:
        coord, _, kr = row
        r, c = coord_2_idx(coord)
        c = c*2+1
        # validation is optional func
        if not validation(table_name, row, df):
            continue
        df.iat[r, c] = kr
    df.to_excel(excel_writer='dst/'+filename, sheet_name=sheet_name,
                header=False, index=False, engine='openpyxl')
    return True


def omission_check(filename: str) -> bool:
    # idx2coord function
    # if cell is not na
    # db tablename:filename coord:idx2coord cn:cell value check
    # if there is not in db then logging
    omission_cnt = 0
    def col_index_to_letter(column_index:int):
        letter = ''
        while column_index > 25:
            letter += chr(65 + int(column_index/26) -1)
            column_index -= int(column_index/26)*26
        letter += chr(65 + int(column_index))
        return letter

    xlsx = pd.ExcelFile('src/'+filename)
    sheet_name = xlsx.sheet_names[0]
    df = pd.read_excel(xlsx, engine='openpyxl',
                       header=None, dtype='string', sheet_name=sheet_name)

    con = sqlite3.connect('db.db')
    cur = con.cursor()
    table_name = filename[:-5]

    df = df.fillna('')
    for c, s in df.items():
        df[c] = s.str.removesuffix('_x000d_')
        for r, v in s.items():
            if v:
                coord = col_index_to_letter(c)+str(r+1)
                cn = df[c][r]
                rows = cur.execute(f'SELECT * FROM {table_name} WHERE coord = "{coord}"').fetchall()
                if len(rows) == 1:
                    _, dbcn, _ = rows[0]
                    if dbcn != cn:
                        omission_cnt += 1
                        logging.warning(f'mismatch: [{table_name}:{coord}] {dbcn},{cn}')
                else:
                    omission_cnt += 1
                    logging.warning("rows length is not 1")
    cur.close()
    return omission_cnt


if __name__ == '__main__':
    translate_mode = True
    createFolder('src')
    createFolder('dst')
    if os.path.isfile('db.db'):
        print('기존 db 발견')
    else:    
        print('db 생성 시작')
        excel_2_sql('main.xlsx')
        print('db 생성 완료')
    file_list = [x for x in os.listdir('./src/') if x[-5:] == '.xlsx']
    if translate_mode:
        print('변환 시작')
        result = multiprocessing.Pool().map(translate, file_list)
        if all(result):
            print('변환 완료')
        else:
            print('일부 변환 완료')
            print(result)
    else:
        print('누락 검출 시작')
        result = multiprocessing.Pool().map(omission_check(), file_list)
        omission_file_list = []
        for file, res in zip(file_list, result):
            if res:
                omission_file_list.append((file, res))
        for file, res in omission_file_list:
            print(f'{file} has {res} missing values')