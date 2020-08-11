import os

from pandas import read_excel


def filter_xlsx():  # 현재 폴더에 있는 엑셀 파일의 이름을 리스트에 담아서 리턴함
    result = list()
    for file in os.listdir(os.getcwd()):
        if file.endswith('.xlsx') and not file.startswith('~$'):
            result.append(file)
    return result


def filter_process_num(data):  # 공정번호가 30번인거만 남김
    is_data = data['공정번호'] == 30
    filter_data = data[is_data]
    return filter_data


# print(__file__)
# print(os.path.realpath(__file__))
# print(os.path.abspath(__file__))
# print(os.getcwd())
# print(os.listdir(os.getcwd()))
# todo clean code


def make_to_list():
    excel_file_list = filter_xlsx()  # 폴더 안에 있는 엑셀 파일 이름을 가져옴
    first_file = list()  # 첫번째 파일의 rows를 담을 리스트
    second_file = list()  # 두번째 파일의 rows를 담을 리스트
    column_name = list()  # 컬럼명을 담을 리스트
    for index, file in enumerate(excel_file_list):
        data = read_excel(f'{os.getcwd()}/{file}', sheet_name='Sheet1')  # 엑셀 파일 읽어옴
        data = filter_process_num(data)  # 공정번호 30번만 남김
        column_name = data.head(0)  # 컬럼명 세팅

        for row_num in range(len(data)):
            tmp_row_list = list()  # 한개의 행을 담을 리스트

            for i in data.loc[row_num]:
                tmp_row_list.append(i)
            if index == 0:  # 첫번째 파일 실행시
                first_file.append(tmp_row_list)
            else:  # 두번째 파일 실행시
                second_file.append(tmp_row_list)

    column_name = list(column_name.columns)  # 컬럼명을 리스트로 변환
    return first_file, second_file, column_name, excel_file_list


first_file, second_file, column_name, excel_file_name = make_to_list()  # 필요한 데이터들을 리스트로 불러옴
tmp_first_file = list(first_file)  # 데이터 복사
tmp_second_file = list(second_file)  # 데이터 복사

for i in first_file:
    if i in tmp_second_file:
        tmp_second_file.remove(i)  # 첫번째 파일과 다른것만 남김

for i in second_file:
    if i in tmp_first_file:
        tmp_first_file.remove(i)  # 두번째 파일과 다른것만 남김

with open('result.txt', mode='w') as f:  # 결과를 메모장으로 작성함
    f.write(f'-------------{excel_file_name[0]}-------------\n')
    for i in tmp_first_file:
        for index, value in enumerate(i):
            f.write(column_name[index] + ' : ' + str(value) + "\t")
        f.write('\n')

    f.write(f'-------------{excel_file_name[1]}-------------\n')
    for i in tmp_second_file:
        for index, value in enumerate(i):
            f.write(column_name[index] + ' : ' + str(value) + "\t")
        f.write('\n')
