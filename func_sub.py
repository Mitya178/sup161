import os
import pandas as pd
from func_io import print_cmd, sys_exit, get_key_list, get_key


# Константы наименований колонок в IO_LIST
TAG_COLUMN = 'TAG'
DECIMAL_DIGITS_COLUMN = 'DEC_DIG'


# Объвление функции подсчета количества строчек/колонок в таблице Excel
def get_dict_row_col(path, book, sheet, rows, column_enable=False,
                     row_enable=True):
    if sheet in book.sheet_names:
        df = pd.read_excel(os.path.join(path, book),
                           sheet_name=sheet, dtype=object, engine='openpyxl')
        DICT = df.to_dict()
    else:
        print_cmd("Лист " + sheet + " отсутствует в таблице " + book + "!\n")
        DICT = sys_exit()
    LB_ROW = len(DICT[rows])
    LB_COLUMN = len(DICT)
    print_cmd(DICT, flag_print=False)
    print_cmd('', flag_print=False)
    if column_enable and row_enable:
        res = DICT, LB_ROW, LB_COLUMN
    elif row_enable:
        res = DICT, LB_ROW
    else:
        res = DICT
    return res


# Объявление функции получения значений параметров из IO_list
def get_parameters(loop_id: int, loop: dict, swtype_rules: dict,
                   io_list: dict, calc: dict, parameters: dict) -> dict:
    key_sw = get_key_list(swtype_rules['SW_TYPE'], loop['SW_TYPE'][loop_id])
    for i in key_sw:
        io_name = loop[swtype_rules['IO_NAME'][i]][loop_id]
        io_param = swtype_rules['IO_PARAM'][i]
        block_name = loop[swtype_rules['BLOCK_NAME'][i]][loop_id]
        block_param = swtype_rules['BLOCK_PARAM'][i]
        tag_key = get_key(io_list[TAG_COLUMN], io_name)
        param_value = []
        # параметр является CALC функцией
        if io_param in calc['FUNCTION'].values():
            # номер строки параметра в листе CALC
            calc_key = get_key(calc['FUNCTION'], io_param)
            # количество знаков после запятой
            digits_decimal = io_list[DECIMAL_DIGITS_COLUMN][tag_key]
            formul = calc['FORMUL'][calc_key]
            for k in range(len(formul)):
                # символ формулы (столбец или знак)
                symbol = formul[k]
                if symbol in calc.keys():
                    # параметр IO листа из столбца CALC
                    parametr = calc[formul[k]][calc_key]
                    if parametr in io_list.keys():
                        param_value.append(io_list[parametr][tag_key])
                    elif pd.isnull(parametr):
                        print_cmd(f"Ошибка при формировании функции {io_param}"
                                  " в листе CALC!\n")
                        sys_exit()
                    else:
                        param_value.append(parametr)
                # знаки в формуле
                else:
                    param_value.append(formul[k])
            param_value = ''.join(map(str, param_value))
            # формула не "OFF"/"ON"
            if not param_value.isalpha():
                # param_value = round(float(ne.evaluate(param_value)), digits_decimal)
                param_value = round(float(eval(param_value)), digits_decimal)
        # параметр найден в IO листе
        elif io_param in io_list.keys():
            if pd.isnull(io_list[io_param][tag_key]):
                param_value = '0'
            else:
                param_value = io_list[io_param][tag_key]
        else:
            print_cmd(f"Таблица 'IO_LIST_BGCC' недоступна или допущена ошибка "
                      "в формировании SW_TYPE_BLOCK_RULES шаблона\n")
            sys_exit()
        # запись значения параметра в param DICT (для замены параметров FB в csv)
        if block_name not in parameters:
            parameters[block_name] = {}
        parameters[block_name][block_param] = param_value
    return parameters


# Объявление функции подсчета количества параметров функционального блока
def parameter_counter(sheet, fb_name) -> int:
    fb_column = 'FB Tag Name'
    param_name_column = 'Param Name'
    key_fb = get_key(sheet[fb_column], fb_name)
    result = 0
    # если в листе присутствует 2+ функциональных блоков
    for i in range(key_fb + 1, len(sheet[fb_column])):
        if sheet[fb_column][i] != sheet[fb_column][i]:
            pass
        else:
            result = i - key_fb - 1
            break
    # если в листе один (последний) функциональный блок
    if result == 0:
        for i in range(key_fb, len(sheet[param_name_column])):
            if sheet[param_name_column][i] != sheet[param_name_column][i]:
                break
            else:
                result += 1
    return result
