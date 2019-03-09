import sys
import openpyxl


def r_chop(the_string, ending):
    if the_string.endswith(ending):
        return the_string[:-len(ending)]
    return the_string


def save_to_file(name, value):
    file = open(name, "w+")
    file.write(value)
    file.close()


def main(argv):
    excel_file = openpyxl.load_workbook(argv[1])
    sheet = excel_file.get_sheet_by_name('Export Worksheet')
    sql_query = 'with last_forecast as ('

    for row in range(2, sheet.max_row + 1):
        sql_query += ('\n select \'' + sheet['C' + str(row)].value + '\' item, ' + str(int(sheet['G' + str(row)].value)
                                                                            + int(sheet['H' + str(row)].value)
                                                                            + int(sheet['I' + str(row)].value)
                                                                            + int(sheet['J' + str(row)].value)
                                                                            + int(sheet['K' + str(row)].value)
                                                                            + int(sheet['L' + str(row)].value)
                                                                            + int(sheet['M' + str(row)].value)
                                                                            + int(sheet['N' + str(row)].value)
                                                                            + int(sheet['O' + str(row)].value)
                                                                            + int(sheet['P' + str(row)].value)
                                                                            + int(sheet['Q' + str(row)].value)
                                                                            + int(sheet['R' + str(row)].value)
                                                                            + int(sheet['S' + str(row)].value)
                                                                            )
              + ' qty from dual \n union all')

    sql_query = r_chop(sql_query, ' union all')
    sql_query += ')'

    if argv[2]:
        save_to_file(str(argv[2]).replace(" ", "_"), sql_query)
    else:
        save_to_file(str(argv[1]).replace(" ", "_").replace("xlsx", "txt"), sql_query)


if __name__ == "__main__":
    main(sys.argv)
