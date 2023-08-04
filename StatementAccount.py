from openpyxl import load_workbook

# Подключение вебхука с путем на файл
wb = load_workbook('выписка.xlsx')

# печатаем список листов
sheets = wb.sheetnames
for sheet in sheets:
    print(sheet)

sheet = wb.active

debet_sum = 0
return_sum = 0
index = 0
for i in sheet.values:
    index += 1
    if (i[1] != 'Итого обороты в валюте счета'):
        # Дебет
        if (isinstance(i[2], float) or isinstance(i[2], int)):
            if (index >= 14):
                if ('Возврат' in i[8]):
                    return_sum+=i[2]
        # Кредит
        if (isinstance(i[3], float) or isinstance(i[3], int)):
            if (index >= 14):
                print(index)
                print(i[3])
                if ('Возврат' in i[8]):
                    print('Возврат')
                else:
                    debet_sum += i[3]
    else:
        break


deduction_returns = debet_sum - return_sum
three_percent_tax = deduction_returns * 0.03

debet_sum = '{0:,}'.format(debet_sum).replace(',', ' ')
return_sum = '{0:,}'.format(return_sum).replace(',', ' ')
deduction_returns = '{0:,}'.format(deduction_returns).replace(',', ' ')
three_percent_tax = '{0:,}'.format(three_percent_tax).replace(',', ' ')

print(f'Итого оборота: {debet_sum}тг')
print(f'Возвраты на сумму: {return_sum}тг')
print(f'Обороты с вычтенным возвратом: {deduction_returns}тг')
print(f'Чистые 3%: {three_percent_tax}тг')
