# turnover_tax_kaspi_kz
Расчет налогов ИП 3% с вычетом возвратов. Работает с выпиской из каспи.

Очень простая программа которая легко упрощает работу консалтинговой компании. В период сдачи 910 формы, приходится в ручную считать чистый оборот своего каспи магазина.

Для активации программы нужно установить дополнительную библиотеку openpyxl
`
pip install openpyxl
`
На данный момент название документа хранится в переменной 

wb = load_workbook('выписка.xlsx')

Нужно сменить название выписка.xlsx на свой документ.
