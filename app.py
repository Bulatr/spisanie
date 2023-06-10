import pandas as pd
import PyPDF2
import random
from datetime import datetime, timedelta

reasons = {
    'Компьютер': [
        'Неисправность цепи питания, что приводит к неработоспособности всей платы.',
        'Повреждение разъема процессора, что делает невозможным установку процессора или вызывает ошибки при его работе',
        'Неисправность чипсета, ведущая к проблемам с распознаванием и управлением других компонентов на плате',
        'Повреждение слотов памяти, вызывающее ошибки при установке оперативной памяти или приводящее к неработоспособности системы.',
        'Неисправность графического разъема или чипа, что приводит к отсутствию вывода видеосигнала.',
        'Повреждение разъемов SATA или IDE, что приводит к неработоспособности жестких дисков или оптических приводов.',
        'Неисправность разъемов USB, ведущая к проблемам с подключением устройств.',
        'Неисправность BIOS-чипа, приводящая к невозможности загрузки системы и настройки платы.',
        'Повреждение микрокомпонентов или трасс на плате, что может вызывать различные ошибки и сбои в работе системы.',
        'Неисправность контроллера Ethernet, вызывающая проблемы с сетевым подключением.',
        'Неисправность разъема питания, приводящая к проблемам с подачей электропитания на плату.'
    ],
    'Ноутбук': [
        'Неисправность цепи питания, что приводит к неработоспособности всей платы.',
        'Неисправность BIOS-чипа, приводящая к невозможности загрузки системы и настройки платы.',
        'Повреждение разъемов SATA или IDE, что приводит к неработоспособности жестких дисков или оптических приводов.',
        'Неисправность контроллера SATA или RAID, приводящая к проблемам с обнаружением и работой жестких дисков или массивов данных.',
        'Неисправность контроллера Ethernet, вызывающая проблемы с сетевым подключением.',
        'Повреждение BIOS-флэш-памяти, что может привести к неработоспособности системы или ошибкам при загрузке.'
    ],
    'Принтер': [
        'Повреждение лазерного модуля',
        'Повреждение блока формирования изображения',
        'Неисправность платы управления.',
        'Неисправность платы форматирования.',
        'Механическое повреждение принтера. Прекращение поддержки и отсутствие запчастей'
    ],
    'Принтер цветной':[
        'Повреждение печатающей головки, неисправность системы подачи чернил'
    ],
    'СканерШК': [
        'Неисправность платы управления или процессора.'
    ],
    'Сканер': [
        'Неисправность платы управления или процессора.'
    ],
    'ИБП' : [
        'Неисправность платы управления. Износ аккумулятора'
    ],
    'АТС' : [
        'Неисправность чипа, приводящая к невозможности загрузки системы и настройки платы. Неисправность платы управления.'
    ],
    'Проектор' : [
        'Повреждение оптической системы или объектива. Неисправность матрицы или DLP-чипа.'
    ],
    'Монитор' : [
        'Повреждение ЖК-панели.',
        'Неисправность платы инвертора или платы подсветки. ',
        'Повреждение графического контроллера или платы графического вывода. '
    ]
}

def generate_date():
    # Задаем начальную и конечную дату
    start_date = datetime(2023, 1, 15)
    end_date = datetime(2023, 6, 15)

    # Создаем список выходных дней
    weekend_days = [5, 6]  # Предполагаем, что суббота (5) и воскресенье (6) - выходные дни

    generated_dates = []
    current_date = start_date

    while current_date <= end_date:
        rand_days = random.randint(3, 5)
        if current_date.weekday() not in weekend_days and current_date not in generated_dates:
            generated_dates.append(current_date)
        current_date += timedelta(rand_days)

    # Выводим даты в порядке возрастания
    generated_dates.sort()
    return generated_dates

def next_5_day(random_date):
    # Добавляем 5 дней, пропуская выходные дни
    weekend_days = [5, 6]
    days_added = 0
    while days_added < 5:
        random_date += timedelta(days=1)
        if random_date.weekday() not in weekend_days:
            days_added += 1
    return random_date

def fill_pdf_form(template_file, output_file, inventory_number, name, equipment_type, num, date:datetime):
    # Открытие существующего PDF-файла
    with open(template_file, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        # Копирование страниц из шаблонного файла в новый файл
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            writer.add_page(page)

        # Получение полей формы из первой страницы PDF-файла
        form_fields = reader.get_fields()
        # {'inventory_number': {}, 'name': {}, 'number': {}, 'malfunction_reason': {}, 'name1': {}, 'name2': {}}

        # генерация стоимости случайным порядком
        price = random.randrange(5800, 12301, 5)  # Генерация случайного числа в диапазоне от 5800 до 12300

        # Генерация процента износа с шагом 5% в диапазоне от 70% до 95%
        wear_percentage = random.randrange(70, 96, 5)

        # Заполнение полей формы данными
        name1 = name
        name2 = name
        price = f'{price} рублей'
        wear_percentage = f"{wear_percentage}%"
        next_5 = next_5_day(date)
        random_date = date.strftime('%d.%m.%Y')
        next_5 = next_5.strftime('%d.%m.%Y')
        form_fields['inventory_number'] = str(inventory_number)
        form_fields['name'] = str(name)
        form_fields['number'] = str(num)
        form_fields['name1'] = str(name1)
        form_fields['name2'] = str(name2)
        form_fields['price'] = str(price)
        form_fields['wear_percentage'] = str(wear_percentage)
        form_fields['random_date'] = str(random_date)
        form_fields['next_5_day'] = str(next_5)

        # Генерация случайной причины неисправности в зависимости от типа оборудования
        if equipment_type in reasons:
            random_reason = random.choice(reasons[equipment_type])
            form_fields['malfunction_reason'] = str(random_reason)

        # Запись заполненных полей формы в новый PDF-файл
        with open(output_file, 'wb') as output:
            writer.update_page_form_field_values(writer.pages[0], form_fields)
            writer.write(output)

# Загрузка Excel-файла с данными
filename = "spisok.xlsx"
df = pd.read_excel(filename)

# Загрузка шаблонного PDF-файла с полями формы
template_file = "шаблон2.pdf"

num = 25
date_index = 0
iteration_date = 1
rand_dates = generate_date()
# Перебор строк в Excel-файле
for index, row in df.iterrows():
    inventory_number = row['инвентарный номер']
    name = row['наименование']
    equipment_type = row['тип оборудования']
    num = num + 1
    # Генерация уникального имени файла для текущей строки
    output_file = f"spisanie_{num}.pdf"
    if iteration_date < random.randrange(2,5):
        iteration_date += 1
    else:
        date_index += 1
        iteration_date = 1
    date = rand_dates[date_index]
    # Заполнение полей формы в PDF-файле
    fill_pdf_form(template_file, output_file, inventory_number, name, equipment_type, str(num), date)