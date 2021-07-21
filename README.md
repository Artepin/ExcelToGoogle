# ExcelToGoogle
Разработанная программа работает на python 3.8
В написанном коде используются следующие библиотеки сторонние:
1.	gspread
2.	datetime
3.	re
4.	google-api-python-client
5.	oauth2client
6.	numpy
А также библиотеки, написанные или доработанные командой разработчиков:
1.	spreadsheetgoogle
2.	exellib
Внешние подключаемые модули:
1.	deptControl
2.	columnRevise
Модуль deptControl отвечает за контроль, перенос и группировку из одной таблицы в 3 задолженностей по заказам с фильтром по фамилии работника. Группировка происходит в 3 таблицы. Одна из них «красная», т.е. срок задолженности либо сегодня, либо уже прошёл. Вторая – «желтая» с подходящим списком и датой по плану не более 14 дней с текущего момента по планируемую дату завершения. Наконец, 3 таблица с уже выполненными заказами – при поступлении даты в листах с «красными» задолженностями в поле фактическое выполнение строка переносится в таблицу «выполнено».
Данный модуль реализует следующие функции:
1.	deptControl – авторизует пользователя и понимает, какую таблицу мы будем сейчас редактировать.
2.	dateTransform – преобразует дату из документа google в дату для последующей обработки в программе.
3.	changeOfColor –меняет цвет незаполненной колонки фактической даты.
4.	isItLate- проверяет давность даты
5.	prohod – главный цикл обхода колонки в документе. В зависимости от соответствия дат в колонках «дата план»  и «дата факт» вызывает функции выше.
…(надеюсь Артем это написал)
Модуль columnRevise отвечает за перенос нужных значений в диапазоне под нужным столбцом(указывается имя столбца, например «Длина») в Гугл таблицу в столбец с тем же названием, причем с проверкой условий. На данный момент перенос происходит, если значение из excel больше, в случае, когда на этом месте в Гугл таблице нет значения, то в него записывается значения из excel документа, если оно есть. При наличии нескольких столбцов с одним названием в любой из таблиц, пользователю будет предложен выбор в какой / из какого столбца производить перенос.
Библиотека exellib внутри себя имеет класс Exellib, и его экземпляр вида “el”. Основные функции библиотеки:
1.	redFile, на вход получает путь к excel файлу. Результат: данный файл готов к работе.
2.	sheetID, на вход получает номер нужного листа. Результат: данный лист готов к работе.
3.	getRows, возвращает максимальное количество строк в листе.
4.	getColumns, возвращает максимальное количество колонок в листе.
5.	getNumber, на вход получает координату ячейки, возвращает ее значение.
6.	columnLetter, на вход получает численный номер столбца, возвращает его буквенный эквивалент.
7.	getMerged, возвращает список с диапазонами всех объединённых ячеек.
