KONVERT: DOCX -> FB2 / EPUB
===========================

Программа создаёт FB2, EPUB или оба формата из DOCX. Перед конвертацией она
сохраняет нормализованную копию DOCX и преобразует текстовую разметку:

*текст*       -> курсив
**текст**     -> жирный
***текст***   -> жирный курсив
> текст       -> цитата первого уровня
>> текст      -> цитата второго уровня
> > текст     -> цитата второго уровня

Маркер цитаты распознаётся только в начале абзаца. Разметка может охватывать
несколько слов и может встречаться несколько раз в одном абзаце.

Файлы
-----

Пути и форматы задаются в config.ini. Все относительные пути отсчитываются от
папки, в которой расположен config.ini:

[book]
input = book.docx
metadata = metadata.txt
cover = cover.jpg
normalized = book.normalized.docx
output_dir = .

[output]
formats = fb2, epub

[normalization]
enabled = yes
report = conversion-report.txt

cover.jpg необязателен. Исходный DOCX никогда не перезаписывается. Результаты:

book.normalized.docx
book.fb2
book.epub
conversion-report.txt

Установка (PowerShell)
----------------------

python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt

Если PowerShell запрещает активацию:

Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1

Можно не активировать окружение:

.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe main.py config.ini

Запуск
------

python main.py config.ini

Также можно передать папку, в которой находится config.ini:

python main.py "D:\Книги\Моя книга"

metadata.txt
------------

title=Название книги
author=Имя Фамилия
genre=fantasy
lang=ru
sequence=Название цикла
sequence_number=1
annotation=Первый абзац||Второй абзац
date=2026

Для глав следует применять встроенные стили Word «Заголовок 1» —
«Заголовок 6». Обычное форматирование Word (жирный и курсив) также переносится.

Ограничения
-----------

Таблицы пропускаются с предупреждением. Абзацы с гиперссылками, формулами,
полями, разрывами или встроенными изображениями не перестраиваются
нормализатором, чтобы не повредить сложные объекты Word; предупреждение об этом
попадает в conversion-report.txt.

EPUB сохраняет вложенные цитаты как вложенные blockquote. FB2 не разрешает
вложенный cite по своей схеме, поэтому первый уровень сохраняется как cite, а
дополнительные уровни обозначаются оставшимися маркерами "> ".
