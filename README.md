# excel-2-tilda-shop-2-satu-marker
<b>Simple script to transform excel book to tilda import format and satu import format</b>

<p>Подготовительная работа:
<p>Скачать и установить python (скрипт создан для версии Python 3.11)
<ul>
  <li>- Открыть командную строку (терминал)</li>
  <li>- Перейти в директорию с сохраненным репозиторием.</li>
  <li>- Выполнить следующие команды:</li>
</ul>
<br>
<p><b>python -m venv myenv</b>       # для создания и запуска виртуальрого окружения</p>
<p><b>source myenv/bin/activate</b>  # для Linux и macOS</p>
<p><b>myenv\Scripts\activate.bat</b>     # для Windows</p>
<p><b>pip install -r requirements.txt</b> # Для установки дополнительных библиотек, используемых в скрипте</p>
<br>
<p>Запуск скрипта из виртуального окружения:</p>
<ul><li><b>python xls2marktplss.py sample.xlsx</b></li></ul>
<br>
<p>Тут:</p>
<p>- <i>sample.xlsx</i>: пример excel-файла c данными вашего товара. Имя файла может быть иным, но мтруктуру желательно повторить.</p>
<p>-- "Дерево категорий" - обязательно и не изменно.</p>
<p>-- "Основные характеристики" - обязательны.</p>
<p>-- Первую строку можете заполнять собственными данными.</p>
<p>-- Листы "Цены" и "оглавление" - обязательны</p>
