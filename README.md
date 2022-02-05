# REHAU SKU плагин для MS Excel
## Назначение
Помощь в обработке клиентских заявок путем поиска продукции REHAU по запросу, выраженному в свободной форме, и в работе с прайс-листом BS REHAU

## Принцип работы
Плагин делает поисковый запрос в интернет-магазин REHAU, выполняет парсинг полученного ответа и выдает результат

## Реализованные функции
- Формулы для поиска информации 
    - Отображение наименования с помощью `=RAUNAME()`
    - Отображение артикула с помощью `=RAUSKU()`
    - Отображение цены с помощью формулы `=RAUPRICE()`
- Экспорт массива ячеек вида "Артикул - Количество" в прайс-лист
- Актуализация прайс-листа до последней версии
- Объединение нескольких прайс-листов в один файл
    - Сложением всех позиций по артикулам
    - С разнесением данных по колонкам в конечном файле

*Для работы функций "Экспорт", "Актуализация" и "Объединение" требуется указать путь к файлу пустого прайс-листа REHAU*

## Работа без установки
1. Запустить файл `RehauSku.Assist-AddIn-packed.xll` или `RehauSku.Assist-AddIn64-packed.xll` в зависимости от архитектуры приложения
2. Включить надстройку для данного сеанса в извещении системы безопасности

## Постоянная установка
1. Скопировать файл плагина в папку 
```
%AppData%\Microsoft\AddIns
```
2. В приложении Excel:

    Файл -> Параметры -> Надстройки -> 
    Управление: Надстройки Excel -> Перейти... -> Обзор

    Выбрать и включить файл плагина

## Использованные библиотеки
- [ExcelDna](https://github.com/Excel-DNA/ExcelDna)
- [Newtonsoft.Json](https://github.com/JamesNK/Newtonsoft.Json)
- [AngleSharp](https://github.com/AngleSharp/AngleSharp)