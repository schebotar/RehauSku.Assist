# REHAU SKU плагин для MS Excel
## Назначение
Служит для ускоренной обработки клиентских заявок путем поиска наименования и артикула по запросу, выраженному в свободной форме.

После загрузки плагина в Excel позволяет с помощью простой формулы вида
```excel
=RAUNAME(A1)
```
или
```excel
=RAUNAME("Запрос")
```
получить полное наименование и артикул позиции из каталога интернет-магазина REHAU https://shop-rehau.ru/
## Принцип работы
Прототип делает поисковый запрос в интернет-магазине, выполняет парсинг полученного ответа и выдает первый результат
![Видео работы](./README.files/video.webm.mov/)