## Введение:

Данный репозиторий содержит скрипты которые мне приходится использовать в работе.


Основной скрипт позволяет проанализировать списки материалов, выгруженных из системы САПР в Эксель формат, и на основе анализа составить необходимое представление о том какие количества данных материалов необходимо внести в производственную систему.

## Основной скрипт

### Алгоритм основного скрипта:

1. Создать объект таблицы данных из которой будут загружаться данные по расходу материалов.

2. Создать объект компонента для которого будет создаваться список материалов.

3. Создать объекты для каждого из используемых материалов.

4. Взять объект материала, проверить его наличие в таблице данных в достаточном колличестве.

   * В случае если материала недостаточно, расширить приближение.
   * Получить датафрейм на этот материал.

5. Взять объект компонента, применить его к датафрейму на материал. Проверить достаточно ли данных для расчета
   	* В случае если материала недостаточно, расширить приближение по продукту.
   	* Получить датафрейм для расчета

6. Расчитать колличество материала в продукте.
7. Записать полученое значение в данные.

Повторить для следующего объекта материала.

8. Записать данные в список материалов.
9. Проверить данные о создаваемом продкте и флаги.
10. Составить финальный список материалов с учетом продукта и флагов.

### Данные основного скрипта:

**Класс Таблица_данных:**

    Объект таблица_данных:
        Атрибуты:
            Название файла(ДБ)
            Датафрейм
    
        Функция Создать датафрейм:
            Создаем датафрейм из готовой Эксель таблицы.
            Возвращаем датафрейм.

**Класс Компонент:**

     Атрибуты класса:
        Параметры расширения приближения:
            1. Номер продукта(на случай если проверяем существующий продукт)
            2. Номер компонента(для всех его продуктов)
            3. Описание компонента(обычно они уникальны для СОС)
            4. Группа материала(данные доступны для всех материалов)
            5. СОС
            6. Все материалы
    Объект компонент:
        Атрибуты объекта:
            Название файла
            1. Номер продукта(на случай если проверяем существующий продукт)
            2. Номер компонента(для всех его продуктов)
            3. Описание компонента(обычно они уникальны для СОС)
            4. Группа материала(данные доступны для всех материалов)
            5. СОС
            6. Тип ремонта
            7. Датафрейм со списком материалов (через него будет проходить итерация)
            
        Функция создать объект(название файла)
            занести атрибуты из файла
            Создать датафрейм
            
        Функция создание датафрейма материалов:
            Для каждой строки датафрейма списка материалов
                материал.создать_объект
                материал.проверить_объект_в_тд
                материал.расчет значения
            
        Функция добавить_материал_в_данные:
            добавить материал к данным по индексу общего списка материалов.
            
        Функция соединить_данные.
            соеденить готовые данные в датафрейм
            
        Функция создания сводной таблицы:
            получить сводную таблицу для занесения данных напрямую в систему.

**Класс Материал:**

    Атрибуты класса:
        Параметры расширения приближения:
            1. Номер материала
            2. Описание материала
            3. Группа материала(данные доступны для всех материалов)
            4. СОС
            5. Все материалы
            
    Объект материал:
        Атрибуты объекта:
            1. Номер материала
            2. Описание материала
            3. Группа материала(данные доступны для всех материалов)
            4. СОС
            5. Словарь расширения
            6. Продукт к которому относиться
            
        Функция создать объект:
            Занести атрибуты из датафрейма и листа материалов с их данными.
            Создать словарь расширения
        
        Функция проверить объект в таблице данных:
            таблица_данных.создать_датафрейм
            петля для всех ключей словаря расширения    
                проверить материал (значение ключа расширения) в таблице данных
                если материал имеет недостаточное потребление или отсутсвует:
                    перейти на следующий ключ словаря расширения
                наоборот:
                    вернуть датафрейм материала
            петля для всех ключей словаря расширения  продукта   
                проверить продукт (значение ключа расширения) в датафрейме материалов
                если материал имеет недостаточное потребление или отсутсвует:
                    перейти на следующий ключ словаря расширения продукта.
                наоборот:
                    вернуть датафрейм материала
                    
         Функция расчет значения:
            посчитать значение расхода материала, применив статистический подход
            продукт.добавить_материал_в_данные

**Класс Аутпут:**

    Объект аутпут:
        атрибуты:
            датафрейм
        
        Функция выгрузить_данные:
            выгрузить данные в эксель файл

