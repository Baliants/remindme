# remindme
The program to remid dates from Excel data file
To compile
1. Change paths in RemidMe.spec
2. Compile by comand > pyinstaller RemidMe.spec

===============

Данная программа была написана как выпускной проект на языке Python. Реализация на другом языке была бы более оптимальна, но требовалось использовать именно Python.

Цель программы –используя базу в формате Excel напоминать о событиях за разное количество дней.
Программа будет особенно полезна в отделах кадров. Она позволяет легко проследить разные сроки, связанные с большим количеством разных людей и событий. При этом сохраняются все возможности гибкой работы с базой данных в формате Excel.

Для корректной работы программы необходимо
- файл Excel должен называться data.xlsx и находится в папке с программой;
- ячейка Z1 должна содержать количество дней, за сколько нужно напомнить до события;
- данные о дате события должны начинается с ячейки A2 и ниже;
- один лист = один срок напоминания. Т.е. если у даны однотипные события, которые должны иметь разные сроки напоминания, для корректной работы программы, их нужно разделить на разные листы и выставить соответствующий срок напоминания
- имена листам лучше всего давать информативные, т.к. они будут обображатся в программе.

