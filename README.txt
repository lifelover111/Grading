Программа для выставления баллов.
Формат входного файла - таблица xlsx (название файла пишется в файле config) с следующими столбцами:
ФИО/Ожидаемая оценка/Посещаемость/Задание 1/Задание 2/.../Задание n
Последовательность столбцов строго в указанном порядке, первая строка - заголовки, количество
столбцов заданий неограничено. Результат работы программы - столбец с соответствующими баллами.

ФИО - любая строка
Ожидаемая оценка - количество ожидаемых баллов. В случае пустой ячейки считается равным 60
Посещаемость - доля посещенных занятий, отрезок [0,1]
Задание - любая строка. Заполненная ячейка задания означает, что студент выполнил это задание


Оценка считается следующим образом:
1. grade = pointsForTasks * (tasks/maxTasks) + (100 - pointsForTasks) * attendance  
(предварительная оценка, где tasks - число выполненных студентом заданий; 
maxTasks - максимальное число заданий, выполненных одним студентом; 
attendance - посещаемость;
pointsForTasks - число баллов, отведенных на задания, берется из config файла)

2. Если grade ниже ожидаемой оценки, считается их разница delta. В случае grade<60 ожидаемой оценкой считается 60
2.1 grade += delta*avg, где avg - среднее значение между долей посещаемости и долей выполненных заданий(tasks/maxTasks)

3. Если grade < 60 считается значение p = 60 - grade. Если p < offsetThreshold то grade = 60, где offsetThreshold берется
из файла конфигурации

4. Значение grade записывается в качестве результата