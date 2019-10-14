# импорт необходимых библиотек
import re
import xlwt
import nltk
import string
from nltk.corpus import stopwords
import pymorphy2
morph = pymorphy2.MorphAnalyzer()
#Загрузку данных для nltk произвести при первом запускке
#nltk.download()
#nltk.download('punkt')

# коэффициент частоты повторения слов
key_index = 0.5

# чтение текста из нужного файла
fText = input("Расположение исходного файла:  ")
#fText = "test.txt"
f = open(fText, "r")
data_f = f.read()
f.close()

# очистка текста от  чисел
data_f = re.sub(r"\d+", "", data_f, flags=re.UNICODE)

#токенизация слов
data_tok = nltk.word_tokenize(data_f)

# чистка списка слов от знаков пунктуации
data_tok = [i for i in data_tok if (i not in string.punctuation)]
# все слова строчными буквами
data_norm = [word.lower() for word in data_tok]
#постановка слов в начальную форму
for i in range(len(data_norm)):
    data_norm[i] = morph.parse(data_norm[i])[0].normal_form

# подсчет повторяющихся слов
frequency = {}
for word in data_norm:
    count = frequency.get(word, 0)
    frequency[word] = count + 1

# удаление стоп-слов
stop_words = stopwords.words('russian')
stop_words.extend(['что', 'это', 'так', 'вот', 'быть', 'как', 'в', '—', 'к', 'на', '»', '«'])
data_norm = [i for i in data_norm if (i not in stop_words)]

data_clean = []
sw = []
#чистка списка слов от повторений
for word in data_norm:
    if word not in data_clean:
        data_clean.append(word)
        sw.append(frequency[word])

#сортировка списка по частоте повторения в тексте
data_sort = []
s = max(sw) + 1
for i in range(s):
    for word in data_clean:
        if frequency[word] == i:
            data_sort.append(word)
    i = i - 1

data_sort = data_sort[::-1]

#запись результата в файл с указанием частоты повторения в тексте
wb = xlwt.Workbook()
name_out_f = input('Введите имя выходного файла ')
ws = wb.add_sheet(name_out_f)

i = 0
for word in data_sort:
    if frequency[word] > (max(sw)*key_index):
        ws.write(i, 0, word)
        ws.write(i, 1, frequency[word])
        i = i + 1

wb.save(name_out_f+'.xls')