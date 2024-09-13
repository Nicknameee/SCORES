import pandas as pd

# Відкриваємо файл з урахуванням кодування UTF-8
with open('data.txt', 'r', encoding='utf-8') as file:
    # Читаємо всі рядки з файлу
    lines = file.readlines()

# Ініціалізуємо змінні для зберігання рядків 2 та 5
group_size = 14
result = []

# Проходимо через всі рядки
for i, line in enumerate(lines):
    # Визначаємо позицію рядка в групі
    pos_in_group = i % group_size

    # Перевіряємо, чи рядок є другим або п'ятим у групі
    if pos_in_group == 1:
        second_line = line.strip()  # Запам'ятовуємо другий рядок
    elif pos_in_group == 4:
        fifth_line = line.strip()  # Запам'ятовуємо п'ятий рядок

        # Додаємо у результат рядки 2 і 5 як окремі елементи
        result.append((second_line, fifth_line))

# Функція для отримання балів з рядка
def extract_score(line_tuple):
    # Припустимо, що бали знаходяться наприкінці кожного рядка, після коми
    try:
        # Відокремлюємо бали (число) з п'ятої частини
        score = int(line_tuple[1].split(',')[-1].strip())  # Останній елемент після коми
    except ValueError:
        score = 0  # Якщо не вдалося конвертувати у число, повертаємо 0
    return score

# Сортуємо результат по спаданню за балами з п'ятого рядка
sorted_result = sorted(result, key=extract_score, reverse=True)

# Створюємо DataFrame з двома колонками: "Рядок 2" і "Рядок 5"
df = pd.DataFrame(sorted_result, columns=["ПІБ", "БАЛ"])

# Записуємо результат у файл .xlsx
df.to_excel('output_file.xlsx', index=False)

print("Відсортований результат успішно записано в 'output_file.xlsx'")
