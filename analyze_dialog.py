import json
import pandas as pd
from nltk.sentiment import SentimentIntensityAnalyzer
from random import randint

# Функція для аналізу настрою тексту
def analyze_sentiment(text):
    sia = SentimentIntensityAnalyzer()
    return sia.polarity_scores(text)

# Функція для генерації випадкового балу
def generate_random_score(min_score, max_score):
    return randint(min_score, max_score)

# Функція для аналізу діалогу
def analyze_dialog(dialog):
    operator_lines = [line for line in dialog['messages'] if line['type'] == 'Оператор']
    subscriber_lines = [line for line in dialog['messages'] if line['type'] == 'Абонент']
    operator_text = ' '.join([line['message'] for line in operator_lines])
    subscriber_text = ' '.join([line['message'] for line in subscriber_lines])
    operator_sentiment = analyze_sentiment(operator_text)
    subscriber_sentiment = analyze_sentiment(subscriber_text)
    results = {
        'Оператор': dialog['messages'][0]['name'],
        'Репліки': len(operator_lines),
        'Пунктуація': generate_random_score(8, 10),
        'Простота розмови': generate_random_score(7, 10),
        'Рейтинг оператора': generate_random_score(8, 10),
        'Рейтинг діалогу': generate_random_score(7, 10),
        'Розуміння абонента': generate_random_score(8, 10),
        'Operator sentiment': operator_sentiment,
        'Subscriber sentiment': subscriber_sentiment,
    }
    return results

# Завантаження JSON-файлу
with open('GetDummyChats.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

# Аналіз кожного діалогу
results_list = []
for item in data['data']:
    results = analyze_dialog(item)
    results_list.append(results)

# Створення DataFrame з результатами
df = pd.DataFrame(results_list)

# Збереження DataFrame в Excel-файл
df.to_excel('dialog_analysis_results.xlsx', index=False)

# Висновок результатів
print("Analysis results saved to 'dialog_analysis_results.xlsx'")