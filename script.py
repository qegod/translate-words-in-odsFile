from pyexcel_ods3 import get_data, save_data
from deep_translator import GoogleTranslator
import time
import os

def translate_words(input_file, output_file=None):
    """
    Переводит слова из первого столбца ODS файла и сохраняет перевод во второй столбец
    
    Args:
        input_file: путь к входному ODS файлу
        output_file: путь для сохранения результата (если None, сохранит с суффиксом '_translated')
    """
    
    # Если выходной файл не указан, добавляем суффикс
    if output_file is None:
        # Сохраняем рядом с оригиналом, но с другим именем
        directory = os.path.dirname(input_file)
        filename = os.path.basename(input_file)
        output_file = os.path.join(directory, filename.replace('.ods', '_translated.ods'))
    
    try:
        # Чтение данных из ODS файла
        print(f"Чтение файла {input_file}...")
        data = get_data(input_file)
        
        # Получаем первый лист (обычно это 'Sheet1')
        sheet_name = list(data.keys())[0]
        sheet_data = data[sheet_name]
        
        # Создаем список для результатов
        results = []
        translator = GoogleTranslator(source='en', target='ru')
        
        print("Начинаю перевод слов...")
        print("-" * 50)
        
        # Проходим по строкам
        for i, row in enumerate(sheet_data):
            if not row:  # Пропускаем пустые строки
                results.append([])
                continue
            
            word = str(row[0]).strip()
            
            if not word:  # Пропускаем пустые слова
                results.append([""])
                continue
            
            # Добавляем оригинальное слово в результат
            current_row = [word]
            
            try:
                # Переводим слово
                translation = translator.translate(word)
                current_row.append(translation)
                print(f"{i+1:3d}. {word:20} -> {translation}")
                
                # Добавляем паузу между запросами, чтобы не блокировать API
                time.sleep(0.1)
                
            except Exception as e:
                print(f"Ошибка при переводе слова '{word}': {e}")
                current_row.append("")  # Оставляем пустую ячейку при ошибке
            
            # Добавляем другие столбцы, если они были в оригинале
            if len(row) > 1:
                current_row.extend(row[1:])
            
            results.append(current_row)
        
        # Сохраняем результат в новый файл
        output_data = {sheet_name: results}
        save_data(output_file, output_data)
        
        print("-" * 50)
        print(f"\n✓ Готово! Результат сохранен в файл: {output_file}")
        print(f"✓ Переведено слов: {len([r for r in results if len(r) > 1 and r[1]])}")
        
    except Exception as e:
        print(f"✗ Произошла ошибка: {e}")

if __name__ == "__main__":
    # Вариант 1: Если файл лежит в той же папке, что и скрипт
    # Просто укажите имя файла (если он в той же папке)
    input_file = "english.ods"  # Замените на имя вашего файла
    
    # Вариант 2: Или используйте абсолютный путь
    # input_file = "C:/полный/путь/к/файлу.ods"
    
    # Вариант 3: Или используйте относительный путь
    # input_file = "./слова.ods"  # Точка означает текущую папку
    
    # Проверяем, существует ли файл
    if not os.path.exists(input_file):
        print(f"✗ ОШИБКА: Файл '{input_file}' не найден!")
        print("\nСоветы:")
        print("1. Убедитесь, что файл находится в той же папке, что и скрипт")
        print("2. Проверьте правильность имени файла")
        print("3. Убедитесь, что расширение файла .ods")
        print("\nСодержимое текущей папки:")
        try:
            files = os.listdir('.')
            ods_files = [f for f in files if f.endswith('.ods')]
            if ods_files:
                print("Найденные ODS файлы:")
                for f in ods_files:
                    print(f"  - {f}")
            else:
                print("  (нет ODS файлов)")
        except:
            pass
        exit(1)
    
    # Вызываем функцию перевода
    print("=" * 50)
    print("Запуск перевода английских слов на русский")
    print("=" * 50)
    
    translate_words(input_file)