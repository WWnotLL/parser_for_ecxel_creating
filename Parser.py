import os #для работы с путями файлов
import pandas as pd #для работы с Excel и CSV файлами
import re #для поиска и извлечения текста по шаблонам
from pathlib import Path #для поиска папок по шаблону, перебора файлов в директории
import sys #проверка пути исполняемого файла и запущен как .py или .exe


def extract_requirement_id(text):
    ###Извлечение ID требования из текста###
    if pd.isna(text) or not isinstance(text, str):
        return None
    match = re.search(r'<Вставьте сюда постоянную часть идентификатора>-\d{4}', text) #4 - количество символов в изменяющейся части идентификатора
    return match.group() if match else None


def process_test_directory(root_dir):
    ###Основная функция обработки директории###
    results = []
    
    #Сначала находим все папки верхнего уровня
    for item in Path(root_dir).iterdir():
        if item.is_dir() and not item.name.startswith('Вставляем постоянную часть названия верхнеуровневых папок'):  #пропускаем сами папки на верхнем уровне
            print(f"\nОбработка папки верхнего уровня: {item.name}")
            
            #Ищем внутри папки верхнего уровня папки
            for test_dir in item.rglob('Вставляем постоянную часть названия верхнеуровневых папок*'): #* в конце предложения нужно оставить
                if test_dir.is_dir():
                    test_id = test_dir.name
                    print(f"  Найдена папка с тестами: {test_id}")

                    #Обработка XLSM файлов
                    xlsm_files = list(test_dir.rglob('*.xlsm'))
                    for xlsm_file in xlsm_files:
                        print(f"    Обработка XLSM: {xlsm_file.name}")
                        results.extend(process_xlsm_file(xlsm_file, test_id))

                    #Обработка TBI файлов, в которых расположен анализ кода
                    tbi_files = list(test_dir.rglob('*.tbi'))
                    for tbi_file in tbi_files:
                        print(f"    Обработка TBI: {tbi_file.name}")
                        results.extend(process_tbi_file(tbi_file, test_id))

    return results


def process_xlsm_file(file_path, test_id):
    ###Обработка XLSM файла###
    results = []
    try:
        xl = pd.ExcelFile(file_path)
        
        #Обрабатываем все листы, начиная со второго (индекс 1)
        for sheet_name in xl.sheet_names[1:]:
            if 'Группа' in sheet_name:
                group_match = re.search(r'Группа\s*(\d+)', sheet_name)
                if group_match:
                    group_num = group_match.group(1)
                    try:
                        #Читаем лист без заголовков
                        df = pd.read_excel(xl, sheet_name, header=None)
                        results.extend(parse_group_sheet(df, group_num, test_id, file_path.name))
                    except Exception as e:
                        print(f"      Ошибка в листе {sheet_name}: {e}")

    except Exception as e:
        print(f"    Ошибка чтения файла {file_path}: {e}")

    return results


def parse_group_sheet(df, group_num, test_id, file_name):
    ###Парсинг листа группы###
    results = []
    
    #Множество для хранения уже обработанных пар
    processed_pairs = set()

    try:
        #Проверяем, что в листе достаточно строк
        if len(df) < 3:
            return results
            
        #Строка 0 - заголовки примеров
        #Строка 1 - заголовки шагов
        example_row = df.iloc[0, :]
        step_row = df.iloc[1, :]

        #Проходим по строкам с требованиями (начиная со строки 2)
        for row in range(2, len(df)):
            #Проверяем столбец B (индекс 1) на наличие требования
            if len(df.columns) > 1:
                req_id = extract_requirement_id(df.iloc[row, 1])

                if req_id:
                    #Ищем отметки X в строке (начиная со столбца C, индекс 2)
                    for col in range(2, len(df.columns)):
                        try:
                            cell_value = df.iloc[row, col]
                            #Проверяем на X (учитываем разные варианты написания)
                            if cell_value in ['X', 'x', 'Х', 'х'] or (
                                    isinstance(cell_value, str) and cell_value.strip().upper() in ['X', 'Х']):
                                
                                #Получаем значения из строк заголовков
                                example_header = str(example_row.iloc[col]) if col < len(example_row) else ""
                                step_header = str(step_row.iloc[col]) if col < len(step_row) else ""
                                
                                #Извлекаем номера примеров и шагов
                                example_match = re.search(r'(\d+)', example_header)
                                step_match = re.search(r'(\d+)', step_header)
                                
                                if example_match and step_match:
                                    example_num = example_match.group(1)
                                    step_num = step_match.group(1)
                                    
                                    #Формируем идентификатор тестового примера
                                    test_example_id = f"Группа_{group_num}_Пример_{example_num}_Шаг_{step_num}"
                                    
                                    #Проверяем на дубликаты
                                    pair_key = (req_id, test_id, file_name, test_example_id)
                                    if pair_key not in processed_pairs:
                                        processed_pairs.add(pair_key)
                                        results.append({
                                            'Требования к ПО низкого уровня': req_id,
                                            'Идентификатор модульного теста': test_id,
                                            'Идентификатор тестового файла': file_name,
                                            'Идентификатор тестового примера': test_example_id
                                        })
                        except Exception as e:
                            continue
    except Exception as e:
        print(f"      Ошибка парсинга листа: {e}")

    return results


def process_tbi_file(file_path, test_id):
    ###Обработка TBI файла с разными кодировками###
    results = []

    #Пробуем разные кодировки
    encodings = ['utf-8', 'windows-1251', 'cp1251', 'iso-8859-1', 'cp866']

    for encoding in encodings:
        try:
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()

            #Успешно прочитали файл, парсим содержимое
            results.extend(parse_tbi_content(content, test_id, file_path.name))
            break  #Если успешно, выходим из цикла кодировок

        except UnicodeDecodeError:
            continue  #Пробуем следующую кодировку
        except Exception as e:
            print(f"      Ошибка чтения TBI файла {file_path} с кодировкой {encoding}: {e}")
            break

    return results


def parse_tbi_content(content, test_id, file_name):
    ###Парсинг содержимого TBI файла###
    results = []
    
    try:
        #Разделяем содержимое на блоки по разделителям (**********)
        blocks = re.split(r'\*{10,}', content)
        
        #Множество для хранения уже обработанных пар
        processed_pairs = set()
        
        for block in blocks:
            block = block.strip()
            if not block:
                continue
                
            #Ищем идентификатор тестового примера в блоке
            example_match = re.search(r'Идентификатор тестового примера:\s*Группа\s*(\d+)\s*Пример\s*(\d+)', 
                                     block, re.IGNORECASE)
            if not example_match:
                continue
                
            group_num = example_match.group(1)
            example_num = example_match.group(2)
            test_example_id = f"Группа_{group_num}_Пример_{example_num}"
            
            #Ищем ВСЕ требования в этом блоке
            requirements = re.findall(r'<Вставьте сюда постоянную часть идентификатора>-\d{4}', block)
            
            #Добавляем уникальные требования для этого примера
            for req_id in requirements:
                pair_key = (req_id, test_id, file_name, test_example_id)
                if pair_key not in processed_pairs:
                    processed_pairs.add(pair_key)
                    results.append({
                        'Требования к ПО низкого уровня': req_id,
                        'Идентификатор модульного теста': test_id,
                        'Идентификатор тестового файла': file_name,
                        'Идентификатор тестового примера': test_example_id
                    })
                
    except Exception as e:
        print(f"      Ошибка парсинга содержимого TBI: {e}")

    return results


def save_to_excel(results, output_file):
    ###Сохранение результатов в Excel###
    if not results:
        print("Нет данных для сохранения")
        return

    try:
        df = pd.DataFrame(results)
        
        #Удаляем возможные дубликаты из финального результата
        df = df.drop_duplicates()

        #Убеждаемся, что имя файла имеет правильное расширение
        if not output_file.endswith('.xlsx'):
            output_file += '.xlsx'

        #Упорядочиваем столбцы
        columns_order = [
            'Требования к ПО низкого уровня',
            'Идентификатор модульного теста',
            'Идентификатор тестового файла',
            'Идентификатор тестового примера'
        ]

        #Проверяем наличие всех столбцов
        for col in columns_order:
            if col not in df.columns:
                df[col] = ""

        df = df[columns_order]
        
        #Сортируем для удобства просмотра
        df = df.sort_values(by=['Требования к ПО низкого уровня', 'Идентификатор модульного теста'])

        #Сохраняем в Excel
        df.to_excel(output_file, index=False, engine='openpyxl')
        print(f"\nУспешно сохранено {len(df)} уникальных записей в файл: {output_file}")

    except Exception as e:
        print(f"Ошибка при сохранении в Excel: {e}")
        #Пробуем сохранить как CSV в случае ошибки
        try:
            csv_file = output_file.replace('.xlsx', '.csv')
            df.to_csv(csv_file, index=False, encoding='utf-8-sig')
            print(f"Данные также сохранены в CSV файл: {csv_file}")
        except:
            print("Не удалось сохранить данные в альтернативном формате")


def get_executable_directory():
    ###Получение директории, где находится исполняемый файл###
    if getattr(sys, 'frozen', False):
        #Если запущен как exe файл
        return os.path.dirname(sys.executable)
    else:
        #Если запущен как Python скрипт
        return os.path.dirname(os.path.abspath(__file__))


def main():
    ###Основная функция - автоматический режим###
    print("=== Парсер трассировки требований ===")
    print("Автоматический режим - поиск тестов в текущей директории")
    
    #Получаем директорию, где находится исполняемый файл
    exe_dir = get_executable_directory()
    print(f"Директория исполняемого файла: {exe_dir}")
    
    #Проверяем, есть ли папки верхнего уровня
    top_level_folders = [item for item in Path(exe_dir).iterdir() 
                        if item.is_dir() and not item.name.startswith('Вставляем постоянную часть названия верхнеуровневых папок')]
    
    if not top_level_folders:
        print("ОШИБКА: В директории исполняемого файла не найдено папок верхнего уровня (ADC, MAIN и т.д.)!")
        print("Поместите исполняемый файл в папку, содержащую папки с тестами, и запустите снова.")
        input("Нажмите Enter для выхода...")
        return
    
    print(f"Найдено {len(top_level_folders)} папок верхнего уровня")
    
    #Обрабатываем директорию исполняемого файла
    print("\nНачинаю обработку...")
    data = process_test_directory(exe_dir)
    
    #Создаем выходной файл рядом с исполняемым файлом
    output_file = os.path.join(exe_dir, "trace.xlsx")
    
    if data:
        print(f"\nНайдено {len(data)} записей трассировки")
        save_to_excel(data, output_file)
    else:
        print("Не найдено данных для трассировки")
    
    #Пауза перед закрытием
    print("\nОбработка завершена.")
    input("Нажмите Enter для выхода...")


if __name__ == "__main__":
    main()
