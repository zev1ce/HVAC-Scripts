"""
Главный скрипт для автоматизации переноса данных ОВИК в AutoCAD
"""
import os
import sys
from pathlib import Path

# Добавляем папку src в путь для импорта модулей
current_dir = Path(__file__).parent
src_dir = current_dir / "src"
sys.path.insert(0, str(src_dir))

try:
    from excel_reader import ExcelDataReader
    from autocad_controller import AutoCADController
except ImportError as e:
    print(f"❌ Ошибка импорта модулей: {e}")
    print("Убедитесь, что файлы excel_reader.py и autocad_controller.py находятся в папке src/")
    sys.exit(1)


def main():
    """Основная функция программы"""
    print("=" * 60)
    print("АВТОМАТИЗАЦИЯ ПЕРЕНОСА ДАННЫХ ОВИК В AUTOCAD")
    print("=" * 60)
    
    # Проверяем наличие необходимых файлов
    config_path = "config/block_mapping.json"
    if not os.path.exists(config_path):
        print(f"❌ Файл конфигурации {config_path} не найден!")
        return
    
    # Инициализируем компоненты
    print("📖 Инициализация компонентов...")
    excel_reader = ExcelDataReader(config_path)
    autocad_controller = AutoCADController(config_path)
    
    # Подключаемся к AutoCAD
    print("🔗 Подключение к AutoCAD...")
    if not autocad_controller.connect_to_autocad():
        print("❌ Не удалось подключиться к AutoCAD!")
        print("   Убедитесь, что AutoCAD запущен и имеет открытый документ")
        return
    
    # Показываем информацию о блоках в чертеже
    print("\n🔍 Анализ блоков в чертеже...")
    blocks_info = autocad_controller.get_block_info()
    available_blocks = autocad_controller.get_available_blocks()
    
    # Проверяем наличие базового блока-шаблона
    default_block = autocad_controller.config.get("default_template_block", "HVAC_ROOM_DATA")
    has_template = autocad_controller.block_exists(default_block)
    print(f"📦 Блок-шаблон '{default_block}': {'✅ найден' if has_template else '❌ не найден'}")
    
    # Считаем блоки помещений
    room_blocks_count = 0
    if blocks_info:
        for block in blocks_info:
            if any(target in block['name'].upper() 
                  for target in autocad_controller.target_blocks):
                room_blocks_count += 1
    
    print(f"🏠 Блоков помещений в чертеже: {room_blocks_count}")
    
    # Выбираем режим работы с данными
    print("\n📂 Выбор источника данных...")
    print("  1. Один Excel файл (как раньше)")
    print("  2. Множественные источники данных (теплопотери + воздухообмены)")
    print("  3. Использовать пример данных")
    
    while True:
        data_mode = input("\nВыберите режим работы с данными (1-3): ").strip()
        if data_mode in ['1', '2', '3']:
            break
        print("Введите 1, 2 или 3")
    
    room_data = None
    excel_path = None
    
    if data_mode == '1':
        # Режим одного файла из папки data
        print("\n📄 Режим одного Excel файла...")
        
        excel_path = excel_reader.select_file_from_data_folder("Excel файл с данными ОВИК")
        
        if not excel_path:
            print("❌ Файл не выбран!")
            return
        
        # Проверяем структуру Excel файла
        print("📊 Проверка структуры файла...")
        if not excel_reader.validate_excel_structure(excel_path):
            print("❌ Структура Excel файла не соответствует ожидаемой!")
            return
        
        # Читаем данные из Excel
        print("📖 Чтение данных из Excel...")
        room_data = excel_reader.read_room_data(excel_path)
        
    elif data_mode == '2':
        # Режим множественных источников
        print("\n📊 Режим множественных источников данных...")
        
        # Создаем конфигурацию источников
        data_sources = excel_reader.create_data_sources_from_user_input()
        
        if not data_sources:
            print("❌ Не настроено ни одного источника данных!")
            return
        
        # Валидируем источники
        validation_results = excel_reader.validate_multiple_sources(data_sources)
        
        valid_sources = [name for name, is_valid in validation_results.items() if is_valid]
        
        if not valid_sources:
            print("❌ Ни один источник данных не прошел валидацию!")
            return
        
        if len(valid_sources) < len(data_sources):
            print(f"⚠️  Будут использованы только валидные источники: {', '.join(valid_sources)}")
        
        # Читаем и объединяем данные
        room_data = excel_reader.read_from_multiple_sources(data_sources)
        excel_path = "multiple_sources"  # Заглушка для дальнейшей логики
        
    else:  # data_mode == '3'
        # Режим примера данных
        print("\n📄 Использование примера данных...")
        example_file = "data/sample_hvac_data.xlsx"
        
        # Проверяем, есть ли уже пример файла
        if os.path.exists(example_file):
            print(f"✅ Найден пример файла: {example_file}")
            excel_path = example_file
        else:
            print("📂 Пример файла не найден, проверяем папку data/...")
            excel_files = excel_reader.get_excel_files_from_data_folder()
            
            if excel_files:
                print("💡 В папке data/ найдены Excel файлы. Хотите использовать один из них?")
                use_existing = input("Да/нет (Enter = создать новый пример): ").strip().lower()
                
                if use_existing in ['да', 'yes', 'y', 'д']:
                    excel_path = excel_reader.select_file_from_data_folder("Excel файл")
                    if not excel_path:
                        print("❌ Файл не выбран!")
                        return
                else:
                    # Создаем новый пример
                    print("   Создание примера файла...")
                    create_sample_excel(example_file)
                    excel_path = example_file
            else:
                # Создаем новый пример
                print("   Создание примера файла...")
                create_sample_excel(example_file)
                excel_path = example_file
        
        if not os.path.exists(excel_path):
            print("❌ Не удалось получить файл данных!")
            return
        
        # Читаем данные
        print("📖 Чтение данных...")
        room_data = excel_reader.read_room_data(excel_path)
    
    if not room_data:
        print("❌ Не удалось прочитать данные!")
        return
    
    print(f"✅ Прочитано данных о {len(room_data)} помещениях")
    
    # Показываем первые несколько записей
    print("\n📋 Пример данных:")
    for i, room in enumerate(room_data[:3]):
        print(f"  {i+1}. Помещение {room.get('room_number')}: {room.get('room_name')}")
        print(f"     Площадь: {room.get('area')} м², Приток: {room.get('air_supply')} м³/ч")
    
    # Определяем режим работы
    existing_room_blocks = autocad_controller.find_room_blocks()
    
    if existing_room_blocks:
        print(f"\n🔍 Найдено существующих блоков помещений: {len(existing_room_blocks)}")
        print("📋 Режимы работы:")
        print("  1. Обновить существующие блоки")
        print("  2. Создать новые блоки из шаблона")
        print("  3. Импортировать данные из блоков в Excel")
        print("  4. Отмена")
        
        while True:
            mode = input("\nВыберите режим (1-4): ").strip()
            if mode in ['1', '2', '3', '4']:
                break
            print("Введите 1, 2, 3 или 4")
        
        if mode == '1':
            # Режим обновления существующих блоков
            print("\n🔄 Обновление существующих блоков...")
            stats = autocad_controller.update_all_room_blocks(room_data)
            
            print("\n📊 Результаты обновления:")
            print(f"   ✅ Обновлено блоков: {stats['updated']}")
            print(f"   ❌ Не найдено данных: {stats['not_found']}")
            print(f"   ⚠️  Ошибок: {stats['errors']}")
            
        elif mode == '2':
            # Режим создания новых блоков
            create_new_blocks(autocad_controller, room_data, available_blocks)
            
        elif mode == '3':
            # Режим импорта данных из блоков
            import_data_from_blocks(autocad_controller, room_data, existing_room_blocks, excel_path)
            
        else:
            print("❌ Операция отменена пользователем")
    else:
        print("\n🏗️  Создание новых блоков из шаблона...")
        create_new_blocks(autocad_controller, room_data, available_blocks)


def create_new_blocks(autocad_controller, room_data, available_blocks):
    """Создает новые блоки из шаблона"""
    
    # Получаем базовый блок из конфигурации
    default_block = autocad_controller.config.get("default_template_block", "HVAC_ROOM_DATA")
    
    # Проверяем существование базового блока
    if autocad_controller.block_exists(default_block):
        template_block = default_block
        print(f"📦 Используется базовый блок-шаблон: {template_block}")
    else:
        print(f"⚠️  Базовый блок '{default_block}' не найден")
        
        if not available_blocks:
            print("❌ Нет доступных блоков для шаблона")
            return
        
        print("📋 Доступные блоки для шаблона:")
        for i, block_name in enumerate(available_blocks[:15], 1):
            print(f"  {i}. {block_name}")
        
        # Выбор альтернативного блока-шаблона
        while True:
            try:
                choice = input(f"\nВыберите блок-шаблон (1-{min(15, len(available_blocks))}): ").strip()
                
                if not choice:
                    print("❌ Создание блоков отменено")
                    return
                
                block_index = int(choice) - 1
                if 0 <= block_index < min(15, len(available_blocks)):
                    template_block = available_blocks[block_index]
                    break
                else:
                    print(f"Введите число от 1 до {min(15, len(available_blocks))}")
                    
            except ValueError:
                print("Введите корректный номер блока")
        
        print(f"📦 Выбран блок-шаблон: {template_block}")
    
    # Спрашиваем о настройках размещения
    print("\n⚙️  Изменить стандартные настройки размещения блоков? (нет)")
    custom_placement = input("Введите 'да' для настройки: ").strip().lower()
    
    if custom_placement in ['да', 'yes', 'y', 'д']:
        # Пользовательские настройки размещения
        print("\n📍 Способы размещения блоков:")
        print("  1. По сетке (автоматически)")
        print("  2. В заданных координатах")
        print("  3. Отмена")
        
        while True:
            placement_mode = input("Выберите способ размещения (1-3): ").strip()
            if placement_mode in ['1', '2', '3']:
                break
            print("Введите 1, 2 или 3")
        
        if placement_mode == '3':
            print("❌ Создание блоков отменено")
            return
            
        if placement_mode == '1':
            # Размещение по сетке с настройками
            print("\n⚙️  Настройка размещения по сетке:")
            
            try:
                start_x = float(input("Начальная X координата (по умолчанию 0): ") or "0")
                start_y = float(input("Начальная Y координата (по умолчанию 0): ") or "0")
                spacing_x = float(input("Расстояние между блоками по X (по умолчанию 100): ") or "100")
                spacing_y = float(input("Расстояние между блоками по Y (по умолчанию 50): ") or "50")
                
                stats = autocad_controller.create_blocks_from_template(
                    room_data, template_block, 
                    start_point=(start_x, start_y),
                    spacing=(spacing_x, spacing_y)
                )
                
            except ValueError:
                print("❌ Некорректные координаты, используются значения по умолчанию")
                stats = autocad_controller.create_blocks_from_template(room_data, template_block)
        else:
            # Размещение по координатам - код остается тот же
            handle_coordinate_placement(autocad_controller, room_data, template_block)
            return
    else:
        # Стандартные настройки размещения по сетке
        print("\n🏗️  Использование стандартных настроек размещения...")
        stats = autocad_controller.create_blocks_from_template(room_data, template_block)
    
    # Показываем результаты
    print("\n📊 Результаты создания блоков:")
    print(f"   ✅ Создано блоков: {stats['created']}")
    print(f"   ❌ Ошибок: {stats['errors']}")
    
    if stats['created'] > 0:
        print("\n🎉 Создание блоков завершено успешно!")
    else:
        print("\n⚠️  Ни один блок не был создан")


def handle_coordinate_placement(autocad_controller, room_data, template_block):
    """Обрабатывает размещение блоков по координатам"""
    reader = ExcelDataReader("config/block_mapping.json")
    
    # Проверяем, есть ли координаты в данных Excel
    if reader.has_coordinates_data(room_data):
        print("\n📍 Обнаружены координаты в Excel файле")
        use_excel_coords = input("Использовать координаты из Excel? (да/нет): ").strip().lower()
        
        if use_excel_coords in ['да', 'yes', 'y', 'д']:
            # Используем координаты из Excel
            coordinates = reader.extract_coordinates(room_data)
            # Фильтруем помещения с координатами
            valid_data = []
            valid_coords = []
            
            for room, coord in zip(room_data, coordinates):
                if coord is not None:
                    valid_data.append(room)
                    valid_coords.append(coord)
                else:
                    print(f"⚠️  Пропущено помещение {room.get('room_number')} - нет координат")
            
            if valid_coords:
                stats = autocad_controller.create_blocks_from_coordinates(
                    valid_data, template_block, valid_coords
                )
            else:
                print("❌ Не найдено помещений с координатами")
                return
        else:
            # Ручной ввод координат
            coordinates = get_manual_coordinates(room_data)
            if coordinates:
                stats = autocad_controller.create_blocks_from_coordinates(
                    room_data, template_block, coordinates
                )
            else:
                return
    else:
        # Ручной ввод координат
        coordinates = get_manual_coordinates(room_data)
        if coordinates:
            stats = autocad_controller.create_blocks_from_coordinates(
                room_data, template_block, coordinates
            )
        else:
            return
    
    # Показываем результаты
    print("\n📊 Результаты создания блоков:")
    print(f"   ✅ Создано блоков: {stats['created']}")
    print(f"   ❌ Ошибок: {stats['errors']}")
    
    if stats['created'] > 0:
        print("\n🎉 Создание блоков завершено успешно!")
    else:
        print("\n⚠️  Ни один блок не был создан")


def import_data_from_blocks(autocad_controller, room_data, existing_blocks, excel_path):
    """Импортирует данные из блоков AutoCAD обратно в Excel"""
    
    print("\n📥 Импорт данных из блоков AutoCAD в Excel...")
    
    # Читаем атрибуты из блоков
    block_data = autocad_controller.read_block_attributes(existing_blocks)
    
    if not block_data:
        print("❌ Не удалось прочитать данные из блоков")
        return
    
    # Создаем объект ExcelDataReader для работы с данными
    reader = ExcelDataReader("config/block_mapping.json")
    
    # Проверяем, работаем ли мы с множественными источниками
    if excel_path == "multiple_sources":
        print("\n⚠️  Обнаружен режим множественных источников данных")
        print("Для сохранения изменений будет создан новый объединенный Excel файл")
        
        # Создаем новый объединенный файл
        output_file = "data/объединенные_данные_обновлен.xlsx"
        
        # Сохраняем объединенные данные с обновлениями из блоков
        success = save_merged_data_with_block_updates(reader, room_data, block_data, output_file)
        
        if success:
            print(f"\n🎉 Создан объединенный файл: {output_file}")
        else:
            print("\n❌ Ошибка при создании объединенного файла")
    else:
        # Обычный режим - обновляем исходный файл
        print(f"\n💾 Текущий Excel файл: {excel_path}")
        save_choice = input("Сохранить изменения в исходный файл? (да/нет): ").strip().lower()
        
        output_file = None
        if save_choice not in ['да', 'yes', 'y', 'д']:
            base_name = os.path.splitext(excel_path)[0]
            output_file = f"{base_name}_обновлен.xlsx"
            print(f"📄 Будет создан новый файл: {output_file}")
        
        # Обновляем Excel данные
        success = reader.update_excel_with_block_data(
            excel_path, room_data, block_data, output_file
        )
        
        if success:
            print("\n🎉 Импорт данных завершен успешно!")
        else:
            print("\n❌ Произошла ошибка при импорте данных")


def save_merged_data_with_block_updates(reader, room_data, block_data, output_file):
    """Сохраняет объединенные данные с обновлениями из блоков в новый файл"""
    try:
        import pandas as pd
        # Получаем поля для обратного импорта из конфигурации  
        reverse_fields = reader.config.get("reverse_import_fields", ["supply_system", "extract_system"])
        
        print(f"\n🔄 Обновление данных полями: {', '.join(reverse_fields)}")
        
        # Создаем словарь для быстрого поиска данных блоков по номеру помещения
        block_dict = {}
        for block in block_data:
            room_num = block.get('room_number', '').strip()
            if room_num:
                block_dict[room_num] = block
        
        # Обновляем данные помещений
        updated_count = 0
        for room in room_data:
            room_num = str(room.get('room_number', '')).strip()
            
            if room_num in block_dict:
                block_info = block_dict[room_num]
                
                # Обновляем только указанные поля
                for field in reverse_fields:
                    if field in block_info and block_info[field]:
                        old_value = room.get(field, '')
                        new_value = block_info[field]
                        
                        if old_value != new_value:
                            room[field] = new_value
                            print(f"  📝 Помещение {room_num}: {field} '{old_value}' → '{new_value}'")
                
                updated_count += 1
        
        # Создаем папку если её нет
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        
        # Конвертируем в формат Excel колонок
        excel_columns = reader.config.get("excel_columns", {})
        
        excel_data = []
        for room in room_data:
            excel_row = {}
            for field, excel_col in excel_columns.items():
                excel_row[excel_col] = room.get(field, '')
            excel_data.append(excel_row)
        
        # Создаем DataFrame и сохраняем
        df = pd.DataFrame(excel_data)
        df.to_excel(output_file, index=False, engine='openpyxl', sheet_name="Объединенные данные ОВИК")
        
        print(f"\n✅ Обновлено {updated_count} помещений")
        print(f"💾 Объединенный файл создан: {len(excel_data)} записей")
        
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при создании объединенного файла: {e}")
        return False


def get_manual_coordinates(room_data):
    """Получает координаты от пользователя вручную"""
    print("\n📍 Ввод координат для размещения блоков:")
    print("Введите координаты в формате: x,y (например: 100,200)")
    print("Для завершения ввода нажмите Enter без координат")
    
    coordinates = []
    for i, room in enumerate(room_data):
        coord_input = input(f"Координаты для помещения {room.get('room_number', i+1)}: ").strip()
        
        if not coord_input:
            break
            
        try:
            x, y = map(float, coord_input.split(','))
            coordinates.append((x, y))
        except ValueError:
            print("❌ Некорректный формат координат, пропущено")
            
    if not coordinates:
        print("❌ Координаты не введены")
        return None
        
    return coordinates


def create_sample_excel(file_path: str):
    """Создает пример Excel файла с данными ОВИК"""
    try:
        import pandas as pd
        
        # Создаем папку если её нет
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # Пример данных
        sample_data = [
            {
                "Номер помещения": "101",
                "Наименование": "Офис главного инженера",
                "Площадь, м²": 25.5,
                "Приток, м³/ч": 150,
                "Вытяжка, м³/ч": 130,
                "Теплопотери, Вт": 1200,
                "Приточная система": "П1-1",
                "Вытяжная система": "В1-1", 
                "Класс чистоты": "А",
                "Температура, °C": 22
            },
            {
                "Номер помещения": "102", 
                "Наименование": "Конференц-зал",
                "Площадь, м²": 45.0,
                "Приток, м³/ч": 450,
                "Вытяжка, м³/ч": 400,
                "Теплопотери, Вт": 2800,
                "Приточная система": "П1-2",
                "Вытяжная система": "В1-2",
                "Класс чистоты": "B", 
                "Температура, °C": 20
            },
            {
                "Номер помещения": "103",
                "Наименование": "Кабинет проектировщика",
                "Площадь, м²": 18.2,
                "Приток, м³/ч": 110,
                "Вытяжка, м³/ч": 90,
                "Теплопотери, Вт": 950,
                "Приточная система": "П1-1",
                "Вытяжная система": "В1-1",
                "Класс чистоты": "А",
                "Температура, °C": 23
            },
            {
                "Номер помещения": "104",
                "Наименование": "Архив документации",
                "Площадь, м²": 12.0,
                "Приток, м³/ч": 60,
                "Вытяжка, м³/ч": 70,
                "Теплопотери, Вт": 600,
                "Приточная система": "П2-1",
                "Вытяжная система": "В2-1",
                "Класс чистоты": "C",
                "Температура, °C": 18
            },
            {
                "Номер помещения": "105",
                "Наименование": "Серверная",
                "Площадь, м²": 8.5,
                "Приток, м³/ч": 200,
                "Вытяжка, м³/ч": 250,
                "Теплопотери, Вт": 3500,
                "Приточная система": "П3-1",
                "Вытяжная система": "В3-1",
                "Класс чистоты": "D",
                "Температура, °C": 25
            }
        ]
        
        # Создаем DataFrame и сохраняем в Excel
        df = pd.DataFrame(sample_data)
        df.to_excel(file_path, index=False, sheet_name="Расчет воздухообмена")
        
        print(f"✅ Создан пример файла: {file_path}")
        
    except (PermissionError, OSError) as e:
        print(f"❌ Ошибка при создании примера файла: {e}")
    except Exception as e:
        print(f"❌ Неожиданная ошибка при создании примера файла: {e}")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n❌ Программа прервана пользователем")
    except SystemExit:
        # Нормальный выход программы
        pass
    except Exception as e:
        print(f"\n❌ Критическая ошибка: {e}")
        print("Обратитесь к разработчику для решения проблемы")
    
    input("\nНажмите Enter для выхода...")
