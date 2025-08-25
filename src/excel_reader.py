"""
Модуль для чтения данных из Excel таблиц с расчетами ОВИК
Поддерживает работу с множественными источниками данных
"""
import pandas as pd
import json
from typing import List, Dict, Optional, Tuple


class ExcelDataReader:
    """Класс для чтения и обработки данных из Excel файлов"""
    
    def __init__(self, config_path: str = "config/block_mapping.json"):
        """
        Инициализация читателя Excel данных
        
        Args:
            config_path: Путь к файлу конфигурации с настройками соответствий
        """
        self.config = self._load_config(config_path)
        self.excel_columns = self.config.get("excel_columns", {})
    
    def _load_config(self, config_path: str) -> Dict:
        """Загружает конфигурацию из JSON файла"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Конфигурационный файл {config_path} не найден")
            return {}
        except json.JSONDecodeError as e:
            print(f"Ошибка разбора JSON в файле {config_path}: {e}")
            return {}
    
    def check_file_availability(self, file_path: str) -> bool:
        """
        Проверяет доступность файла перед чтением
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            True если файл доступен, False иначе
        """
        import os
        
        print(f"🔍 Проверка доступности файла: {file_path}")
        
        # Проверяем существование файла
        if not os.path.exists(file_path):
            print("❌ Файл не найден")
            print("💡 Проверьте правильность пути к файлу")
            return False
        
        # Проверяем расширение файла  
        if not file_path.lower().endswith(('.xlsx', '.xls')):
            print("❌ Неподдерживаемый формат файла")
            print("💡 Используйте файлы .xlsx или .xls")
            return False
        
        # Проверяем права доступа
        if not os.access(file_path, os.R_OK):
            print("❌ Нет прав на чтение файла")
            print("💡 Проверьте права доступа к файлу")
            return False
        
        # Проверяем размер файла
        try:
            file_size = os.path.getsize(file_path)
            if file_size == 0:
                print("❌ Файл пустой")
                return False
            print(f"✅ Файл найден, размер: {file_size:,} байт")
        except OSError:
            print("❌ Ошибка при проверке размера файла")
            return False
        
        return True
    
    def read_room_data(self, excel_path: str, sheet_name: str = None) -> List[Dict]:
        """
        Читает данные помещений из Excel файла
        Автоматически определяет тип таблицы и использует соответствующий парсер
        
        Args:
            excel_path: Путь к Excel файлу
            sheet_name: Имя листа (если None, берется первый лист)
            
        Returns:
            Список словарей с данными помещений
        """
        # Нормализуем путь к файлу
        normalized_path = self.normalize_file_path(excel_path)
        
        # Предварительная проверка файла
        if not self.check_file_availability(normalized_path):
            # Если файл не найден, предлагаем исправления
            if excel_path != normalized_path:
                print("\n💡 Попробуйте использовать нормализованный путь:")
                print(f"📂 {normalized_path}")
            self.suggest_path_fixes(excel_path)
            return []
        
        # Определяем тип таблицы
        table_type = self.detect_table_type(normalized_path, sheet_name)
        
        # Используем соответствующий парсер
        if table_type == 'heat_loss':
            print("🔥 Использование парсера таблиц теплопотерь")
            return self.read_heat_loss_table(normalized_path, sheet_name)
        else:
            print("📊 Использование стандартного парсера таблиц")
            return self._read_standard_table(normalized_path, sheet_name)
    
    def _read_standard_table(self, excel_path: str, sheet_name: str = None) -> List[Dict]:
        """
        Читает данные из стандартной таблицы с заголовками в первой строке
        
        Args:
            excel_path: Путь к Excel файлу
            sheet_name: Имя листа (если None, берется первый лист)
            
        Returns:
            Список словарей с данными помещений
        """
        try:
            # Читаем Excel файл
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(excel_path)
            
            print(f"Прочитано {len(df)} строк из файла {excel_path}")
            
            # Очищаем данные от пустых строк
            df = df.dropna(subset=[self.excel_columns.get("room_number", "Номер помещения")])
            
            # Преобразуем в список словарей
            room_data = []
            for _, row in df.iterrows():
                room_info = self._extract_room_info(row)
                if room_info:
                    room_data.append(room_info)
            
            print(f"Обработано {len(room_data)} помещений")
            return room_data
            
        except FileNotFoundError:
            print(f"❌ Excel файл не найден: {excel_path}")
            self.suggest_path_fixes(excel_path)
            return []
        except PermissionError:
            print(f"❌ Нет доступа к файлу: {excel_path}")
            print("💡 Возможные решения:")
            print("   1. Закройте файл в Excel, если он открыт")
            print("   2. Проверьте права доступа к файлу")
            print("   3. Запустите программу от имени администратора")
            return []
        except UnicodeDecodeError as e:
            print(f"❌ Ошибка кодировки при чтении файла: {e}")
            print("💡 Возможные решения:")
            print("   1. Проверьте наличие кириллических символов в пути")
            print("   2. Переименуйте папки на английский язык")
            print("   3. Скопируйте файл в папку с английским названием")
            return []
        except ValueError as e:
            print(f"❌ Ошибка формата Excel файла: {e}")
            print("💡 Возможные решения:")
            print("   1. Проверьте, что файл имеет расширение .xlsx или .xls")
            print("   2. Откройте файл в Excel и пересохраните его")
            print("   3. Проверьте, что файл не поврежден")
            return []
        except Exception as e:
            print(f"❌ Неожиданная ошибка при чтении Excel файла: {e}")
            print(f"📝 Тип ошибки: {type(e).__name__}")
            print("💡 Обратитесь к разработчику с этой ошибкой")
            return []
    
    def _extract_room_info(self, row: pd.Series) -> Optional[Dict]:
        """
        Извлекает информацию о помещении из строки DataFrame
        
        Args:
            row: Строка DataFrame с данными помещения
            
        Returns:
            Словарь с данными помещения или None если данные некорректны
        """
        try:
            room_info = {}
            
            # Сопоставляем колонки Excel с полями данных
            for field, excel_col in self.excel_columns.items():
                if excel_col in row.index:
                    value = row[excel_col]
                    
                    # Обрабатываем разные типы данных
                    if pd.isna(value):
                        room_info[field] = ""
                    elif field in ["area", "air_supply", "air_extract", "heat_loss", "temperature", "coordinate_x", "coordinate_y"]:
                        # Числовые поля
                        room_info[field] = float(value) if value != "" else 0.0
                    else:
                        # Текстовые поля
                        room_info[field] = str(value).strip()
            
            # Проверяем обязательные поля
            if not room_info.get("room_number"):
                return None
                
            return room_info
            
        except (ValueError, KeyError, TypeError) as e:
            print(f"Ошибка при извлечении данных помещения: {e}")
            return None
    
    def get_column_mapping(self) -> Dict[str, str]:
        """Возвращает настройки соответствия колонок Excel и полей данных"""
        return self.excel_columns.copy()
    
    def has_coordinates_data(self, room_data_list: List[Dict]) -> bool:
        """
        Проверяет, содержат ли данные информацию о координатах
        
        Args:
            room_data_list: Список данных помещений
            
        Returns:
            True если есть координаты, False иначе
        """
        if not room_data_list:
            return False
            
        # Проверяем первые несколько записей на наличие координат
        for room_data in room_data_list[:5]:
            if (room_data.get("coordinate_x") is not None and 
                room_data.get("coordinate_y") is not None and
                room_data.get("coordinate_x") != 0 and 
                room_data.get("coordinate_y") != 0):
                return True
        
        return False
    
    def extract_coordinates(self, room_data_list: List[Dict]) -> List[Tuple[float, float]]:
        """
        Извлекает координаты из данных помещений
        
        Args:
            room_data_list: Список данных помещений
            
        Returns:
            Список кортежей (x, y) с координатами
        """
        coordinates = []
        
        for room_data in room_data_list:
            x = room_data.get("coordinate_x", 0)
            y = room_data.get("coordinate_y", 0)
            
            # Используем координаты только если они заданы
            if x != 0 or y != 0:
                coordinates.append((float(x), float(y)))
            else:
                coordinates.append(None)  # Помечаем отсутствие координат
        
        return coordinates
    
    def update_excel_with_block_data(self, excel_file_path: str, room_data: List[Dict], 
                                   block_data: List[Dict], output_file_path: str = None) -> bool:
        """
        Обновляет Excel файл данными из блоков AutoCAD
        
        Args:
            excel_file_path: Путь к исходному Excel файлу
            room_data: Исходные данные из Excel
            block_data: Данные прочитанные из блоков AutoCAD
            output_file_path: Путь для сохранения (если None - перезаписывает исходный)
            
        Returns:
            True если успешно, False иначе
        """
        try:
            # Получаем поля для обратного импорта из конфигурации  
            reverse_fields = self.config.get("reverse_import_fields", ["supply_system", "extract_system"])
            
            print(f"\n🔄 Обновление Excel данных полями: {', '.join(reverse_fields)}")
            
            # Создаем словарь для быстрого поиска данных блоков по номеру помещения
            block_dict = {}
            for block in block_data:
                room_num = block.get('room_number', '').strip()
                if room_num:
                    block_dict[room_num] = block
            
            # Обновляем данные помещений
            updated_count = 0
            for room_info in room_data:
                room_num = str(room_info.get('room_number', '')).strip()
                
                if room_num in block_dict:
                    block_info = block_dict[room_num]
                    
                    # Обновляем только указанные поля
                    for field in reverse_fields:
                        if field in block_info and block_info[field]:
                            old_value = room_info.get(field, '')
                            new_value = block_info[field]
                            
                            if old_value != new_value:
                                room_info[field] = new_value
                                print(f"  📝 Помещение {room_num}: {field} '{old_value}' → '{new_value}'")
                    
                    updated_count += 1
                else:
                    print(f"  ⚠️  Помещение {room_num}: данные в блоке не найдены")
            
            # Сохраняем обновленные данные в Excel
            output_path = output_file_path or excel_file_path
            self._save_updated_excel(room_data, output_path)
            
            print(f"\n✅ Обновлено {updated_count} помещений")
            print(f"💾 Файл сохранен: {output_path}")
            
            return True
            
        except Exception as e:
            print(f"❌ Ошибка обновления Excel: {e}")
            return False
    
    def _save_updated_excel(self, room_data: List[Dict], file_path: str):
        """Сохраняет обновленные данные в Excel файл"""
        # Конвертируем обратно в формат Excel колонок
        excel_columns = self.config.get("excel_columns", {})
        
        excel_data = []
        for room_info in room_data:
            excel_row = {}
            for field, excel_col in excel_columns.items():
                excel_row[excel_col] = room_info.get(field, '')
            excel_data.append(excel_row)
        
        # Создаем DataFrame и сохраняем
        df = pd.DataFrame(excel_data)
        df.to_excel(file_path, index=False, engine='openpyxl')
        
        print(f"💾 Excel файл обновлен: {len(excel_data)} записей")
    
    def validate_excel_structure(self, excel_path: str, sheet_name: str = None) -> bool:
        """
        Проверяет структуру Excel файла на соответствие ожидаемой
        
        Args:
            excel_path: Путь к Excel файлу
            sheet_name: Имя листа
            
        Returns:
            True если структура корректна, False иначе
        """
        try:
            if sheet_name:
                df = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=1)
            else:
                df = pd.read_excel(excel_path, nrows=1)
            
            columns = df.columns.tolist()
            missing_columns = []
            
            # Проверяем наличие обязательных колонок
            required_columns = [
                self.excel_columns.get("room_number"),
                self.excel_columns.get("room_name")
            ]
            
            for col in required_columns:
                if col and col not in columns:
                    missing_columns.append(col)
            
            if missing_columns:
                print(f"Отсутствуют обязательные колонки: {missing_columns}")
                print(f"Найденные колонки: {columns}")
                return False
            
            return True
            
        except (FileNotFoundError, PermissionError, ValueError) as e:
            print(f"Ошибка при проверке структуры Excel: {e}")
            return False
    
    def read_from_multiple_sources(self, data_sources: List[Dict]) -> List[Dict]:
        """
        Читает и объединяет данные из множественных Excel источников
        
        Args:
            data_sources: Список источников данных в формате:
                [
                    {
                        "file_path": "path/to/file.xlsx",
                        "sheet_name": "Лист1", 
                        "name": "heat_loss",
                        "fields": ["room_number", "room_name", "heat_loss", "temperature"],
                        "priority": 1
                    },
                    ...
                ]
                
        Returns:
            Объединенный список данных помещений
        """
        print("\n📊 Чтение данных из множественных источников...")
        
        all_room_data = {}  # Словарь по номеру помещения
        source_stats = {}   # Статистика по источникам
        
        for source in data_sources:
            source_name = source.get("name", "Неизвестный источник")
            file_path = source.get("file_path")
            sheet_name = source.get("sheet_name")
            allowed_fields = source.get("fields", [])
            priority = source.get("priority", 0)
            
            print(f"\n📂 Обработка источника '{source_name}': {file_path}")
            
            if not file_path:
                print(f"  ❌ Не указан путь к файлу для источника '{source_name}'")
                continue
            
            # Читаем данные из источника
            source_data = self.read_room_data(file_path, sheet_name)
            
            if not source_data:
                print(f"  ❌ Не удалось прочитать данные из '{source_name}'")
                source_stats[source_name] = {"rooms": 0, "fields_added": 0}
                continue
            
            rooms_processed = 0
            fields_added = 0
            
            # Обрабатываем каждое помещение из источника
            for room_record in source_data:
                room_num = str(room_record.get("room_number", "")).strip()
                
                if not room_num:
                    continue
                
                # Инициализируем данные помещения если его еще нет
                if room_num not in all_room_data:
                    all_room_data[room_num] = {
                        "room_number": room_num,
                        "_sources": {},  # Отслеживаем источники полей
                        "_priority": {}  # Отслеживаем приоритеты полей
                    }
                
                # Добавляем/обновляем поля из источника
                for field, value in room_record.items():
                    if field == "room_number":
                        continue
                        
                    # Проверяем, разрешено ли это поле для данного источника
                    if allowed_fields and field not in allowed_fields:
                        continue
                    
                    # Проверяем, нужно ли обновить поле (приоритет)
                    should_update = (
                        field not in all_room_data[room_num] or 
                        all_room_data[room_num].get(f"_{field}_priority", -1) <= priority
                    )
                    
                    if should_update and value not in ["", 0, 0.0, None]:
                        all_room_data[room_num][field] = value
                        all_room_data[room_num]["_sources"][field] = source_name
                        all_room_data[room_num]["_priority"][field] = priority
                        fields_added += 1
                
                rooms_processed += 1
            
            source_stats[source_name] = {
                "rooms": rooms_processed, 
                "fields_added": fields_added
            }
            
            print(f"  ✅ Обработано помещений: {rooms_processed}, добавлено полей: {fields_added}")
        
        # Преобразуем словарь обратно в список и очищаем служебные поля
        merged_data = []
        for room_num, room_data in all_room_data.items():
            # Удаляем служебные поля
            clean_room = {k: v for k, v in room_data.items() if not k.startswith("_")}
            merged_data.append(clean_room)
        
        # Выводим итоговую статистику
        print("\n📈 Статистика объединения данных:")
        print(f"  🏠 Всего помещений: {len(merged_data)}")
        
        for source_name, stats in source_stats.items():
            print(f"  📂 {source_name}: {stats['rooms']} помещений, {stats['fields_added']} полей")
        
        return merged_data
    
    def validate_multiple_sources(self, data_sources: List[Dict]) -> Dict[str, bool]:
        """
        Проверяет доступность и структуру множественных источников данных
        
        Args:
            data_sources: Список источников данных
            
        Returns:
            Словарь с результатами валидации для каждого источника
        """
        validation_results = {}
        
        print("\n🔍 Валидация источников данных...")
        
        for source in data_sources:
            source_name = source.get("name", "Неизвестный источник")
            file_path = source.get("file_path")
            sheet_name = source.get("sheet_name")
            
            print(f"\n📋 Проверка источника '{source_name}': {file_path}")
            
            if not file_path:
                print("  ❌ Не указан путь к файлу")
                validation_results[source_name] = False
                continue
            
            # Проверяем доступность файла и базовую структуру
            is_valid = self.validate_excel_structure(file_path, sheet_name)
            validation_results[source_name] = is_valid
            
            if is_valid:
                print(f"  ✅ Источник '{source_name}' прошел валидацию")
            else:
                print(f"  ❌ Источник '{source_name}' не прошел валидацию")
        
        return validation_results
    
    def create_data_sources_from_user_input(self) -> List[Dict]:
        """
        Создает конфигурацию источников данных на основе пользовательского ввода
        
        Returns:
            Список источников данных
        """
        print("\n⚙️ Настройка источников данных:")
        print("Выберите ваши Excel файлы из папки data/:")
        
        data_sources = []
        
        # Источник 1: Теплопотери
        print("\n🔥 Источник 1: Расчет теплопотерь")
        heat_loss_file = self.select_file_from_data_folder("файл расчета теплопотерь")
        if heat_loss_file:
            heat_loss_sheet = input("Имя листа (Enter для первого): ").strip() or None
            
            data_sources.append({
                "file_path": heat_loss_file,
                "sheet_name": heat_loss_sheet,
                "name": "heat_loss",
                "fields": ["room_number", "room_name", "heat_loss", "temperature"],
                "priority": 1
            })
        
        # Источник 2: Воздухообмены  
        print("\n💨 Источник 2: Расчет воздухообменов")
        air_exchange_file = self.select_file_from_data_folder("файл расчета воздухообменов")
        if air_exchange_file:
            air_exchange_sheet = input("Имя листа (Enter для первого): ").strip() or None
            
            data_sources.append({
                "file_path": air_exchange_file,
                "sheet_name": air_exchange_sheet,
                "name": "air_exchange", 
                "fields": ["room_number", "room_name", "area", "air_supply", "air_extract", 
                          "supply_system", "extract_system", "cleanliness_class", "temperature"],
                "priority": 2  # Более высокий приоритет для воздухообменов
            })
        
        if not data_sources:
            print("❌ Не указано ни одного источника данных")
            return []
        
        print(f"\n✅ Настроено источников данных: {len(data_sources)}")
        for source in data_sources:
            import os
            file_name = os.path.basename(source['file_path'])
            print(f"  📂 {source['name']}: {file_name}")
            
        return data_sources
    
    @staticmethod
    def normalize_file_path(file_path: str) -> str:
        """
        Нормализует путь к файлу для корректной работы с различными форматами
        
        Args:
            file_path: Исходный путь к файлу
            
        Returns:
            Нормализованный путь к файлу
        """
        import os
        
        # Убираем лишние пробелы
        normalized_path = file_path.strip()
        
        # Убираем кавычки если они есть
        if normalized_path.startswith('"') and normalized_path.endswith('"'):
            normalized_path = normalized_path[1:-1]
        
        # Нормализуем путь (конвертируем слеши, убираем дублирование)
        normalized_path = os.path.normpath(normalized_path)
        
        # Для Windows конвертируем в формат с прямыми слешами для лучшей совместимости
        if os.name == 'nt':  # Windows
            normalized_path = normalized_path.replace('\\', '/')
        
        return normalized_path
    
    @staticmethod
    def get_excel_files_from_data_folder() -> List[str]:
        """
        Получает список Excel файлов из папки data/
        
        Returns:
            Список путей к Excel файлам в папке data/
        """
        import os
        import glob
        
        data_folder = "data"
        
        # Создаем папку data если её нет
        if not os.path.exists(data_folder):
            os.makedirs(data_folder)
            print(f"📁 Создана папка {data_folder}/")
            return []
        
        # Ищем Excel файлы
        excel_patterns = [
            os.path.join(data_folder, "*.xlsx"),
            os.path.join(data_folder, "*.xls")
        ]
        
        excel_files = []
        for pattern in excel_patterns:
            excel_files.extend(glob.glob(pattern))
        
        # Убираем временные файлы Excel (начинающиеся с ~$)
        excel_files = [f for f in excel_files if not os.path.basename(f).startswith('~$')]
        
        # Сортируем по времени изменения (новые сверху)
        def get_modification_time(file_path):
            return os.path.getmtime(file_path)
        
        excel_files.sort(key=get_modification_time, reverse=True)
        
        return excel_files
    
    @staticmethod
    def select_file_from_data_folder(file_type: str = "Excel файл") -> str:
        """
        Позволяет пользователю выбрать файл из папки data/
        
        Args:
            file_type: Описание типа файла для пользователя
            
        Returns:
            Путь к выбранному файлу или None если файл не выбран
        """
        import os
        import datetime
        
        excel_files = ExcelDataReader.get_excel_files_from_data_folder()
        
        if not excel_files:
            print("❌ В папке data/ не найдено Excel файлов")
            print("💡 Скопируйте ваши .xlsx или .xls файлы в папку data/")
            return None
        
        print(f"\n📂 Найдено Excel файлов в папке data/: {len(excel_files)}")
        print(f"📋 Выберите {file_type}:")
        
        for i, file_path in enumerate(excel_files, 1):
            file_name = os.path.basename(file_path)
            file_size = os.path.getsize(file_path)
            mod_time = os.path.getmtime(file_path)
            
            # Форматируем размер файла
            if file_size < 1024:
                size_str = f"{file_size} б"
            elif file_size < 1024 * 1024:
                size_str = f"{file_size / 1024:.1f} КБ"
            else:
                size_str = f"{file_size / (1024 * 1024):.1f} МБ"
            
            # Форматируем время изменения
            mod_time_str = datetime.datetime.fromtimestamp(mod_time).strftime("%d.%m.%Y %H:%M")
            
            print(f"  {i}. {file_name}")
            print(f"     📊 {size_str}, изменен: {mod_time_str}")
        
        print(f"  {len(excel_files) + 1}. Отмена")
        
        while True:
            try:
                choice = input(f"\nВыберите файл (1-{len(excel_files) + 1}): ").strip()
                
                if not choice:
                    continue
                
                choice_num = int(choice)
                
                if choice_num == len(excel_files) + 1:
                    print("❌ Выбор файла отменен")
                    return None
                
                if 1 <= choice_num <= len(excel_files):
                    selected_file = excel_files[choice_num - 1]
                    print(f"✅ Выбран файл: {os.path.basename(selected_file)}")
                    return selected_file
                else:
                    print(f"Введите число от 1 до {len(excel_files) + 1}")
                    
            except ValueError:
                print("Введите корректный номер файла")
    
    def read_heat_loss_table(self, excel_path: str, sheet_name: str = None) -> List[Dict]:
        """
        Читает данные из специализированной таблицы теплопотерь
        Структура: начало с 20 строки, столбцы A (номер), B (название), S (теплопотери)
        
        Args:
            excel_path: Путь к Excel файлу
            sheet_name: Имя листа (если None, берется первый лист)
            
        Returns:
            Список словарей с данными помещений из таблицы теплопотерь
        """
        # Нормализуем путь к файлу
        normalized_path = self.normalize_file_path(excel_path)
        
        # Предварительная проверка файла
        if not self.check_file_availability(normalized_path):
            return []
        
        try:
            print(f"📊 Чтение таблицы теплопотерь из {normalized_path}")
            
            # Читаем Excel файл, начиная с 20 строки (индекс 19)
            if sheet_name:
                df = pd.read_excel(normalized_path, sheet_name=sheet_name, header=None, skiprows=19)
            else:
                df = pd.read_excel(normalized_path, header=None, skiprows=19)
            
            print(f"Прочитано {len(df)} строк данных из таблицы теплопотерь")
            
            room_data = []
            processed_count = 0
            
            for index, row in df.iterrows():
                try:
                    # Столбец A (индекс 0) - номер помещения
                    room_number = row.iloc[0] if len(row) > 0 and not pd.isna(row.iloc[0]) else None
                    
                    # Столбец B (индекс 1) - название помещения  
                    room_name = row.iloc[1] if len(row) > 1 and not pd.isna(row.iloc[1]) else None
                    
                    # Столбец S (индекс 18) - теплопотери
                    heat_loss = row.iloc[18] if len(row) > 18 and not pd.isna(row.iloc[18]) else None
                    
                    # Проверяем, что у нас есть основные данные
                    if room_number is None or room_name is None:
                        continue
                    
                    # Очищаем и проверяем номер помещения
                    room_number_str = str(room_number).strip()
                    if not room_number_str or room_number_str in ['nan', 'None', '']:
                        continue
                    
                    # Очищаем название помещения
                    room_name_str = str(room_name).strip()
                    if not room_name_str or room_name_str in ['nan', 'None', '']:
                        continue
                    
                    # Обрабатываем теплопотери
                    heat_loss_value = 0.0
                    if heat_loss is not None and not pd.isna(heat_loss):
                        try:
                            heat_loss_value = float(heat_loss)
                        except (ValueError, TypeError):
                            heat_loss_value = 0.0
                    
                    # Создаем запись в стандартном формате
                    room_info = {
                        "room_number": room_number_str,
                        "room_name": room_name_str,
                        "heat_loss": heat_loss_value,
                        "temperature": 20.0  # Значение по умолчанию для температуры
                    }
                    
                    room_data.append(room_info)
                    processed_count += 1
                    
                except Exception as e:
                    print(f"  ⚠️  Ошибка обработки строки {index + 20}: {e}")
                    continue
            
            print(f"✅ Обработано {processed_count} помещений из таблицы теплопотерь")
            
            # Показываем первые несколько записей для проверки
            if room_data:
                print("📋 Пример обработанных данных:")
                for i, room in enumerate(room_data[:3]):
                    print(f"  {i+1}. {room['room_number']}: {room['room_name']} - {room['heat_loss']} Вт")
            
            return room_data
            
        except Exception as e:
            print(f"❌ Ошибка при чтении таблицы теплопотерь: {e}")
            return []
    
    def detect_table_type(self, excel_path: str, sheet_name: str = None) -> str:
        """
        Определяет тип таблицы (стандартная или таблица теплопотерь)
        
        Args:
            excel_path: Путь к Excel файлу
            sheet_name: Имя листа
            
        Returns:
            'standard' - обычная таблица с заголовками в первой строке
            'heat_loss' - таблица теплопотерь со сложной структурой
        """
        try:
            # Читаем первые несколько строк файла
            if sheet_name:
                df_sample = pd.read_excel(excel_path, sheet_name=sheet_name, nrows=25, header=None)
            else:
                df_sample = pd.read_excel(excel_path, nrows=25, header=None)
            
            # Ищем признаки таблицы теплопотерь
            for index, row in df_sample.iterrows():
                if len(row) > 0:
                    cell_text = str(row.iloc[0]).lower()
                    if 'расчет теплопотерь' in cell_text or 'теплопотерь' in cell_text:
                        print("🔥 Обнаружена таблица теплопотерь")
                        return 'heat_loss'
                    
                    # Проверяем наличие стандартных заголовков
                    if any(header in cell_text for header in ['номер помещения', 'наименование', 'площадь']):
                        print("📊 Обнаружена стандартная таблица")
                        return 'standard'
            
            # Если не удалось точно определить, проверяем первую строку
            if len(df_sample) > 0:
                first_row = df_sample.iloc[0]
                if len(first_row) > 2:
                    # Если в первой строке есть осмысленные заголовки
                    headers = [str(cell).lower() for cell in first_row[:5] if not pd.isna(cell)]
                    if any('номер' in h or 'название' in h or 'площадь' in h for h in headers):
                        return 'standard'
            
            print("📋 Тип таблицы не определен, используется стандартный")
            return 'standard'
            
        except Exception as e:
            print(f"⚠️  Ошибка определения типа таблицы: {e}")
            return 'standard'
    
    @staticmethod
    def suggest_path_fixes(file_path: str):
        """
        Предлагает исправления для проблемного пути к файлу
        
        Args:
            file_path: Проблемный путь к файлу
        """
        print("\n🔧 Предложения по исправлению пути:")
        print(f"📂 Исходный путь: {file_path}")
        print(f"📂 Нормализованный: {ExcelDataReader.normalize_file_path(file_path)}")
        
        # Предлагаем альтернативные варианты
        if '\\' in file_path:
            alt1 = file_path.replace('\\', '/')
            print(f"📂 Вариант 1 (прямые слеши): {alt1}")
            
            alt2 = file_path.replace('\\', '\\\\')
            print(f"📂 Вариант 2 (двойные слеши): {alt2}")
        
        # Проверяем проблемы с кириллицей
        has_cyrillic = any(ord(char) > 127 for char in file_path)
        if has_cyrillic:
            print("⚠️  Обнаружены кириллические символы в пути")
            print("💡 Рекомендации:")
            print("   1. Скопируйте файл в папку с английским названием")
            print("   2. Или в папку data/ вашего проекта")
        
        # Проверяем пробелы
        if ' ' in file_path:
            print("⚠️  Обнаружены пробелы в пути")
            print("💡 Это обычно не проблема, но если возникают ошибки:")
            print("   1. Заключите путь в кавычки")
            print("   2. Переименуйте папки без пробелов")


# Пример использования
if __name__ == "__main__":
    reader = ExcelDataReader()
    
    # Пример чтения данных
    data = reader.read_room_data("data/sample_hvac_data.xlsx")
    
    for room in data[:3]:  # Показываем первые 3 помещения
        print(f"Помещение {room.get('room_number')}: {room.get('room_name')}")
        print(f"  Площадь: {room.get('area')} м²")
        print(f"  Приток: {room.get('air_supply')} м³/ч")
        print(f"  Вытяжка: {room.get('air_extract')} м³/ч")
        print()
