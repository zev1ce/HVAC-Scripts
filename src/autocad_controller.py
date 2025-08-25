"""
Модуль для управления AutoCAD через COM интерфейс
"""
import win32com.client
import json
from typing import List, Dict, Optional, Tuple


class AutoCADController:
    """Класс для управления AutoCAD через COM"""
    
    def __init__(self, config_path: str = "config/block_mapping.json"):
        """
        Инициализация контроллера AutoCAD
        
        Args:
            config_path: Путь к файлу конфигурации
        """
        self.acad_app = None
        self.acad_doc = None
        self.config = self._load_config(config_path)
        self.block_attributes = self.config.get("block_attributes", {})
        self.target_blocks = self.config.get("target_blocks", [])
        
    def _load_config(self, config_path: str) -> Dict:
        """Загружает конфигурацию из JSON файла"""
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"Конфигурационный файл {config_path} не найден")
            return {}
        except json.JSONDecodeError:
            print(f"Ошибка разбора JSON в файле {config_path}")
            return {}
    
    def connect_to_autocad(self) -> bool:
        """
        Подключается к запущенному AutoCAD
        
        Returns:
            True если подключение успешно, False иначе
        """
        try:
            # Подключение к активному AutoCAD
            self.acad_app = win32com.client.GetActiveObject("AutoCAD.Application")
            
            # Получаем список открытых документов
            documents = []
            for i in range(self.acad_app.Documents.Count):
                doc = self.acad_app.Documents.Item(i)
                documents.append((doc, doc.Name))
            
            if not documents:
                print("❌ В AutoCAD нет открытых документов")
                return False
            
            # Если документов несколько, даем выбор
            if len(documents) > 1:
                print(f"\n📂 Найдено открытых документов: {len(documents)}")
                for i, (doc, name) in enumerate(documents):
                    active_mark = " (активный)" if doc == self.acad_app.ActiveDocument else ""
                    print(f"  {i+1}. {name}{active_mark}")
                
                while True:
                    try:
                        choice = input(f"\nВыберите документ (1-{len(documents)}) или Enter для активного: ").strip()
                        
                        if not choice:
                            # Используем активный документ
                            self.acad_doc = self.acad_app.ActiveDocument
                            break
                        
                        doc_index = int(choice) - 1
                        if 0 <= doc_index < len(documents):
                            self.acad_doc = documents[doc_index][0]
                            # Делаем выбранный документ активным
                            self.acad_app.ActiveDocument = self.acad_doc
                            break
                        else:
                            print(f"Введите число от 1 до {len(documents)}")
                            
                    except ValueError:
                        print("Введите корректный номер документа")
            else:
                # Только один документ
                self.acad_doc = documents[0][0]
            
            print("✅ Подключение к AutoCAD успешно")
            print(f"📄 Выбранный документ: {self.acad_doc.Name}")
            return True
            
        except ImportError:
            print("❌ Не удается импортировать модуль win32com. Установите: pip install pywin32")
            return False
        except RuntimeError as e:
            print(f"❌ Ошибка подключения к AutoCAD: {e}")
            print("Убедитесь, что AutoCAD запущен и имеет открытый документ")
            return False
        except Exception as e:
            print(f"❌ Неожиданная ошибка подключения к AutoCAD: {e}")
            return False
    
    def find_room_blocks(self) -> List[Tuple]:
        """
        Находит все блоки помещений в чертеже
        
        Returns:
            Список кортежей (блок, номер_помещения)
        """
        if not self.acad_doc:
            print("Нет подключения к AutoCAD")
            return []
        
        room_blocks = []
        
        try:
            # Перебираем все объекты в пространстве модели
            for entity in self.acad_doc.ModelSpace:
                # Проверяем, является ли объект вставкой блока
                if entity.ObjectName == "AcDbBlockReference":
                    block_name = entity.EffectiveName
                    
                    # Проверяем, является ли это блоком помещения
                    if any(target in block_name.upper() for target in self.target_blocks):
                        room_number = self._get_room_number_from_block(entity)
                        if room_number:
                            room_blocks.append(entity)  # Возвращаем только entity, без кортежа
                            print(f"Найден блок помещения {room_number}: {block_name}")
            
            print(f"Всего найдено блоков помещений: {len(room_blocks)}")
            return room_blocks
            
        except AttributeError as e:
            print(f"Ошибка доступа к объектам AutoCAD: {e}")
            return []
        except RuntimeError as e:
            print(f"Ошибка выполнения при поиске блоков: {e}")
            return []
        except Exception as e:
            print(f"Неожиданная ошибка при поиске блоков: {e}")
            return []
    
    def _get_room_number_from_block(self, block_ref) -> Optional[str]:
        """
        Получает номер помещения из атрибутов блока
        
        Args:
            block_ref: Ссылка на блок AutoCAD
            
        Returns:
            Номер помещения или None
        """
        try:
            # Получаем атрибуты блока
            attributes = block_ref.GetAttributes()
            
            for room_attr in attributes:
                tag = room_attr.TagString.upper()
                
                # Ищем атрибут с номером помещения
                room_num_tag = self.block_attributes.get("room_number", "ROOM_NUM").upper()
                if tag == room_num_tag:
                    return room_attr.TextString.strip()
            
            return None
            
        except (AttributeError, IndexError) as e:
            print(f"Ошибка при получении номера помещения: {e}")
            return None
    
    def update_block_attributes(self, block_ref, room_data: Dict) -> bool:
        """
        Обновляет атрибуты блока данными помещения
        
        Args:
            block_ref: Ссылка на блок AutoCAD
            room_data: Данные помещения из Excel
            
        Returns:
            True если обновление успешно, False иначе
        """
        try:
            attributes = block_ref.GetAttributes()
            updated_count = 0
            
            for block_attr in attributes:
                tag = block_attr.TagString.upper()
                
                # Сопоставляем теги атрибутов с данными
                for data_field, attr_tag in self.block_attributes.items():
                    if tag == attr_tag.upper() and data_field in room_data:
                        value = room_data[data_field]
                        
                        # Специальная обработка для названия помещения
                        if data_field == "room_name":
                            new_value = self._abbreviate_room_name(str(value))
                        # Форматируем числовые значения с единицами измерения
                        if data_field in ["area", "air_supply", "air_extract", "heat_loss", "temperature"]:
                            try:
                                num_value = float(value)
                                if data_field == "area":
                                    new_value = f"{num_value:.1f} м²"
                                elif data_field in ["air_supply", "air_extract"]:
                                    new_value = f"{num_value:.0f} м³/ч"
                                elif data_field == "heat_loss":
                                    new_value = f"{num_value:.0f} Вт"
                                elif data_field == "temperature":
                                    new_value = f"{num_value:.1f}°C"
                                else:
                                    new_value = f"{num_value:.0f}"
                            except (ValueError, TypeError):
                                new_value = "0"
                        else:
                            new_value = str(value)
                        
                        # Обновляем атрибут
                        if block_attr.TextString != new_value:
                            block_attr.TextString = new_value
                            updated_count += 1
                            print(f"  Обновлен {tag}: {new_value}")
            
            if updated_count > 0:
                # Обновляем блок в чертеже
                block_ref.Update()
                return True
            else:
                print("  Нет изменений для блока")
                return False
                
        except (AttributeError, ValueError) as e:
            print(f"Ошибка при обновлении атрибутов блока: {e}")
            return False
    
    def update_all_room_blocks(self, room_data_list: List[Dict]) -> Dict[str, int]:
        """
        Обновляет все блоки помещений данными из Excel
        
        Args:
            room_data_list: Список данных помещений
            
        Returns:
            Статистика обновлений
        """
        if not self.acad_doc:
            print("Нет подключения к AutoCAD")
            return {"updated": 0, "not_found": 0, "errors": 0}
        
        # Создаем словарь для быстрого поиска данных по номеру помещения
        room_data_dict = {
            str(room["room_number"]): room 
            for room in room_data_list 
            if room.get("room_number")
        }
        
        # Находим все блоки помещений
        room_blocks = self.find_room_blocks()
        
        stats = {"updated": 0, "not_found": 0, "errors": 0}
        
        for block_ref, room_number in room_blocks:
            try:
                if room_number in room_data_dict:
                    room_data = room_data_dict[room_number]
                    print(f"Обновление помещения {room_number}: {room_data.get('room_name', '')}")
                    
                    if self.update_block_attributes(block_ref, room_data):
                        stats["updated"] += 1
                    else:
                        print("  Не удалось обновить блок")
                        stats["errors"] += 1
                else:
                    print(f"Данные для помещения {room_number} не найдены в Excel")
                    stats["not_found"] += 1
                    
            except Exception as e:
                print(f"Ошибка при обработке блока {room_number}: {e}")
                stats["errors"] += 1
        
        # Сохраняем документ
        try:
            self.acad_doc.Save()
            print("Документ сохранен")
        except (AttributeError, PermissionError) as e:
            print(f"Ошибка при сохранении документа: {e}")
        
        return stats
    
    def read_block_attributes(self, room_blocks: List) -> List[Dict]:
        """
        Читает атрибуты из существующих блоков помещений
        
        Args:
            room_blocks: Список блоков помещений
            
        Returns:
            Список словарей с данными атрибутов блоков
        """
        if not self.acad_doc:
            print("❌ Нет подключения к AutoCAD")
            return []
        
        block_data = []
        
        print(f"\n📖 Чтение атрибутов из {len(room_blocks)} блоков...")
        
        for i, block_ref in enumerate(room_blocks):
            try:
                room_info = {}
                
                # Читаем атрибуты блока
                if hasattr(block_ref, 'HasAttributes') and block_ref.HasAttributes:
                    block_attrs = block_ref.GetAttributes()
                    
                    for block_attr in block_attrs:
                        attr_tag = block_attr.TagString.upper()
                        attr_value = block_attr.TextString.strip()
                        
                        # Сопоставляем атрибуты блока с полями конфигурации
                        for field, block_tag in self.config.get("block_attributes", {}).items():
                            if attr_tag == block_tag.upper():
                                room_info[field] = attr_value
                                break
                
                # Добавляем информацию о блоке если есть номер помещения
                if room_info.get('room_number'):
                    room_info['block_index'] = i
                    block_data.append(room_info)
                    print(f"  📋 Блок {room_info.get('room_number')}: прочитано {len(room_info)-1} атрибутов")
                else:
                    print(f"  ⚠️  Блок #{i+1}: номер помещения не найден")
                    
            except Exception as e:
                print(f"  ❌ Ошибка чтения блока #{i+1}: {e}")
                continue
        
        print(f"✅ Успешно прочитано {len(block_data)} блоков")
        return block_data
    
    def create_blocks_from_template(self, room_data_list: List[Dict], template_block_name: str, 
                                   start_point: Tuple[float, float] = None, 
                                   spacing: Tuple[float, float] = None) -> Dict[str, int]:
        """
        Создает блоки помещений из шаблона на основе данных Excel
        
        Args:
            room_data_list: Список данных помещений
            template_block_name: Имя блока-шаблона
            start_point: Начальная точка размещения (x, y). Если None - из конфига
            spacing: Расстояние между блоками (dx, dy). Если None - из конфига
            
        Returns:
            Статистика создания блоков
        """
        
        # Получаем настройки из конфига, если параметры не заданы
        grid_settings = self.config.get("placement_settings", {}).get("grid", {})
        
        if start_point is None:
            start_point = (
                grid_settings.get("default_start_x", 0),
                grid_settings.get("default_start_y", 0)
            )
            
        if spacing is None:
            spacing = (
                grid_settings.get("default_spacing_x", 100),
                grid_settings.get("default_spacing_y", 50)
            )
            
        blocks_per_row = grid_settings.get("blocks_per_row", 10)
        if not self.acad_doc:
            print("❌ Нет подключения к AutoCAD")
            return {"created": 0, "errors": 0}
        
        # Проверяем существование шаблона блока
        if not self.block_exists(template_block_name):
            print(f"❌ Блок-шаблон '{template_block_name}' не найден в чертеже")
            return {"created": 0, "errors": 0}
        
        stats = {"created": 0, "errors": 0}
        
        print(f"\n🏗️  Создание блоков из шаблона '{template_block_name}'...")
        print(f"   📍 Начальная точка: {start_point}")
        print(f"   📏 Расстояние между блоками: {spacing}")
        print(f"   📋 Блоков в ряду: {blocks_per_row}")
        
        # Размещаем блоки по сетке
        for i, room_data in enumerate(room_data_list):
            try:
                # Вычисляем координаты для текущего блока
                row = i // blocks_per_row  # Количество блоков в строке из конфига
                col = i % blocks_per_row
                
                x = start_point[0] + col * spacing[0]
                y = start_point[1] - row * spacing[1]  # Вниз по Y
                
                print(f"  Создание блока {room_data.get('room_number', 'N/A')} в позиции ({x}, {y})...")
                
                # Проверяем пространство модели
                if not hasattr(self.acad_doc, 'ModelSpace'):
                    raise AttributeError("Нет доступа к пространству модели")
                
                # Создаем точку вставки как массив координат (важно: double тип для COM)
                insert_point = [float(x), float(y), 0.0]
                
                # Пробуем несколько методов вставки блока
                block_ref = None
                
                # Метод 1: InsertBlock с VARIANT (надежный)
                try:
                    # Используем числовые константы для избежания проблем с импортом
                    # VT_ARRAY = 8192, VT_R8 = 5 (double array)
                    variant_point = win32com.client.VARIANT(
                        8192 | 5,  # VT_ARRAY | VT_R8
                        insert_point
                    )
                    
                    block_ref = self.acad_doc.ModelSpace.InsertBlock(
                        variant_point,
                        str(template_block_name),
                        1.0, 1.0, 1.0, 0.0
                    )
                        
                except Exception:
                    # Метод 2: Обычный InsertBlock (fallback)
                    try:
                        block_ref = self.acad_doc.ModelSpace.InsertBlock(
                            insert_point,
                            str(template_block_name),
                            1.0, 1.0, 1.0, 0.0
                        )
                            
                    except Exception as e2:
                        print("    ❌ Не удалось вставить блок")
                        raise RuntimeError(f"Не удалось вставить блок {template_block_name}") from e2
                
                if block_ref is None:
                    raise RuntimeError("Блок не был создан ни одним из методов")
                
                # Заполняем атрибуты нового блока
                if self._fill_block_attributes(block_ref, room_data):
                    stats["created"] += 1
                    print(f"  ✅ Создан блок для помещения {room_data.get('room_number', 'N/A')}")
                else:
                    stats["errors"] += 1
                    print(f"  ⚠️  Блок создан, но ошибка заполнения атрибутов для {room_data.get('room_number', 'N/A')}")
                
            except Exception as e:
                stats["errors"] += 1
                error_msg = str(e)
                if "2147352567" in error_msg:
                    print(f"  ❌ COM ошибка для помещения {room_data.get('room_number', 'N/A')}: Возможно, блок '{template_block_name}' не существует или недоступен")
                    print(f"     Проверьте, что блок определен в текущем чертеже")
                else:
                    print(f"  ❌ Ошибка создания блока для помещения {room_data.get('room_number', 'N/A')}: {error_msg}")
        
        print("\n📊 Создание блоков завершено:")
        print(f"   ✅ Создано: {stats['created']}")
        print(f"   ❌ Ошибок: {stats['errors']}")
        
        return stats
    
    def create_blocks_from_coordinates(self, room_data_list: List[Dict], template_block_name: str,
                                     coordinates_list: List[Tuple[float, float]]) -> Dict[str, int]:
        """
        Создает блоки помещений в заданных координатах
        
        Args:
            room_data_list: Список данных помещений
            template_block_name: Имя блока-шаблона
            coordinates_list: Список координат (x, y) для размещения блоков
            
        Returns:
            Статистика создания блоков
        """
        if not self.acad_doc:
            print("❌ Нет подключения к AutoCAD")
            return {"created": 0, "errors": 0}
        
        if len(coordinates_list) < len(room_data_list):
            print(f"⚠️  Координат ({len(coordinates_list)}) меньше чем помещений ({len(room_data_list)})")
            print("   Будут созданы блоки только для помещений с координатами")
        
        # Проверяем существование шаблона блока
        if not self.block_exists(template_block_name):
            print(f"❌ Блок-шаблон '{template_block_name}' не найден в чертеже")
            return {"created": 0, "errors": 0}
        
        stats = {"created": 0, "errors": 0}
        
        print("\n🎯 Создание блоков в заданных координатах...")
        
        # Создаем блоки по заданным координатам
        for i, room_data in enumerate(room_data_list[:len(coordinates_list)]):
            try:
                x, y = coordinates_list[i]
                
                # Создаем точку вставки как массив координат (важно: double тип для COM)
                insert_point = [float(x), float(y), 0.0]
                
                # Пробуем несколько методов вставки блока
                block_ref = None
                
                # Метод 1: InsertBlock с VARIANT (надежный)
                try:
                    # Используем числовые константы для избежания проблем с импортом
                    # VT_ARRAY = 8192, VT_R8 = 5 (double array)
                    variant_point = win32com.client.VARIANT(
                        8192 | 5,  # VT_ARRAY | VT_R8
                        insert_point
                    )
                    
                    block_ref = self.acad_doc.ModelSpace.InsertBlock(
                        variant_point,
                        str(template_block_name),
                        1.0, 1.0, 1.0, 0.0
                    )
                        
                except Exception:
                    # Метод 2: Обычный InsertBlock (fallback)
                    block_ref = self.acad_doc.ModelSpace.InsertBlock(
                        insert_point,
                        str(template_block_name),
                        1.0, 1.0, 1.0, 0.0
                    )
                
                if block_ref is None:
                    raise RuntimeError("Блок не был создан ни одним из методов")
                
                # Заполняем атрибуты
                if self._fill_block_attributes(block_ref, room_data):
                    stats["created"] += 1
                    print(f"  ✅ Создан блок для помещения {room_data.get('room_number')} в ({x}, {y})")
                else:
                    stats["errors"] += 1
                    
            except (AttributeError, ValueError, IndexError) as e:
                stats["errors"] += 1
                print(f"  ❌ Ошибка создания блока для {room_data.get('room_number')}: {e}")
        
        return stats
    
    def block_exists(self, block_name: str) -> bool:
        """Проверяет существование блока в чертеже"""
        try:
            doc_blocks = self.acad_doc.Blocks
            for i in range(doc_blocks.Count):
                current_block = doc_blocks.Item(i)
                current_name = current_block.Name
                
                if current_name.upper() == block_name.upper():
                    return True
            return False
        except (AttributeError, IndexError):
            return False
    
    def _abbreviate_room_name(self, room_name: str) -> str:
        """
        Сокращает длинные слова в названии помещения
        
        Args:
            room_name: Исходное название помещения
            
        Returns:
            Название с сокращениями
        """
        if not room_name:
            return ""
            
        # Словарь сокращений
        abbreviations = {
            "Помещение": "Пом.",
            "помещение": "пом.",
            "Кабинет": "Каб.",
            "кабинет": "каб.",
            "Лаборатория": "Лаб.",
            "лаборатория": "лаб.",
            "Производственное": "Произв.",
            "производственное": "произв.",
            "Техническое": "Тех.",
            "техническое": "тех."
        }
        
        result = str(room_name)
        
        # Применяем сокращения
        for full_word, abbreviated in abbreviations.items():
            result = result.replace(full_word, abbreviated)
        
        return result

    def _fill_block_attributes(self, block_ref, room_data: Dict) -> bool:
        """
        Заполняет атрибуты блока данными помещения
        
        Args:
            block_ref: Ссылка на вставленный блок
            room_data: Данные помещения
            
        Returns:
            True если успешно, False иначе
        """
        try:
            # Проверяем наличие атрибутов у блока
            if not hasattr(block_ref, 'GetAttributes'):
                return False
                
            attributes = block_ref.GetAttributes()
            
            for fill_attr in attributes:
                tag = fill_attr.TagString.upper()
                
                # Сопоставляем теги атрибутов с данными
                for data_field, attr_tag in self.block_attributes.items():
                    if tag == attr_tag.upper() and data_field in room_data:
                        value = room_data[data_field]
                        
                        # Специальная обработка для названия помещения
                        if data_field == "room_name":
                            new_value = self._abbreviate_room_name(str(value))
                        # Форматируем числовые значения с единицами измерения
                        if data_field in ["area", "air_supply", "air_extract", "heat_loss", "temperature"]:
                            try:
                                num_value = float(value)
                                if data_field == "area":
                                    new_value = f"{num_value:.1f} м²"
                                elif data_field in ["air_supply", "air_extract"]:
                                    new_value = f"{num_value:.0f} м³/ч"
                                elif data_field == "heat_loss":
                                    new_value = f"{num_value:.0f} Вт"
                                elif data_field == "temperature":
                                    new_value = f"{num_value:.1f}°C"
                                else:
                                    new_value = f"{num_value:.0f}"
                            except (ValueError, TypeError):
                                new_value = "0"
                        else:
                            new_value = str(value)
                        
                        fill_attr.TextString = new_value
            
            # Обновляем блок
            block_ref.Update()
            return True
            
        except (AttributeError, ValueError) as e:
            print(f"    ❌ Ошибка заполнения атрибутов: {e}")
            return False
    
    def get_available_blocks(self) -> List[str]:
        """
        Получает список всех блоков в чертеже
        
        Returns:
            Список имен блоков
        """
        if not self.acad_doc:
            return []
        
        block_names = []
        try:
            doc_blocks = self.acad_doc.Blocks
            for i in range(doc_blocks.Count):
                doc_block = doc_blocks.Item(i)
                # Исключаем системные блоки
                if not doc_block.Name.startswith('*'):
                    block_names.append(doc_block.Name)
            
            return sorted(block_names)
            
        except AttributeError as e:
            print(f"❌ Ошибка получения списка блоков: {e}")
            return []
    
    def get_block_info(self) -> List[Dict]:
        """
        Получает информацию о всех блоках в чертеже для анализа
        
        Returns:
            Список словарей с информацией о блоках
        """
        if not self.acad_doc:
            return []
        
        blocks_info = []
        
        try:
            for entity in self.acad_doc.ModelSpace:
                if entity.ObjectName == "AcDbBlockReference":
                    block_info = {
                        "name": entity.EffectiveName,
                        "attributes": []
                    }
                    
                    try:
                        attributes = entity.GetAttributes()
                        for block_attr in attributes:
                            block_info["attributes"].append({
                                "tag": block_attr.TagString,
                                "value": block_attr.TextString
                            })
                    except AttributeError:
                        # Блок без атрибутов
                        pass
                    
                    blocks_info.append(block_info)
            
            return blocks_info
            
        except AttributeError as e:
            print(f"Ошибка при получении информации о блоках: {e}")
            return []


# Пример использования
if __name__ == "__main__":
    controller = AutoCADController()
    
    if controller.connect_to_autocad():
        # Получаем информацию о блоках
        blocks = controller.get_block_info()
        print(f"Найдено блоков в чертеже: {len(blocks)}")
        
        # Показываем первые несколько блоков
        for block in blocks[:5]:
            print(f"Блок: {block['name']}")
            for attr in block['attributes']:
                print(f"  {attr['tag']}: {attr['value']}")