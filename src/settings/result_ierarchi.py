# Настройки

target_columns = ['Подразделение', 'Структурное подразделение', 'Должность, структурное подразделение']
# folder_path = "/home/wol/programs/scripts/matrix_applications/данные/НИЛ2026/выход"
# folder_path = "/home/wol/programs/scripts/matrix_applications/данные/НИЛ2026/Выход2_30.03.2026"
# folder_path = "/home/wol/programs/scripts/matrix_applications/данные/НИЛ2026/тест"
folder_path = "/home/wol/programs/scripts/matrix_applications/данные/НИЛ2026/выход4_1.04.2026"

def normalize_department_name(department):
    """
    Нормализует название подразделения:
    - все строки, начинающиеся с 'Администрация', преобразуются в 'Администрация';
    - остальные названия остаются без изменений.
    """
    if isinstance(department, str) and department.strip().startswith('Администрация'):
        return 'Администрация'

    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Вуктыльское ордена Трудового Красного Знамени газопромысловое управление'):
        return 'ВГПУ'

    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Инженерно-технический центр'):
        return 'ИТЦ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Линейное производственное управление межпромысловых трубопроводов'):
        return 'ЛПУМТ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Управление аварийно- восстановительных работ'):
        return 'УАВР'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Служба корпоративной защиты'):
        return 'СКЗ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Каневское ордена Трудового Красного Знамени газопромысловое управление'):
        return 'КГПУ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Светлоградское'):
        return 'СГППУ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Управление технологического транспорта и спецтехники'):
        return 'УТТиСТ'
    
    if isinstance(department, str) and department.strip().startswith('Филиал ООО "Газпром добыча Краснодар" - Управление материально- технического снабжения и комплектации'):
        return 'УМТСиК'
      
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Медицинская служба при администрации'):
        return 'СППА Медицинская служба'
     
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Служба главного маркшейдера при администрации'):
        return 'СППА Служба главного маркшейдера'
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Служба организации восстановления основных фондов при администрации'):
        return 'СППА Служба организации восстановления основных фондов'
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Служба информационно-управляющих систем при администрации'):
        return 'СППА СИУС'
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Служба организации реконструкции и строительства основных фондов при администрации'):
        return 'СППА СОРиСОФ'
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Учебно-производственный центр при администрации'):
        return 'СППА УПЦ'
    
    if isinstance(department, str) and department.strip().startswith('Структурные подразделения при Администрации / Служба по эксплуатации зданий и сооружений при администрации'):
        return 'СППА СПЭЗиС'
    
    return department
