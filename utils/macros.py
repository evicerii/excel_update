'''
macros copy def
'''
import os
import tempfile
import win32com.client
def copy_macros(source_file, target_file):
    '''
    копируает макрос между указаными файлами
    '''
    # Запускаем Excel
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible = False
    source_wb = None
    target_wb = None
    try:
        # Проверяем существование файлов
        if not os.path.exists(source_file):
            raise FileNotFoundError(f"Исходный файл не найден: {source_file}")
        if not os.path.exists(target_file):
            raise FileNotFoundError(f"Целевой файл не найден: {target_file}")
        # Открываем исходный файл
        source_wb = excel.Workbooks.Open(source_file)
        # Открываем целевой файл
        target_wb = excel.Workbooks.Open(target_file)
        # Получаем доступ к модулям VBA исходного файла
        vba_components = source_wb.VBProject.VBComponents
        # Копируем каждый модуль
        for component in vba_components:
            # Экспортируем модуль во временный файл
            temp_file = os.path.join(tempfile.gettempdir(), f"{component.Name}.bas")
            component.Export(temp_file)
            # Импортируем модуль в целевой файл
            target_wb.VBProject.VBComponents.Import(temp_file)
            # Удаляем временный файл
            os.remove(temp_file)
        # Сохраняем и закрываем целевой файл
        target_wb.Save()
    except Exception as e:
        print(f"Произошла ошибка: {e}")
    finally:
        # Закрываем рабочие книги и Excel, если они были открыты
        if source_wb is not None:
            source_wb.Close(SaveChanges=False)
        if target_wb is not None:
            target_wb.Close(SaveChanges=True)
        excel.Quit()