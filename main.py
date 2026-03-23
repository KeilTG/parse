import pandas as pd
import os
import sys
import argparse
from pathlib import Path

def split_excel_by_group(input_file, group_column, output_dir=None, output_format='xlsx'):
    """
    Разделяет Excel файл на несколько по уникальным значениям в указанном столбце
    
    Args:
        input_file: путь к исходному Excel файлу
        group_column: название столбца для группировки
        output_dir: папка для сохранения результатов (по умолчанию папка с исходным файлом)
        output_format: формат выходных файлов ('xlsx' или 'csv')
    """
    
    # Проверяем существование файла
    if not os.path.exists(input_file):
        print(f"❌ Ошибка: Файл '{input_file}' не найден!")
        return False
    
    try:
        # Определяем папку для сохранения
        input_path = Path(input_file)
        if output_dir is None:
            output_dir = input_path.parent / f"{input_path.stem}_split"
        else:
            output_dir = Path(output_dir)
        
        # Создаем папку для результатов
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Читаем Excel файл
        print(f"📖 Чтение файла: {input_file}")
        df = pd.read_excel(input_file)
        
        # Проверяем существование столбца
        if group_column not in df.columns:
            print(f"❌ Ошибка: Столбец '{group_column}' не найден в файле!")
            print(f"📋 Доступные столбцы: {', '.join(df.columns)}")
            return False
        
        # Получаем уникальные значения для группировки
        unique_groups = df[group_column].unique()
        print(f"📊 Найдено {len(unique_groups)} уникальных групп в столбце '{group_column}'")
        
        # Создаем файлы для каждой группы
        created_files = []
        for group in unique_groups:
            # Фильтруем данные для текущей группы
            group_df = df[df[group_column] == group]
            
            # Формируем имя файла (заменяем недопустимые символы)
            safe_name = str(group).replace('/', '_').replace('\\', '_').replace(':', '_')
            if output_format == 'xlsx':
                output_file = output_dir / f"{safe_name}.xlsx"
                group_df.to_excel(output_file, index=False)
            else:
                output_file = output_dir / f"{safe_name}.csv"
                group_df.to_csv(output_file, index=False, encoding='utf-8-sig')
            
            created_files.append(output_file)
            print(f"   ✅ Создан: {output_file.name} ({len(group_df)} строк)")
        
        print(f"\n✨ Готово! Создано {len(created_files)} файлов в папке:")
        print(f"   {output_dir}")
        return True
        
    except Exception as e:
        print(f"❌ Ошибка при обработке: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(
        description='Разделение Excel файла на несколько по группам',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Примеры использования:
  python split_excel.py data.xlsx -c "Отдел"
  python split_excel.py data.xlsx -c "Город" -o ./output
  python split_excel.py data.xlsx -c "Категория" -f csv
        """
    )
    
    parser.add_argument('input_file', help='Путь к исходному Excel файлу')
    parser.add_argument('-c', '--column', required=True, help='Название столбца для группировки')
    parser.add_argument('-o', '--output', help='Папка для сохранения результатов')
    parser.add_argument('-f', '--format', choices=['xlsx', 'csv'], default='xlsx', 
                       help='Формат выходных файлов (по умолчанию: xlsx)')
    
    args = parser.parse_args()
    
    print("=" * 50)
    print("🔄 Разделитель Excel файлов по группам")
    print("=" * 50)
    
    success = split_excel_by_group(
        input_file=args.input_file,
        group_column=args.column,
        output_dir=args.output,
        output_format=args.format
    )
    
    sys.exit(0 if success else 1)

if __name__ == "__main__":
    main()