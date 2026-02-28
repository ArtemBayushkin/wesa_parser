
import json
from dwg_parser import AutoCADProcessor
from multiprocessing import Pool

if __name__ == '__main__':
    config_path = '.\config.json'
    with open(config_path, 'r', encoding='utf-8') as f:
                config_data = json.load(f)
    file_rename_rules = config_data.get('АЭС Эль-Дабаа Блоки 1 и 2', {}).get("file_rename", {})
    print(file_rename_rules)
    input_path = "D:\Новая папка\ED.D.P000.2.0UKD&&JEW&&&.021.DC.0001.E-MLV0001"
    input_path2 = "D:\Новая папка\ED.D.P000.2.0UKD&&JEW&&&.021.DC.0001.E-MLV0002"
    output_path = "D:\Новая папка\ED.D.P000.1.0UKD&&JEW&&&.021.DC.0001.E-MLV0001"
    output_path2 = "D:\Новая папка\ED.D.P000.1.0UKD&&JEW&&&.021.DC.0001.E-MLV0002"
    inp = (input_path, input_path2)
    out = (output_path, output_path2)

    processor = AutoCADProcessor('1', 'El Dabaa', file_rename_rules)
    processor.process_file(input_path, output_path)
    print('Heelo')